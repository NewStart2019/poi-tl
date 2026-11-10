package com.deepoove.poi.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.IdentityHashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;

/**
 * Bean 转 Map 工具。
 *
 * <p>占位/取值规则：
 * <ul>
 *     <li>字符串、原始类型包装类、{@link Number}（含 BigDecimal、原子类）等标量原样保留；</li>
 *     <li>{@link Date}、{@link LocalDateTime} 统一格式化为 {@code yyyy-MM-dd HH:mm:ss}，
 *     {@link LocalDate} 格式化为 {@code yyyy-MM-dd}；</li>
 *     <li>枚举、{@link Class}、{@link Method} 原样保留；</li>
 *     <li>集合、对象数组转换为 List，其中每个元素独立转换，{@code null} 元素保留为 {@code null}；</li>
 *     <li>原始类型数组（如 {@code int[]}）原样保留；</li>
 *     <li>命中 {@code noTransferClass} 的对象原样保留，不再向下转换；</li>
 *     <li>最大递归深度为 {@link #MAX_DEPTH} 层，到达上限后容器对象原样保留，避免数据丢失与栈溢出；</li>
 *     <li>循环引用按对象身份识别，复用正在构建的同一 Map，避免死循环；</li>
 *     <li>字段解析会沿父类链向上收集（不含 {@code Object}），跳过 {@code static} 与匿名类的
 *     {@code this$0} 外部类引用。</li>
 * </ul>
 *
 * <p><b>线程安全说明：</b>本类不再保存跨调用的可变状态，日期统一使用不可变的
 * {@link DateTimeFormatter}，因此可在多线程下共享同一个实例。每次顶层调用都会创建独立的
 * 循环引用上下文，转换结束即释放，不会返回陈旧数据，也不会产生内存泄漏。
 */
public class TlBeanUtil {

    private static final Logger log = LoggerFactory.getLogger(TlBeanUtil.class);

    /** 最大递归深度（层数）。 */
    private static final int MAX_DEPTH = 100;

    /** 日期时间格式（线程安全）。 */
    private static final DateTimeFormatter DATE_TIME_FORMATTER =
        DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    /** 日期格式（线程安全）。 */
    private static final DateTimeFormatter DATE_FORMATTER =
        DateTimeFormatter.ofPattern("yyyy-MM-dd");
    /** 旧版 {@link Date} 转 java.time 时使用的系统时区。 */
    private static final ZoneId SYSTEM_ZONE = ZoneId.systemDefault();

    public <T> Map<String, Object> beanToMap(Object obj, Class<T> noTransferClass) {
        return this.beanToMap(obj, noTransferClass, 0);
    }

    /**
     * Bean 转 Map，指定不做转换而原样保留的类型。
     *
     * @param obj             原始对象
     * @param noTransferClass 需要原样保留的类型，不能为 {@code null}
     * @param depth           起始深度
     * @return 转换结果；顶层为 {@code null} 或非 Bean 的简单类型时返回空 Map
     */
    public <T> Map<String, Object> beanToMap(Object obj, Class<T> noTransferClass, int depth) {
        if (noTransferClass == null) {
            throw new IllegalArgumentException("noTransferClass must not be null");
        }
        if (obj == null) {
            return new HashMap<>();
        }
        // 顶层命中保留类：受返回类型（Map）限制，包一层 {简单类名: 对象}，与历史行为保持一致。
        if (noTransferClass.isInstance(obj)) {
            Map<String, Object> wrapper = new HashMap<>();
            wrapper.put(obj.getClass().getSimpleName(), obj);
            return wrapper;
        }

        // 每次顶层调用使用独立的循环引用上下文。
        ConvertContext context = new ConvertContext();
        Object converted = this.convertValue(obj, noTransferClass, depth, context);
        if (converted instanceof Map) {
            @SuppressWarnings("unchecked")
            Map<String, Object> result = (Map<String, Object>) converted;
            return result;
        }
        // 顶层为标量等非 Bean 类型：返回空 Map（与历史行为一致），并给出明确提示。
        log.warn("Top-level object of type {} is not a bean, return an empty map",
            obj.getClass().getName());
        return new HashMap<>();
    }

    /**
     * 统一的值转换分发器：字段、集合、数组、Map 中的元素都走这一套逻辑，保证口径一致。
     *
     * @param value 待转换的值
     * @param noTransferClass 保留类型
     * @param depth 当前深度
     * @param context 本次转换的循环引用上下文
     * @return 转换后的值（可能是 Map、List、字符串或原对象）
     */
    private <T> Object convertValue(Object value, Class<T> noTransferClass, int depth,
                                    ConvertContext context) {
        if (value == null) {
            return null;
        }
        // 1. 命中保留类型：原样返回。
        if (noTransferClass.isInstance(value)) {
            return value;
        }
        // 2. 标量：字符串、字符、所有 Number（含包装类、BigDecimal、原子类）。
        if (value instanceof String || isScalar(value)) {
            return value;
        }
        // 3. 日期/时间：统一格式化为字符串。
        if (value instanceof LocalDateTime) {
            return ((LocalDateTime) value).format(DATE_TIME_FORMATTER);
        }
        if (value instanceof LocalDate) {
            return ((LocalDate) value).format(DATE_FORMATTER);
        }
        if (value instanceof Date) {
            return formatLegacyDate((Date) value);
        }
        // 4. 枚举、Class、Method：原样保留。
        if (value instanceof Enum || value instanceof Class || value instanceof Method) {
            return value;
        }
        // 5. 原始类型数组：原样保留。
        if (isPrimitiveArray(value)) {
            return value;
        }
        // 6. 深度兜底：容器/复杂对象到达最大深度后原样保留，避免数据丢失与栈溢出。
        if (depth >= MAX_DEPTH) {
            return value;
        }
        // 7. 容器类型。
        if (value.getClass().isArray()) {
            return convertObjectArray((Object[]) value, noTransferClass, depth, context);
        }
        if (value instanceof Collection) {
            return convertCollection((Collection<?>) value, noTransferClass, depth, context);
        }
        if (value instanceof Map) {
            return convertMap((Map<?, ?>) value, noTransferClass, depth, context);
        }
        // 8. 普通复杂 Bean。
        return convertBean(value, noTransferClass, depth, context);
    }

    /**
     * 转换复杂 Bean：处理循环引用并逐字段转换。
     */
    private <T> Map<String, Object> convertBean(Object bean, Class<T> noTransferClass, int depth,
                                                ConvertContext context) {
        // 循环引用：直接复用正在构建的同一 Map（身份匹配）。
        Map<String, Object> existing = context.cycles.get(bean);
        if (existing != null) {
            return existing;
        }
        Map<String, Object> result = new HashMap<>();
        // 先占位，保证自引用/相互引用能命中同一 Map。
        context.cycles.put(bean, result);

        for (Field field : resolveFields(bean.getClass())) {
            try {
                field.setAccessible(true);
                result.put(field.getName(),
                    convertValue(field.get(bean), noTransferClass, depth + 1, context));
            } catch (Exception e) {
                log.error("Read field [{}] of [{}] failed: {}",
                    field.getName(), bean.getClass().getName(), e.getMessage());
            }
        }
        return result;
    }

    /**
     * 集合转换：逐元素转换，{@code null} 保留为 {@code null}，日期/枚举等与字段口径一致。
     */
    private <T> List<Object> convertCollection(Collection<?> collection, Class<T> noTransferClass,
                                               int depth, ConvertContext context) {
        List<Object> list = new ArrayList<>(collection.size());
        for (Object item : collection) {
            list.add(convertValue(item, noTransferClass, depth + 1, context));
        }
        return list;
    }

    /**
     * 对象数组转换：逐元素独立转换，不再只依据首元素决定整体策略。
     */
    private <T> List<Object> convertObjectArray(Object[] array, Class<T> noTransferClass,
                                                int depth, ConvertContext context) {
        List<Object> list = new ArrayList<>(array.length);
        for (Object item : array) {
            list.add(convertValue(item, noTransferClass, depth + 1, context));
        }
        return list;
    }

    /**
     * Map 转换：安全处理 {@code null} 键与非字符串键，并对重名键去重，避免覆盖丢数据。
     */
    private <T> Map<String, Object> convertMap(Map<?, ?> map, Class<T> noTransferClass,
                                               int depth, ConvertContext context) {
        Map<String, Object> result = new HashMap<>();
        for (Map.Entry<?, ?> entry : map.entrySet()) {
            Object key = entry.getKey();
            String baseName;
            if (key == null) {
                baseName = "null";
            } else if (key instanceof String) {
                baseName = (String) key;
            } else {
                baseName = key.getClass().getSimpleName();
            }
            String uniqueName = uniqueKey(result, baseName);
            result.put(uniqueName,
                convertValue(entry.getValue(), noTransferClass, depth + 1, context));
        }
        return result;
    }

    /**
     * 解析类的全部实例字段：沿父类链向上收集（不含 {@code Object}），
     * 跳过静态字段与匿名类的 {@code this$} 外部类引用；子类同名字段优先。
     */
    private static List<Field> resolveFields(Class<?> clazz) {
        Map<String, Field> byName = new LinkedHashMap<>();
        Class<?> current = clazz;
        while (current != null && current != Object.class) {
            for (Field field : current.getDeclaredFields()) {
                int modifiers = field.getModifiers();
                if (Modifier.isStatic(modifiers)) {
                    continue;
                }
                if (field.isSynthetic() && field.getName().contains("this$")) {
                    continue;
                }
                // 子类字段先加入，父类同名字段不覆盖。
                byName.putIfAbsent(field.getName(), field);
            }
            current = current.getSuperclass();
        }
        return new ArrayList<>(byName.values());
    }

    /**
     * 判断是否为不可再分的标量：原始包装类、字符或任意 {@link Number}。
     */
    private static boolean isScalar(Object value) {
        return isPrimitive(value) || value instanceof Number;
    }

    /**
     * 旧版 {@link Date} 的线程安全格式化：转 java.time 后使用不可变 formatter。
     */
    private static String formatLegacyDate(Date date) {
        return date.toInstant()
            .atZone(SYSTEM_ZONE)
            .toLocalDateTime()
            .format(DATE_TIME_FORMATTER);
    }

    /**
     * 当目标 Map 已存在同名键时，追加 {@code #序号} 保证唯一。
     */
    private static String uniqueKey(Map<String, ?> map, String baseName) {
        if (!map.containsKey(baseName)) {
            return baseName;
        }
        int index = 1;
        String candidate;
        do {
            candidate = baseName + "#" + index++;
        } while (map.containsKey(candidate));
        return candidate;
    }

    /**
     * 判断对象是否为原始类型的包装类。
     */
    public static boolean isPrimitive(Object obj) {
        if (obj == null) {
            return false;
        }
        Class<?> clazz = obj.getClass();
        return clazz.equals(Integer.class) || clazz.equals(Character.class) ||
            clazz.equals(Boolean.class) || clazz.equals(Byte.class) ||
            clazz.equals(Short.class) || clazz.equals(Double.class) ||
            clazz.equals(Float.class) || clazz.equals(Long.class);
    }

    /**
     * 判断类是否为原始类型或其包装类。
     */
    public static boolean isPrimitive(Class<?> clazz) {
        if (clazz == null) {
            return false;
        }
        return clazz.isPrimitive() || clazz.equals(Integer.class) || clazz.equals(Character.class) ||
            clazz.equals(Boolean.class) || clazz.equals(Byte.class) ||
            clazz.equals(Short.class) || clazz.equals(Double.class) ||
            clazz.equals(Float.class) || clazz.equals(Long.class);
    }

    /**
     * 判断对象是否为原始类型数组（元素为原始类型，而非对象）。
     */
    public static boolean isPrimitiveArray(Object obj) {
        if (obj == null || !obj.getClass().isArray()) {
            return false;
        }
        return isPrimitive(obj.getClass().getComponentType());
    }

    /**
     * 判断是否为常见的数字引用类型。
     */
    public static boolean isNumber(Object obj) {
        return obj instanceof AtomicInteger || obj instanceof AtomicLong
            || obj instanceof BigDecimal;
    }

    /**
     * 单次转换的上下文：按对象身份记录正在/已经转换的 Bean，用于打破循环引用。
     * 每次顶层调用独立创建，转换结束即可被回收。
     */
    private static final class ConvertContext {
        private final IdentityHashMap<Object, Map<String, Object>> cycles =
            new IdentityHashMap<>();
    }
}
