package com.deepoove.poi.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

public class TlBeanUtil {
    private static final Logger log = LoggerFactory.getLogger(TlBeanUtil.class);
    private final ConcurrentHashMap<Object, Map<String, Object>> cache = new ConcurrentHashMap<>();
    private final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    private final DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public <T> Map<String, Object> beanToMap(Object obj, Class<T> noTransferClass) {
        return this.beanToMap(obj, noTransferClass, 0);
    }

    /**
     * <p>Bean to map, specify certain classes not to convert and preserve the original data. </p>
     * <p><red>Please read the following precautions: </red></p>
     * <li> Resolved recursive references to speed up parsing</li>
     * <li> If the input parameter is Map and its key is Object, the name of this object will be obtained,
     * and then the value will be converted into a Map to form a key value pair</li>
     * <li>If the input is an array or collection, the object does not support conversion</li>
     * <li>The maximum recursive depth is 10 layers</li>
     * <li>this$0-the reference property name of the external class of the anonymous class is not converted </li>
     * <p> Examples:
     * <blockquote><pre>
     *     TlBeanUtil beanUtil = new TlBeanUtil();
     *     Map<String, Object> map = beanUtil.beanToMap(map1, Address.class);
     * </pre></blockquote>
     *
     * @param obj             Original Object
     * @param noTransferClass Objects that need to be preserved
     * @return Map<String, Object>
     */
    public <T> Map<String, Object> beanToMap(Object obj, Class<T> noTransferClass, int depth) {
        Map<String, Object> map = new HashMap<>();
        int MAX_DEPTH = 10;
        if (obj == null || depth > MAX_DEPTH) {
            return map;
        }
        if (cache.containsKey(obj)) {
            return cache.get(obj);
        } else {
            cache.put(obj, map);
        }
        if (noTransferClass.isAssignableFrom(obj.getClass())) {
            map.put(obj.getClass().getSimpleName(), obj);
            return map;
        }
        try {
            if (obj instanceof Map) {
                ((Map) obj).forEach((k, v) -> {
                    try {
                        if (k instanceof String) {
                            this.dealProperty((String) k, v, noTransferClass, map, depth);
                        } else {
                            this.dealProperty(k.getClass().getSimpleName(), v, noTransferClass, map, depth);
                        }
                    } catch (IllegalAccessException e) {
                        throw new RuntimeException(e);
                    }
                });
            } else if (obj.getClass().isArray() || obj instanceof Collection || obj instanceof Date
                || obj instanceof LocalDate || obj instanceof Enum || obj instanceof Class) {
                log.error("Unsupported data type: " + obj.getClass());
                return map;
            } else {
                Class<?> clazz = obj.getClass();
                for (Field field : clazz.getDeclaredFields()) {
                    try {
                        field.setAccessible(true);
                        this.dealProperty(field.getName(), field.get(obj), noTransferClass, map, depth);
                    } catch (Exception exception) {
                        log.error("Error while accessing field: " + field.getName() + " in class: " + clazz.getName(), exception);
                    }
                }
            }
        } catch (Exception exception) {
            log.error("Error: ", exception);
        }
        return map;
    }

    private <T> void dealProperty(String fieldName, Object fieldValue, Class<T> noTransferClass, Map<String, Object> map, int depth)
        throws IllegalAccessException {
        // If the field value is a basic type or string
        if (fieldValue == null || fieldValue instanceof String || isPrimitive(fieldValue)) {
            map.put(fieldName, fieldValue);
        } else if (fieldValue instanceof Collection) {
            map.put(fieldName, collectionToMap((Collection<?>) fieldValue, noTransferClass, depth));
        } else if (fieldValue instanceof Date) {
            map.put(fieldName, sdf.format((Date) fieldValue));
        } else if (fieldValue instanceof LocalDateTime) {
            map.put(fieldName, dtf.format((LocalDateTime) fieldValue));
        } else if (fieldValue instanceof LocalDate) {
            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd");
            map.put(fieldName, dtf.format((LocalDate) fieldValue));
        } else if (fieldValue.getClass().isArray()) {
            if (isPrimitiveArray(fieldValue)) {
                map.put(fieldName, fieldValue);
            } else {
                map.put(fieldName, arrayToMap((Object[]) fieldValue, noTransferClass, depth));
            }
        } else {
            // If it is a filtering object， here is no need to cache it
            //  SpEL will put function objects without conversion
            if (noTransferClass.isAssignableFrom(fieldValue.getClass()) || fieldValue instanceof Method || fieldName.contains("this$")) {
                map.put(fieldName, fieldValue);
            } else {
                // If the field value is a complex object
                map.put(fieldName, beanToMap(fieldValue, noTransferClass, depth + 1));
            }
        }
    }

    public static boolean isPrimitive(Object obj) {
        if (obj == null) {
            return false;
        }

        // 获取对象的类信息
        Class<?> clazz = obj.getClass();

        // 检查是否是原始类型的封装类
        return clazz.equals(Integer.class) || clazz.equals(Character.class) ||
            clazz.equals(Boolean.class) || clazz.equals(Byte.class) ||
            clazz.equals(Short.class) || clazz.equals(Double.class) ||
            clazz.equals(Float.class) || clazz.equals(Long.class);
    }

    public static boolean isPrimitive(Class<?> clazz) {
        if (clazz == null) {
            return false;
        }
        // 检查是否是原始类型的封装类
        return clazz.isPrimitive() || clazz.equals(Integer.class) || clazz.equals(Character.class) ||
            clazz.equals(Boolean.class) || clazz.equals(Byte.class) ||
            clazz.equals(Short.class) || clazz.equals(Double.class) ||
            clazz.equals(Float.class) || clazz.equals(Long.class);
    }

    // Check if the object is a primitive array
    public static boolean isPrimitiveArray(Object obj) {
        if (obj == null) {
            return false;
        }
        Class<?> clazz = obj.getClass();
        if (!clazz.isArray()) {
            return false;
        }
        // Get the type of array element
        Class<?> componentType = clazz.getComponentType();
        // Check if the type of array elements is primitive
        return isPrimitive(componentType);
    }

    private <T> List<Object> collectionToMap(Collection<?> collection, Class<T> noTransferClass, int depth) {
        List<Object> list = new ArrayList<>();
        for (Object item : collection) {
            if (noTransferClass.isInstance(item)) {
                list.add(item);
            } else {
                list.add(beanToMap(item, noTransferClass, depth + 1));
            }
        }
        return list;
    }

    private <T> List<Object> arrayToMap(Object[] array, Class<T> noTransferClass, int depth) {
        List<Object> list = new ArrayList<>();
        for (Object item : array) {
            if (noTransferClass.isInstance(item)) {
                list.add(item);
            } else {
                list.add(beanToMap(item, noTransferClass, depth + 1));
            }
        }
        return list;
    }
}