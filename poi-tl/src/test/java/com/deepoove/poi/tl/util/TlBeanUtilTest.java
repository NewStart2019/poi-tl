package com.deepoove.poi.tl.util;

import com.deepoove.poi.util.TlBeanUtil;
import com.fasterxml.jackson.core.JsonProcessingException;
import org.junit.jupiter.api.Test;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

public class TlBeanUtilTest {
    public static class Person {
        public String name;
        public int age;
        public Address address;
        public List<Car> cars;
        public Car[] carArray;
        public Person father;
    }

    public static class Address {
        public String street;
        public String city;
    }

    public static class Car {
        public String model;
        public int year;
        public List<Address> points;
    }


    @Test
    public void testBeanToMap() throws IllegalAccessException, JsonProcessingException {
        Person person = new Person();
        person.name = "John Doe";
        person.age = 30;

        Address address = new Address();
        address.street = "123 Main St";
        address.city = "Springfield";
        person.address = address;

        Car car1 = new Car();
        car1.model = "Toyota Camry";
        car1.year = 2020;
        List<Address> addressList = new ArrayList<>();
        addressList.add(address);
        car1.points = addressList;

        Car car2 = new Car();
        car2.model = "Honda Civic";
        car2.year = 2018;

        person.cars = Arrays.asList(car1, car2);

        Car[] carArray = new Car[2];
        carArray[0] = new Car();
        carArray[0].model = "Ford Mustang";
        carArray[0].year = 2019;
        carArray[1] = new Car();
        carArray[1].model = "Chevrolet Corvette";
        carArray[1].year = 2021;
        person.carArray = carArray;
        person.father = person;

        TlBeanUtil beanUtil = new TlBeanUtil();
        Map<String, Object> map = beanUtil.beanToMap(person, Address.class, 0);
        map.put("atomicInteget", new AtomicInteger(100));
        System.out.println(map);

        // ---- 断言自检 ----
        check("name", "John Doe".equals(map.get("name")));
        check("age", Integer.valueOf(30).equals(map.get("age")));
        // Address 命中保留类，作为字段原样保留
        check("address preserved", map.get("address") == address);
        // 集合被转换为 List
        check("cars is List", map.get("cars") instanceof List);
        // 对象数组被转换为 List
        check("carArray is List", map.get("carArray") instanceof List);
        // 自引用：father 指向同一 Map，且不发生死循环
        check("father self reference", map.get("father") == map);
        check("atomic kept", map.get("atomicInteget") instanceof AtomicInteger);
        System.out.println("ORIGINAL TEST SCENARIO OK");
    }

    private static void check(String name, boolean ok) {
        if (!ok) {
            throw new AssertionError("FAILED: " + name);
        }
        System.out.println("OK: " + name);
    }

    /** 仅作为 noTransferClass 占位，测试对象都不是它的子类。 */
    static final class Marker { }

    enum Color { RED, GREEN }

    static class Bean {
        String name;
        Color color;
        Date when;
        List<Object> mixed;
        Object[] arr;
        Bean child;
    }

    /**
     * 实证 TlBeanUtil 的若干逻辑边界。Marker 不被任何测试对象继承，确保不触发"保留"分支。
     */
    @Test
    public void testEnum() {
        // 1) 枚举字段 + 集合里的 null/Date/字符串/包装类型
        Bean b = new Bean();
        b.name = "x";
        b.color = Color.RED;
        b.when = new Date(0);
        b.mixed = new ArrayList<>();
        b.mixed.add(null);
        b.mixed.add(new Date(0));
        b.mixed.add("str");
        b.mixed.add(123);
        b.arr = new Object[]{ new Bean(), "tail-string" };

        TlBeanUtil u = new TlBeanUtil();
        Map<String, Object> m = u.beanToMap(b, Marker.class, 0);
        System.out.println("1. enum field => " + m.get("color"));
        System.out.println("2. mixed list => " + m.get("mixed"));
        System.out.println("3. array(complex first, string second) => " + m.get("arr"));

        Bean b2 = new Bean();
        b2.arr = new Object[]{ "head-string", new Bean() };
        Map<String, Object> m2 = new TlBeanUtil().beanToMap(b2, Marker.class, 0);
        System.out.println("4. array(string first, complex second) => " + m2.get("arr"));

        // 缓存复用：转换后修改对象，再次转换
        Bean b3 = new Bean();
        b3.name = "old";
        TlBeanUtil u3 = new TlBeanUtil();
        Map<String, Object> first = u3.beanToMap(b3, Marker.class, 0);
        b3.name = "new";
        Map<String, Object> second = u3.beanToMap(b3, Marker.class, 0);
        System.out.println("5. cache reuse first=" + first.get("name") + " second=" + second.get("name"));

        // Map 含 null key
        Map<Object, Object> raw = new HashMap<>();
        raw.put(null, "v");
        try {
            Map<String, Object> mk = new TlBeanUtil().beanToMap(raw, Marker.class, 0);
            System.out.println("6. null key map => " + mk);
        } catch (Exception e) {
            System.out.println("6. null key threw => " + e);
        }

        // 深度：15 层不同对象链，验证是否有 10 层限制
        Bean root = new Bean();
        Bean cur = root;
        for (int i = 0; i < 15; i++) {
            Bean n = new Bean();
            n.name = "L" + i;
            cur.child = n;
            cur = n;
        }
        Map<String, Object> deep = new TlBeanUtil().beanToMap(root, Marker.class, 0);
        int levels = 0;
        Object node = deep;
        while (node instanceof Map && ((Map<?, ?>) node).containsKey("child")) {
            node = ((Map<?, ?>) node).get("child");
            levels++;
        }
        System.out.println("7. nested levels converted (no 10 limit?) => " + levels);
    }
}
