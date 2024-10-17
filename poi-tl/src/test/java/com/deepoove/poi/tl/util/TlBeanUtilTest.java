package com.deepoove.poi.tl.util;

import com.deepoove.poi.util.TlBeanUtil;
import com.fasterxml.jackson.core.JsonProcessingException;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class TlBeanUtilTest {
    static class Person {
        private String name;
        private int age;
        private Address address;
        private List<Car> cars;
        private Car[] carArray;
        private Person father;

        public Person getFather() {
            return father;
        }

        public void setFather(Person father) {
            this.father = father;
        }

        // Getters and Setters
        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public int getAge() {
            return age;
        }

        public void setAge(int age) {
            this.age = age;
        }

        public Address getAddress() {
            return address;
        }

        public void setAddress(Address address) {
            this.address = address;
        }

        public List<Car> getCars() {
            return cars;
        }

        public void setCars(List<Car> cars) {
            this.cars = cars;
        }

        public Car[] getCarArray() {
            return carArray;
        }

        public void setCarArray(Car[] carArray) {
            this.carArray = carArray;
        }
    }

    static class Address {
        private String street;
        private String city;

        // Getters and Setters
        public String getStreet() {
            return street;
        }

        public void setStreet(String street) {
            this.street = street;
        }

        public String getCity() {
            return city;
        }

        public void setCity(String city) {
            this.city = city;
        }
    }

    static class Car {
        private String model;
        private int year;
        private List<Address> points;

        // Getters and Setters
        public String getModel() {
            return model;
        }

        public void setModel(String model) {
            this.model = model;
        }

        public int getYear() {
            return year;
        }

        public void setYear(int year) {
            this.year = year;
        }

        public List<Address> getPoints() {
            return points;
        }

        public void setPoints(List<Address> points) {
            this.points = points;
        }
    }

    @Test
    public void testBeanToMap() throws IllegalAccessException, JsonProcessingException {
        Person person = new Person();
        person.setName("John Doe");
        person.setAge(30);

        Address address = new Address();
        address.setStreet("123 Main St");
        address.setCity("Springfield");
        person.setAddress(address);

        Car car1 = new Car();
        car1.setModel("Toyota Camry");
        car1.setYear(2020);
        List<Address> addressList = new ArrayList<>();
        addressList.add(address);
        car1.setPoints(addressList);

        Car car2 = new Car();
        car2.setModel("Honda Civic");
        car2.setYear(2018);

        person.setCars(Arrays.asList(car1, car2));

        Car[] carArray = new Car[2];
        carArray[0] = new Car();
        carArray[0].setModel("Ford Mustang");
        carArray[0].setYear(2019);

        carArray[1] = new Car();
        carArray[1].setModel("Chevrolet Corvette");
        carArray[1].setYear(2021);

        person.setCarArray(carArray);
        person.setFather(person);
        TlBeanUtil beanUtil = new TlBeanUtil();
        Map<String, Object> map = beanUtil.beanToMap(person, Address.class, 0);
        System.out.println(map);
    }
}
