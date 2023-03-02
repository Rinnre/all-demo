package com.wj.wj.eight;


import java.util.Arrays;
import java.util.List;
import java.util.function.Supplier;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/25 8:50
 */
public class MethodTest {
    public static void main(String[] args) {
        Car car = Car.create(Car::new);
        List<Car> cars = Arrays.asList(car);

        cars.forEach(Car::collide);

        cars.forEach(Car::repair);

        final Car police  = Car.create(Car::new);
        cars.forEach(police::follow);
    }

}

class Car{

    /**
     * 构造器
     * @param supplier
     * @return
     */
    public static Car create(final Supplier<Car> supplier){
        return supplier.get();
    }
    // 静态方法
    public static void collide(final Car car){
        System.out.println("Collide "+car.toString());
    }


    /**
     * 特点类任意对象的方法引用
     * @param another
     */
    public void follow(final Car another){
        System.out.println("Following the "+another.toString());
    }


    /**
     * 特定对象的方法引用
     */
    public void repair(){
        System.out.println("Repaired "+this.toString());
    }
}
