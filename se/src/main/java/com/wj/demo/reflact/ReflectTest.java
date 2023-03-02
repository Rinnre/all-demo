package com.wj.demo.reflact;

import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2023/2/16 11:53
 */
@Slf4j
public class ReflectTest {
    public static void main(String[] args) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, InstantiationException {
        User user = new User();
        user.setName("张三");
        Class<? extends User> aClass = user.getClass();

        Field[] declaredFields = aClass.getDeclaredFields();
        String[] split = declaredFields[0].toString().split("\\.");
        String fieldName = split[split.length - 1];
        fieldName = upperCaseFirst(fieldName);
        String functionName = "get" + fieldName;
        Method method = aClass.getMethod(functionName);
        Object invoke = method.invoke(user);
        log.debug(invoke.toString());

    }

    private static String upperCaseFirst(String fieldName) {
        char[] chars = fieldName.toCharArray();
        if (chars[0] >= 'a' && chars[0] <= 'z') {
            chars[0] -= 32;
        }
        return new String(chars);
    }
}
