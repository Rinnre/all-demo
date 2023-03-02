package com.wj.aop.pointcut;

import sun.reflect.annotation.TypeAnnotation;

import java.lang.annotation.*;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2023/3/2 15:57
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ViewCountAdd {

}
