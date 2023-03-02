package com.wj.aop.section;

import lombok.extern.slf4j.Slf4j;
import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Before;
import org.aspectj.lang.annotation.Pointcut;
import org.springframework.stereotype.Component;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2023/3/2 16:02
 */
@Aspect
@Component
@Slf4j
public class AddAspect {

    /**
     * 注解声明切点
     */
    @Pointcut("@annotation(com.wj.aop.pointcut.ViewCountAdd)")
    public void ViewCountAddPointcut() {

    }
    @Before("ViewCountAddPointcut()")
    private void addViewCount(JoinPoint joinPoint){
        System.out.println(joinPoint);
        log.debug("view count add");
    }
}
