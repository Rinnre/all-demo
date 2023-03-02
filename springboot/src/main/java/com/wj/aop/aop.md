## AOP

### 理论

- 切面：拦截器类，其中会定义切点以及通知
- 切点：具体拦截某个业务点
- 通知：切面当中的方法，声明通知方法在目标业务层的执行位置，通知类型如下：
    1. 前置通知：@Before 在目标业务方法执行之前执行
    2. 后置通知：@After 在目标业务方法执行之后执行
    3. 返回通知：@AfterReturning 在目标业务方法返回结果之后执行
    4. 异常通知：@AfterThrowing 在目标业务方法抛出异常之后
    5. 环绕通知：@Around 功能强大，可代替以上四种通知，还可以控制目标业务方法是否执行以及何时执行

### 实现方式

1. 定义切点

```java
// 作用范围-方法
@Target(ElementType.METHOD)
// 声明周期-运行时
@Retention(RetentionPolicy.RUNTIME)
public @interface ViewCountAdd {
}
```

2. 定义切面
```java
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
```
3. 业务代码增强
```java
@RestController
public class ArticleController {
    @ViewCountAdd
    @GetMapping("/article")
    public Article getArticle(){
        return new Article();
    }
}
```

### 原理
***动态代理***
- 动态代理和静态代理的区别是，静态代理的的代理类是我们自己定义好的，在程序运行之前就已经变异完成，但是动态代理的代理类是在程序运行时创建的。 
- 相比于静态代理，动态代理的优势在于可以很方便的对代理类的函数进行统一的处理，而不用修改每个代理类中的方法。
