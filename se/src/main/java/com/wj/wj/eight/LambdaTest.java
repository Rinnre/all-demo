package com.wj.wj.eight;

/**
 * Created by IntelliJ IDEA
 * 这是一个平凡的Class
 *
 * @author wj
 * @date 2022/10/24 17:52
 */
public class LambdaTest {

    public static void main(String[] args) {
        MathOperation addOperation = (a, b) -> {
            return a + b;
        };

        MathOperation subOperation = (a, b) ->a - b;

        LambdaTest lambdaTest = new LambdaTest();

        System.out.println("a-b:"+lambdaTest.operate(2,1,subOperation));
        System.out.println("a+b:"+lambdaTest.operate(2,1,addOperation));

        ConsoleOperation consoleOperation = arg -> System.out.println(arg);

        consoleOperation.operate("hello world");
    }

    /**
     * 定义接口
     */
    interface MathOperation {
        int operate(int num1, int num2);
    }

    interface ConsoleOperation{
        void operate(String arg);
    }

    private int operate(int num1, int num2, MathOperation mathOperation) {
        return mathOperation.operate(num1, num2);
    }
}
