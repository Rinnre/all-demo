package com.wj.demo.reflact;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2023/2/16 11:51
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class User {
    private String name;

    private int age;

    private Dog dog;
}
