package com.wj;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;
import org.springframework.boot.jdbc.DatabaseDriver;
import org.springframework.boot.jdbc.EmbeddedDatabaseConnection;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date ${DATE} ${TIME}
 */
@SpringBootApplication(exclude = DataSourceAutoConfiguration.class)
public class Main {
    public static void main(String[] args) {
        SpringApplication.run(Main.class,args);
    }
}
