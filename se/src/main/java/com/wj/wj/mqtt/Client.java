package com.wj.wj.mqtt;

import org.eclipse.paho.client.mqttv3.MqttClient;
import org.eclipse.paho.client.mqttv3.MqttConnectOptions;

/**
 * Created by IntelliJ IDEA
 * 这是一个平凡的Class
 *
 * @author wj
 * @date 2022/10/20 15:48
 */
public class Client {
    public static void main(String[] args) throws Exception {

        String host = "tcp://localhost:1883";
        String topic = "keenyoda/deviceInfo/download/EmsGtw001";
        // clientId不能重复
        String clientId = "12345";
        // 1.设置mqtt连接属性
        MqttConnectOptions options = new MqttConnectOptions();
        options.setCleanSession(true);
        // 2.实例化mqtt客户端
        MqttClient client = new MqttClient(host, clientId);
        // 3.连接
        client.connect(options);

        client.setCallback(new PushCallback());
        while (true) {
            client.subscribe(topic, 2);
        }
        // client.disconnect();
    }
}


