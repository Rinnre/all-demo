package com.wj.wj.mqtt;

import org.eclipse.paho.client.mqttv3.MqttClient;
import org.eclipse.paho.client.mqttv3.MqttConnectOptions;
import org.eclipse.paho.client.mqttv3.MqttMessage;

import java.util.Scanner;

/**
 * Created by IntelliJ IDEA
 * 这是一个平凡的Class
 *
 * @author wj
 * @date 2022/10/20 15:46
 */
public class Server {
    public static void main(String[] args) throws Exception {
        String host = "tcp://localhost:1883";
        String topic = "keenyoda/deviceInfo/upload/EmsGtw001";
        // clientId不能重复
        String clientId = "server";
        MqttConnectOptions options = new MqttConnectOptions();
        options.setCleanSession(true);

        MqttClient client = new MqttClient(host, clientId);
        client.connect(options);

        MqttMessage message = new MqttMessage();

        @SuppressWarnings("resource")
        Scanner scanner = new Scanner(System.in);
        System.out.println("请输入要发送的内容：");
        while (true) {
            String line = scanner.nextLine();
            message.setPayload(line.getBytes());
            client.publish(topic, message);
        }

    }
}
