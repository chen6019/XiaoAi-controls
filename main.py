import paho.mqtt.client as mqtt
import os
import wmi
from windows_toasts import Toast, WindowsToaster
import os
import json

# 获取APPDATA目录的路径
appdata_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'Ai-controls')

# 创建mqtt_config.json文件的路径
config_path = os.path.join(appdata_path, 'mqtt_config.json')

# 从json文件中读取MQTT配置
with open(config_path, 'r') as f:
    mqtt_config = json.load(f)

# 从MQTT配置中获取值并赋值给变量
broker = mqtt_config['broker']
topic1 = mqtt_config['topic1']
topic2 = mqtt_config['topic2']
topic3 = mqtt_config['topic3']
secret_id = mqtt_config['secret_id']
port = mqtt_config['port']

# 初始化Windows通知
toaster = WindowsToaster('Python')
newToast = Toast()
info = '加载中..'
newToast.text_fields = [info]
newToast.on_activated = lambda _: print('Toast clicked!')
toaster.show_toast(newToast)


def toast(info):
    """
    显示一个toast通知。

    参数:
    - info: 要显示在通知中的文本信息。
    """
    newToast.text_fields = [info]
    toaster.show_toast(newToast)

def on_subscribe(client, userdata, mid, reason_code_list, properties):
    """
    MQTT订阅确认回调。

    参数:
    - client: MQTT客户端实例。
    - userdata: 用户自定义数据，这里未使用。
    - mid: 订阅的消息ID。
    - reason_code_list: 订阅结果的状态码列表。
    - properties: 订阅的附加属性。
    """
    for sub_result in reason_code_list:
        if sub_result >= 128:
            print("Subscribe failed")

def on_unsubscribe(client, userdata, mid, reason_code_list, properties):
    """
    MQTT取消订阅确认回调。

    参数:
    - client: MQTT客户端实例。
    - userdata: 用户自定义数据，这里未使用。
    - mid: 取消订阅的消息ID。
    - reason_code_list: 取消订阅结果的状态码列表。
    - properties: 取消订阅的附加属性。
    """
    if len(reason_code_list) == 0 or not reason_code_list[0].is_failure:
        print("Unsubscribe succeeded")
    else:
        print(f"Broker replied with failure: {reason_code_list[0]}")
    client.disconnect()

def set_brightness(value):
    wmi.WMI(namespace='wmi').WmiMonitorBrightnessMethods()[0].WmiSetBrightness(value, 0)
    
def on_message(client, userdata, message):
    """
    MQTT消息接收回调。

    参数:
    - client: MQTT客户端实例。
    - userdata: 用户自定义数据，用于存储接收到的消息。
    - message: 接收到的消息对象，包含主题和负载。
    """
    userdata.append(message.payload)
    command = message.payload.decode()
    print(f"Received `{command}` from `{message.topic}` topic")
    # 处理不同主题的消息
    if message.topic == topic1:
        # 电脑开关控制
        if command == 'on':
            toast("电脑已经开着啦")
        elif command == 'off':
            os.system("shutdown -s -t 10")
        else:
            client.publish(topic1)
    if message.topic == topic2:
        # 电脑屏幕亮度控制
        if command == 'off' or command == '0':
            set_brightness(0)
        elif command == 'on':
            set_brightness(100)
        else:
            brightness = int(command[3:])
            set_brightness(brightness)
    if message.topic == topic3:
        # 电脑远程控制
        if command == 'off':
            os.system('')
        elif command == 'on':
            os.system('shell:Applications\com.maxframing.mipc')
           
def on_connect(client, userdata, flags, reason_code, properties):
    """
    MQTT连接确认回调。

    参数:
    - client: MQTT客户端实例。
    - userdata: 用户自定义数据，这里未使用。
    - flags: 连接标志。
    - reason_code: 连接结果的状态码。
    - properties: 连接的附加属性。
    """
    if reason_code.is_failure:
        toast("连接MQTT失败: {reason_code}. 重新连接中...")
        print(f"Failed to connect: {reason_code}. loop_forever() will retry connection")
    else:
        toast("MQTT成功连接至"+broker)
        print("Connected to", broker)
        client.subscribe(topic1)
        client.subscribe(topic2) 
        client.subscribe(topic3)




# 创建MQTT客户端并配置回调
mqttc = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
mqttc.on_connect = on_connect
mqttc.on_message = on_message
mqttc.on_subscribe = on_subscribe
mqttc.on_unsubscribe = on_unsubscribe

# 设置用户数据并连接MQTT服务器
mqttc.user_data_set([])
mqttc._client_id = secret_id
mqttc.connect(broker, port)
mqttc.loop_forever()
print(f"Received the following message: {mqttc.user_data_get()}")

