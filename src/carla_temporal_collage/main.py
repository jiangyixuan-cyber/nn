import sys
import os
# 补全完整CARLA dist路径
sys.path.append('D:\WindowsNoEditor\PythonAPI\carla\dist')

import carla
import time

# CARLA仿真环境初始化
try:
    client = carla.Client('127.0.0.1', 2000)
    client.set_timeout(10.0)
    world = client.get_world()
    print('CARLA客户端连接成功')
except Exception as e:
    print('连接失败：',e)