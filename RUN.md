# CARLA Temporal Collage Prompting 运行指南
1. 先启动 CARLA 0.9.14 仿真客户端 CarlaUE4
2. 激活虚拟环境 venv37
    Windows CMD 执行：
    venv37\Scripts\activate
3. 启动数据采集（常驻运行，不能关闭窗口）
    python data_collect.py
4. 新开终端，可视化查看相机输出画面
    python visualize.py