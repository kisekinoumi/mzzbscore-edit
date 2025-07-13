# -*- coding: utf-8 -*-
"""
Excel动漫评分排名系统 - 程序入口

这是一个面向对象的动漫评分数据处理和排名计算系统。
"""

import sys
import os

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

try:
    from app import Application
    from app.config.settings import Settings
    from app.utils.exceptions import BaseApplicationError
except ImportError as e:
    print(f"错误：无法导入必要的模块 - {e}")
    print("请确保所有依赖项已正确安装，并且模块文件存在。")
    print("运行 'pip install -r requirements.txt' 安装依赖项。")
    sys.exit(1)


def main():
    """程序主入口函数"""
    try:
        # 创建应用程序设置
        settings = Settings()
        
        # 创建应用程序实例
        app = Application(settings)
        
        # 运行应用程序
        app.run()
        
    except KeyboardInterrupt:
        print("\n程序被用户中断，正在退出...")
        
    except BaseApplicationError as e:
        print(f"应用程序错误: {e}")
        sys.exit(1)
        
    except Exception as e:
        print(f"程序发生严重错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()