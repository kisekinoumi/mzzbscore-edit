"""
核心组件包

包含应用程序的核心控制器和基础抽象类。
"""

from app.core.application import Application
from app.core.base import BaseService, BaseHandler

__all__ = ["Application", "BaseService", "BaseHandler"] 