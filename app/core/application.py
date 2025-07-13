"""
主应用程序控制器

协调所有服务组件，提供统一的业务逻辑接口。
"""

import sys
import io
from typing import Optional, Dict, Any

from app.core.base import BaseService
from app.services.excel_service import ExcelService
from app.services.ranking_service import RankingService
from app.models.data_models import ProcessingConfig, RankingResult
from app.config.settings import Settings
from app.config.constants import APP_NAME, APP_VERSION
from app.utils.logger import Logger
from app.utils.exceptions import (
    BaseApplicationError, InitializationError, 
    ProcessingInterruptedError, ConfigurationError
)


class Application(BaseService):
    """主应用程序控制器"""
    
    def __init__(self, settings: Optional[Settings] = None):
        """
        初始化应用程序
        
        Args:
            settings: 应用程序设置
        """
        # 初始化日志系统
        self._logger_instance = Logger()
        super().__init__(self._logger_instance.get_logger("Application"))
        
        # 初始化配置
        self._settings = settings or Settings()
        
        # 初始化服务
        self._excel_service: Optional[ExcelService] = None
        self._ranking_service: Optional[RankingService] = None
        
        # 应用程序状态
        self._running = False
        
        self.logger.info(f"{APP_NAME} v{APP_VERSION} 初始化开始")
    
    def initialize(self) -> bool:
        """
        初始化应用程序
        
        Returns:
            bool: 初始化是否成功
        """
        try:
            self.logger.info("应用程序初始化开始")
            
            # 设置编码
            self._setup_encoding()
            
            # 初始化服务
            self._initialize_services()
            
            # 验证配置
            self._validate_configuration()
            
            self._set_initialized(True)
            self.logger.info("应用程序初始化成功")
            return True
            
        except Exception as e:
            self.logger.error(f"应用程序初始化失败: {e}")
            return False
    
    def _setup_encoding(self):
        """设置UTF-8编码"""
        try:
            if sys.platform == 'win32':
                if hasattr(sys.stdout, 'buffer'):
                    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
                if hasattr(sys.stderr, 'buffer'):
                    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
            self.logger.debug("编码设置成功")
        except Exception as e:
            self.logger.warning(f"编码设置失败: {e}")
    
    def _initialize_services(self):
        """初始化所有服务"""
        try:
            # 初始化Excel服务
            self._excel_service = ExcelService(self._logger_instance)
            if not self._excel_service.initialize():
                raise InitializationError("Excel服务初始化失败")
            
            # 初始化排名服务
            self._ranking_service = RankingService(self._logger_instance)
            if not self._ranking_service.initialize():
                raise InitializationError("排名服务初始化失败")
            
            self.logger.info("所有服务初始化成功")
            
        except Exception as e:
            self.logger.error(f"服务初始化失败: {e}")
            raise InitializationError(f"服务初始化失败: {e}")
    
    def _validate_configuration(self):
        """验证配置"""
        try:
            # 验证输入文件
            input_file = self._settings.input_file
            if not input_file:
                raise ConfigurationError("输入文件路径未配置")
            
            # 验证输出文件
            if not self._settings.output_file_monthly:
                raise ConfigurationError("首月评分输出文件路径未配置")
            
            if not self._settings.output_file_final:
                raise ConfigurationError("完结评分输出文件路径未配置")
            
            self.logger.debug("配置验证通过")
            
        except Exception as e:
            self.logger.error(f"配置验证失败: {e}")
            raise
    
    def run(self):
        """运行应用程序主循环"""
        try:
            if not self.is_initialized:
                if not self.initialize():
                    print("应用程序初始化失败，无法启动")
                    return
            
            self._running = True
            self.logger.info("应用程序启动")
            
            print(f"欢迎使用 {APP_NAME} v{APP_VERSION}")
            print("=" * 50)
            
            # 主循环
            while self._running:
                try:
                    choice = self._get_user_choice()
                    
                    if choice == 'Q':
                        self._shutdown()
                        break
                    
                    self._execute_operation(choice)
                    
                except KeyboardInterrupt:
                    print("\n\n程序被用户中断")
                    self._shutdown()
                    break
                except Exception as e:
                    self.logger.error(f"主循环中发生错误: {e}")
                    print(f"发生错误: {e}")
                    
                    # 询问是否继续
                    try:
                        continue_choice = input("是否继续运行程序？(Y/N): ").strip().upper()
                        if continue_choice != 'Y':
                            self._shutdown()
                            break
                    except:
                        print("无法获取用户输入，程序将退出")
                        self._shutdown()
                        break
            
        except Exception as e:
            self.logger.error(f"应用程序运行时发生严重错误: {e}")
            print(f"程序发生严重错误: {e}")
        finally:
            self._cleanup()
    
    def _get_user_choice(self) -> str:
        """获取用户选择"""
        max_attempts = self._settings.get("max_input_attempts", 3)
        attempts = 0
        
        while attempts < max_attempts:
            try:
                print("\n请选择要执行的操作:")
                print("1. 编辑成首月评分表格")
                print("2. 编辑成完结评分表格")
                print("Q. 退出")
                
                choice = input("请输入选项 (1/2/Q): ").strip().upper()
                
                if choice in ['1', '2', 'Q']:
                    return choice
                else:
                    print("无效的输入，请输入 1、2 或 Q")
                    attempts += 1
                    
            except KeyboardInterrupt:
                raise
            except EOFError:
                print("\n输入流已结束，程序将退出")
                return 'Q'
            except Exception as e:
                self.logger.warning(f"获取用户输入时发生错误: {e}")
                attempts += 1
        
        print(f"尝试次数已达上限（{max_attempts}），程序将退出")
        return 'Q'
    
    def _execute_operation(self, choice: str):
        """执行用户选择的操作"""
        try:
            if choice == '1':
                self._process_monthly_scores()
            elif choice == '2':
                self._process_final_scores()
            else:
                print("无效的操作选择")
                
        except ProcessingInterruptedError:
            print("操作被用户中断")
        except Exception as e:
            self.logger.error(f"执行操作 {choice} 时发生错误: {e}")
            raise
    
    def _process_monthly_scores(self):
        """处理首月评分"""
        try:
            print("开始生成首月评分表格...")
            self.logger.info("开始生成首月评分表格")
            
            # 创建处理配置
            config = ProcessingConfig.for_monthly_processing(
                input_file=self._settings.input_file,
                output_file=self._settings.output_file_monthly
            )
            
            # 执行处理
            result = self.process_anime_scores(config)
            
            # 显示结果
            self._display_processing_result(result, "首月评分表格")
            
        except Exception as e:
            self.logger.error(f"处理首月评分时发生错误: {e}")
            raise
    
    def _process_final_scores(self):
        """处理完结评分"""
        try:
            print("\"编辑成完结评分表格\"功能尚未实现")
            self.logger.warning("完结评分功能被调用，但尚未实现")
            
            # TODO: 实现完结评分处理逻辑
            # config = ProcessingConfig.for_final_processing(
            #     input_file=self._settings.input_file,
            #     output_file=self._settings.output_file_final
            # )
            # result = self.process_anime_scores(config)
            # self._display_processing_result(result, "完结评分表格")
            
        except Exception as e:
            self.logger.error(f"处理完结评分时发生错误: {e}")
            raise
    
    def process_anime_scores(self, config: ProcessingConfig) -> RankingResult:
        """
        处理动漫评分数据
        
        Args:
            config: 处理配置
            
        Returns:
            RankingResult: 处理结果
        """
        try:
            self.logger.info(f"开始处理动漫评分: {config.operation_type}")
            
            # 1. 读取Excel数据
            if self._settings.show_progress:
                print("正在读取Excel数据...")
            
            original_data = self._excel_service.read_file(config.input_file)
            self.logger.info(f"成功读取 {len(original_data)} 条记录")
            
            # 2. 处理排名
            if self._settings.show_progress:
                print("正在处理排名数据...")
            
            ranking_result = self._ranking_service.process_rankings(original_data)
            
            # 3. 写入结果
            if self._settings.show_progress:
                print("正在写入结果文件...")
            
            # 设置结果的输入文件信息（用于Excel服务）
            ranking_result.input_file = config.input_file
            ranking_result.apply_styles = config.apply_styles
            
            success = self._excel_service.write_file(config.output_file, ranking_result)
            
            if not success:
                raise BaseApplicationError("写入结果文件失败")
            
            self.logger.info(f"动漫评分处理完成: {config.operation_type}")
            return ranking_result
            
        except KeyboardInterrupt:
            raise ProcessingInterruptedError("用户中断了处理过程")
        except Exception as e:
            self.logger.error(f"处理动漫评分时发生错误: {e}")
            raise
    
    def _display_processing_result(self, result: RankingResult, operation_name: str):
        """显示处理结果"""
        try:
            print(f"\n{operation_name}处理完成！")
            print("=" * 50)
            
            # 基本统计
            summary = result.get_summary()
            print(f"总处理条目: {summary['total_processed']}")
            print(f"有效条目: {summary['total_valid']}")
            print(f"排除条目: {summary['total_excluded']}")
            print(f"成功率: {summary['success_rate']}")
            print(f"处理时间: {summary['processing_time']}")
            
            # 错误和警告
            if result.has_errors:
                print(f"\n错误 ({summary['error_count']}):")
                for error in result.errors[:5]:  # 只显示前5个错误
                    print(f"  - {error}")
                if len(result.errors) > 5:
                    print(f"  ... 还有 {len(result.errors) - 5} 个错误")
            
            if result.has_warnings:
                print(f"\n警告 ({summary['warning_count']}):")
                for warning in result.warnings[:3]:  # 只显示前3个警告
                    print(f"  - {warning}")
                if len(result.warnings) > 3:
                    print(f"  ... 还有 {len(result.warnings) - 3} 个警告")
            
            # 详细统计（如果启用）
            if self._settings.show_progress:
                stats = self._ranking_service.get_ranking_statistics(result)
                if "platform_stats" in stats:
                    print("\n平台排名统计:")
                    for platform, platform_stats in stats["platform_stats"].items():
                        coverage = platform_stats.get("coverage_rate", 0) * 100
                        print(f"  {platform}: {platform_stats.get('ranked_count', 0)}/{platform_stats.get('total_count', 0)} ({coverage:.1f}%)")
            
            print("=" * 50)
            
        except Exception as e:
            self.logger.error(f"显示处理结果时发生错误: {e}")
    
    def _shutdown(self):
        """关闭应用程序"""
        try:
            self.logger.info("应用程序关闭中...")
            self._running = False
            print("感谢使用，程序已退出")
        except Exception as e:
            self.logger.error(f"关闭应用程序时发生错误: {e}")
    
    def _cleanup(self):
        """清理资源"""
        try:
            # 关闭日志系统
            if hasattr(self, '_logger_instance'):
                # 保存设置（如果需要）
                pass
            
            self.logger.info("资源清理完成")
        except Exception as e:
            print(f"资源清理时发生错误: {e}")
    
    @property
    def settings(self) -> Settings:
        """获取应用程序设置"""
        return self._settings
    
    @property
    def excel_service(self) -> Optional[ExcelService]:
        """获取Excel服务"""
        return self._excel_service
    
    @property
    def ranking_service(self) -> Optional[RankingService]:
        """获取排名服务"""
        return self._ranking_service
    
    def get_version_info(self) -> Dict[str, str]:
        """获取版本信息"""
        return {
            "app_name": APP_NAME,
            "app_version": APP_VERSION,
            "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        } 