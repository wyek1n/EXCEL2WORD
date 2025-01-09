import os
import pandas as pd
from src.config import Config
from src.excel_handler import ExcelHandler
from src.word_handler import WordHandler
from src.utils import clear_screen, validate_path

EXCEL_TEMPLATE_NAME = 'template.xlsx'
OUTPUT_DIR_NAME = 'Output'

class DocumentProcessor:
    def __init__(self):
        self.config = Config()
        self.excel_handler = None
        self.word_handler = None

    def show_menu(self):
        while True:
            clear_screen()
            print("\n=== XLSM2WORD 文档处理系统 ===")
            
            # 显示当前配置状态
            if self.config.word_template:
                print(f"[已配置] Word模板: {self.config.word_template}")
                # 显示输出目录
                if self.config.output_dir:
                    excel_path = os.path.join(self.config.output_dir, EXCEL_TEMPLATE_NAME)
                    if os.path.exists(excel_path):
                        try:
                            df = pd.read_excel(excel_path, engine='openpyxl')
                            if df.empty:
                                print(f"[警告] Excel模板已创建但无数据: {excel_path}")
                            else:
                                print(f"[已配置] Excel数据: {len(df)} 行")
                        except Exception:
                            print(f"[错误] Excel文件读取失败: {excel_path}")
                    else:
                        print("[未配置] Excel模板未创建")
            
            if self.config.replace_items:
                print(f"[已配置] 替换项目数量: {len(self.config.replace_items)}")
            if self.config.output_format:
                print(f"[已配置] 输出格式: {self.config.output_format}")
            
            print("\n1. 输入原始word地址")
            print("2. 输入需要替换的项目名称")
            print("3. 输入输出文件名格式")
            print("4. 开始处理")
            print("5. 退出")
            
            choice = input("\n请选择操作 (1-5): ")
            
            if choice == '1':
                self.set_word_template()
            elif choice == '2':
                self.set_replace_items()
            elif choice == '3':
                self.set_output_format()
            elif choice == '4':
                self.process_documents()
            elif choice == '5':
                break
            else:
                input("无效选择，按Enter继续...")

    def set_word_template(self):
        path = input("\n请输入原始word文件路径: ")
        if validate_path(path, file_type='word'):
            self.config.word_template = path
            # 自动设置输出目录 - 修改为在同级目录创建Output
            word_dir = os.path.dirname(os.path.abspath(path))
            output_dir = os.path.join(word_dir, OUTPUT_DIR_NAME)
            
            try:
                os.makedirs(output_dir, exist_ok=True)
                self.config.output_dir = output_dir
                print(f"\n已设置输出目录: {output_dir}")
                self.config.save_config()  # 保存配置
                input("按Enter返回菜单...")
            except Exception as e:
                input(f"\n创建输出目录失败: {e}，按Enter返回菜单...")
        else:
            input("\n无效路径，按Enter返回菜单...")

    def set_replace_items(self):
        if not self.config.word_template:
            input("\n请先设置Word模板路径，按Enter返回菜单...")
            return

        items = input("\n请输入需要替换的项目名称（用中文分号分隔）: ")
        items = [item.strip() for item in items.split('；') if item.strip()]
        if items:
            self.config.replace_items = items
            self.excel_handler = ExcelHandler(self.config)
            excel_path = os.path.join(self.config.output_dir, EXCEL_TEMPLATE_NAME)
            self.excel_handler.create_template(excel_path)
            print(f"\n已生成Excel模板文件: {excel_path}")
            self.config.save_config()  # 保存配置
            input("请完善Excel内容后按Enter返回菜单...")
        else:
            input("\n输入无效，按Enter返回菜单...")

    def set_output_format(self):
        if not self.config.word_template:
            input("\n请先设置Word模板路径，按Enter返回菜单...")
            return

        format_str = input("\n请输入输出文件名格式: ")
        if format_str:
            self.config.output_format = format_str
            self.config.save_config()  # 保存配置
            excel_path = os.path.join(self.config.output_dir, EXCEL_TEMPLATE_NAME)
            print(f"\n设置成功，请完善 {excel_path} 中的内容")
            input("完成Excel内容编辑后，按Enter返回菜单...")
        else:
            input("\n输入无效，按Enter返回菜单...")

    def process_documents(self):
        if not self.config.word_template:
            input("\n请先设置Word模板路径，按Enter返回菜单...")
            return

        try:
            print("\n开始处理文档...")
            
            # 检查并设置Excel文件路径
            excel_path = os.path.join(self.config.output_dir, EXCEL_TEMPLATE_NAME)
            if not os.path.exists(excel_path):
                raise FileNotFoundError(f"Excel文件不存在: {excel_path}")
            self.config.excel_path = excel_path
            
            # 每次处理都重新初始化handlers
            self.excel_handler = ExcelHandler(self.config)
            self.config.excel_handler = self.excel_handler
            
            # 检查Excel数据
            print(f"正在读取Excel文件: {excel_path}")
            df = self.excel_handler.read_data()
            if df.empty:
                raise ValueError("Excel文件中没有数据，请先完善Excel内容")
            
            # 初始化word处理器并处理文档
            self.word_handler = WordHandler(self.config)
            self.word_handler.process_documents()
            print("\n处理完成，按Enter返回菜单...")
            input()
        except Exception as e:
            print(f"\n处理过程中出错: {str(e)}")
            input("按Enter返回菜单...")

def main():
    processor = DocumentProcessor()
    processor.show_menu()

if __name__ == "__main__":
    main() 