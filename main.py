import os
import sys
import pandas as pd
from src.config import Config
from src.excel_handler import ExcelHandler
from src.word_handler import WordHandler
from src.utils import clear_screen, validate_path

EXCEL_TEMPLATE_NAME = 'template.xlsx'
output_DIR_NAME = 'output'

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
                    items_path = os.path.join(self.config.output_dir, 'items.txt')
                    
                    # 显示items.txt状态
                    if os.path.exists(items_path):
                        try:
                            with open(items_path, 'r', encoding='utf-8') as f:
                                items = [i.strip() for i in f.read().split('；') if i.strip()]
                                if items:
                                    print(f"[已配置] 待替换项: {len(items)}项 ({items_path})")
                                else:
                                    print(f"[未配置] 待替换项为空 ({items_path})")
                        except Exception:
                            print(f"[错误] 无法读取待替换项 ({items_path})")
                    
                    # 显示Excel状态
                    if os.path.exists(excel_path):
                        try:
                            df = pd.read_excel(excel_path, engine='openpyxl')
                            if df.empty:
                                print(f"[警告] Excel模板已创建但无数据 ({excel_path})")
                            else:
                                print(f"[已配置] Excel模板: {len(df.columns)}项 ({excel_path})")
                        except Exception:
                            print(f"[错误] Excel文件读取失败 ({excel_path})")
                    else:
                        print(f"[未配置] Excel模板未创建 ({excel_path})")
            
            if self.config.output_format:
                print(f"[已配置] 输出文件名格式: {self.config.output_format}")
            
            # 显示当前字体颜色设置
            font_color = getattr(self.config, 'font_color', 'red')  # 默认红色
            print(f"[已配置] 替换项字体颜色: {'红色' if font_color == 'red' else '黑色'}")
            
            print("\n1. 输入Word模版地址")
            print("2. 输入待替换项名称（中文分号分隔）")
            print("3. 输入生成文件名格式（含待替换项名）")
            print("4. 打开Excel模板（需手动完善目标值）")
            print("5. 设置替换项字体颜色")
            print("6. 开始替换并生成文件")
            print("7. 退出")
            
            choice = input("\n请选择操作 (1-7): ")
            
            if choice == '1':
                self.set_word_template()
            elif choice == '2':
                self.set_replace_items()
            elif choice == '3':
                self.set_output_format()
            elif choice == '4':
                self.open_excel_template()
            elif choice == '5':
                self.set_font_color()
            elif choice == '6':
                self.process_documents()
            elif choice == '7':
                break
            else:
                input("无效选择，按Enter继续...")

    def set_word_template(self):
        path = input("\n请输入原始word文件路径: ")
        if validate_path(path, file_type='word'):
            self.config.word_template = path
            
            # 修改为在项目根目录创建output
            project_root = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(project_root, output_DIR_NAME)
            
            try:
                os.makedirs(output_dir, exist_ok=True)
                self.config.output_dir = output_dir
                
                # 创建items.txt文件
                items_file = os.path.join(output_dir, 'items.txt')
                if not os.path.exists(items_file):
                    with open(items_file, 'w', encoding='utf-8') as f:
                        f.write('')  # 创建空文件
                
                print(f"\n已设置输出目录: {output_dir}")
                print(f"已创建替换项目文件: {items_file}")
                self.config.save_config()
                input("按Enter返回菜单...")
            except Exception as e:
                input(f"\n创建目录或文件失败: {e}，按Enter返回菜单...")
        else:
            input("\n无效路径，按Enter返回菜单...")

    def set_replace_items(self):
        if not self.config.output_dir:
            raise ValueError("输出目录未设置")
        if not self.config.word_template:
            input("\n请先设置Word模板路径，按Enter返回菜单...")
            return

        items_file = os.path.join(self.config.output_dir, 'items.txt')
        
        print("\n请选择输入方式：")
        print("1. 直接输入（适合项目较少的情况）")
        print("2. 使用已有的items.txt文件")
        
        choice = input("\n请选择 (1-2): ")
        
        items = []
        if choice == '1':
            print("\n请输入需要替换的项目名称（用中文分号分隔）")
            print("注意：如果项目较多，建议使用items.txt文件")
            items_input = input(": ")
            # 不在这里分割，而是直接存储原始输入
            items = [items_input] if items_input.strip() else []
            
            # 将输入的项目写入items.txt
            try:
                with open(items_file, 'w', encoding='utf-8') as f:
                    f.write(items_input)  # 直接写入原始输入
            except Exception as e:
                print(f"\n写入items.txt失败: {str(e)}")
                
        elif choice == '2':
            if not os.path.exists(items_file):
                input("\nitems.txt文件不存在，请先使用选项1输入项目，按Enter返回菜单...")
                return
            
            try:
                with open(items_file, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    items = [content] if content else []
                    
                if not items:
                    input("\nitems.txt文件为空，请先编辑文件内容，按Enter返回菜单...")
                    return
                    
                print("\n从items.txt读取的项目：")
                # 显示分割后的项目列表
                split_items = [i.strip() for i in content.split('；') if i.strip()]
                for i, item in enumerate(split_items, 1):
                    print(f"{i}. {item}")
                print(f"\n共读取到 {len(split_items)} 个项目")
                    
            except Exception as e:
                input(f"\n读取items.txt失败: {str(e)}，按Enter返回菜单...")
                return
        else:
            input("\n无效选择，按Enter返回菜单...")
            return

        if not items:
            input("\n未获取到有效的替换项目，按Enter返回菜单...")
            return

        if choice == '1':
            print(f"\n成功读取 {len(items)} 个替换项目:")
            for i, item in enumerate(items, 1):
                print(f"{i}. {item}")
            print(f"\n项目已保存到: {items_file}")

        confirm = input("\n确认是否使用这些替换项目？(y/n，直接回车默认为y): ").strip().lower()
        if not confirm or confirm == 'y':  # 空输入或'y'都确认
            self.config.replace_items = items
            self.excel_handler = ExcelHandler(self.config)
            excel_path = os.path.join(self.config.output_dir, EXCEL_TEMPLATE_NAME)
            self.excel_handler.create_template(excel_path)
            print(f"\n已生成Excel模板文件: {excel_path}")
            self.config.save_config()
            input("请完善Excel内容后按Enter返回菜单...")
        else:
            input("\n已取消，按Enter返回菜单...")

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

    def open_excel_template(self):
        """使用系统默认程序打开Excel模板文件"""
        if not self.config.output_dir:
            input("\n请先设置Word模板路径，按Enter返回菜单...")
            return
        
        excel_path = os.path.join(self.config.output_dir, EXCEL_TEMPLATE_NAME)
        if not os.path.exists(excel_path):
            input("\nExcel模板文件不存在，请先设置替换项目，按Enter返回菜单...")
            return
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(excel_path)
            elif os.name == 'posix':  # macOS 和 Linux
                import subprocess
                if sys.platform == 'darwin':  # macOS
                    subprocess.run(['open', excel_path])
                else:  # Linux
                    subprocess.run(['xdg-open', excel_path])
            print("\n已打开Excel模板文件")
            input("编辑完成后按Enter返回菜单...")
        except Exception as e:
            input(f"\n打开文件失败: {str(e)}，按Enter返回菜单...")

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

    def set_font_color(self):
        """设置替换项的字体颜色"""
        print("\n请选择替换项的字体颜色：")
        print("1. 红色（默认）")
        print("2. 黑色")
        
        choice = input("\n请选择 (1-2): ").strip()
        
        if choice == '1':
            self.config.font_color = 'red'
            print("\n已设置替换项字体为红色")
        elif choice == '2':
            self.config.font_color = 'black'
            print("\n已设置替换项字体为黑色")
        else:
            print("\n无效选择，将使用默认的红色")
            self.config.font_color = 'red'
        
        self.config.save_config()
        input("按Enter返回菜单...")

def main():
    processor = DocumentProcessor()
    processor.show_menu()

if __name__ == "__main__":
    main() 