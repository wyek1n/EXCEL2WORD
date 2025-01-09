import os
import json

class Config:
    def __init__(self):
        self.word_template = None  # Word模板路径
        self.output_dir = None     # 输出目录
        self.replace_items = None    # 需要替换的项目列表
        self.output_format = None  # 输出文件名格式
        self.excel_path = None     # Excel文件路径
        self.font_color = 'red'  # 默认红色
        self.excel_handler = None  # Excel处理器实例
        
        # 加载保存的配置
        self.load_config()

    def is_valid(self):
        """检查配置是否完整"""
        return all([
            self.word_template,  # Word模板必须存在
            self.replace_items,  # 必须有替换项
            self.output_format   # 必须有输出格式
        ])

    def save_config(self):
        """保存配置到文件"""
        config_data = {
            'word_template': self.word_template,
            'output_dir': self.output_dir,
            'replace_items': self.replace_items,
            'output_format': self.output_format,
            'excel_path': self.excel_path
        }
        
        config_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config')
        os.makedirs(config_dir, exist_ok=True)
        config_path = os.path.join(config_dir, 'config.json')
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"保存配置失败: {e}")

    def load_config(self):
        """从文件加载配置"""
        config_path = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), 
            'config', 
            'config.json'
        )
        
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                self.word_template = config_data.get('word_template')
                self.output_dir = config_data.get('output_dir')
                self.replace_items = config_data.get('replace_items', [])
                self.output_format = config_data.get('output_format')
                self.excel_path = config_data.get('excel_path')
            except Exception as e:
                print(f"加载配置失败: {e}") 