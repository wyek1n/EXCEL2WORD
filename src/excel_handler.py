import pandas as pd
import os

class ExcelHandler:
    def __init__(self, config):
        self.config = config

    def create_template(self, excel_path):
        """创建Excel模板文件"""
        # 创建DataFrame，使用替换项作为列名
        df = pd.DataFrame(columns=self.config.replace_items)
        
        # 保存为xlsx文件
        df.to_excel(excel_path, index=False, engine='openpyxl')
        self.config.excel_path = excel_path

    def read_data(self):
        """读取Excel数据"""
        if not self.config.excel_path or not os.path.exists(self.config.excel_path):
            raise FileNotFoundError(f"Excel文件不存在: {self.config.excel_path}")
        
        print(f"正在读取Excel文件: {self.config.excel_path}")
        df = pd.read_excel(self.config.excel_path, engine='openpyxl')
        
        # 检查数据
        if df.empty:
            print("警告: Excel文件中没有数据行")
        else:
            print(f"成功读取Excel数据，共 {len(df)} 行")
            
            # 只在有空值时显示警告
            empty_cells = df.isnull().sum().sum()
            if empty_cells > 0:
                print(f"警告: 存在 {empty_cells} 个空单元格")
        
        return df 