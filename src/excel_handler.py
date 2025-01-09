import pandas as pd
import os

class ExcelHandler:
    def __init__(self, config):
        self.config = config

    def create_template(self, excel_path):
        """创建Excel模板文件"""
        try:
            # 将分号分隔的字符串转换为列表
            items = []
            for item in self.config.replace_items:
                # 如果项目中包含分号，则分割成多个项目
                if '；' in item:
                    items.extend([i.strip() for i in item.split('；') if i.strip()])
                else:
                    items.append(item.strip())
            
            # 创建DataFrame，使用处理后的项目列表作为列名
            df = pd.DataFrame(columns=items)
            
            # 添加一个空行作为数据输入示例，用'N/A'替代空值
            df.loc[0] = ['N/A'] * len(items)
            
            # 保存到Excel文件
            df.to_excel(excel_path, index=False, engine='openpyxl')
            return True
        except Exception as e:
            print(f"创建Excel模板失败: {str(e)}")
            return False

    def read_data(self):
        """读取Excel数据"""
        try:
            df = pd.read_excel(self.config.excel_path, engine='openpyxl')
            # 将所有的nan值替换为'N/A'
            df = df.fillna('N/A')
            return df
        except Exception as e:
            print(f"读取Excel数据失败: {str(e)}")
            return pd.DataFrame() 