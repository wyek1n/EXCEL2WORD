import os
import platform

def clear_screen():
    """清除控制台屏幕"""
    os.system('cls' if platform.system() == 'Windows' else 'clear')

def validate_path(path, file_type=None):
    """
    验证文件路径
    
    Args:
        path: 文件路径
        file_type: 文件类型 ('word', 'excel', None)
    """
    if not os.path.exists(path):
        return False
        
    if file_type == 'word':
        return path.lower().endswith(('.doc', '.docx'))
    elif file_type == 'excel':
        return path.lower().endswith(('.xls', '.xlsx', '.xlsm'))
        
    return True

def format_filename(format_str, replace_dict):
    """
    格式化文件名
    
    Args:
        format_str: 格式字符串
        replace_dict: 替换字典
    """
    result = format_str
    for key, value in replace_dict.items():
        result = result.replace(key, str(value))
    return result 