#!/usr/bin/env python3
"""创建配置文件脚本"""

import os
import json

def create_config():
    """创建配置文件"""
    # 创建配置目录
    config_dir = os.path.join(os.getenv('APPDATA'), 'TextCaptureTool')
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
        print(f'配置目录已创建: {config_dir}')
    
    # 创建配置文件
    config_file = os.path.join(config_dir, 'config.json')
    config_data = {
        'tesseract_path': r'D:\teseract\tesseract.exe',
        'docx_path': r'd:\music\text_capture.docx',
        'capture_interval': 1.0,
        'min_text_length': 1,
        'max_text_length': 10000,
        'ocr_lang': 'chi_sim+eng',
        'enable_auto_save': True,
        'show_notifications': True
    }
    
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=2, ensure_ascii=False)
    
    print(f'配置文件已创建: {config_file}')
    print('Tesseract路径已设置为: D:\\teseract\\tesseract.exe')
    
    # 验证文件是否存在
    if os.path.exists(config_file):
        print('配置文件验证: 存在')
        
        # 读取并显示配置内容
        with open(config_file, 'r', encoding='utf-8') as f:
            loaded_config = json.load(f)
            print('配置文件内容:')
            print(json.dumps(loaded_config, indent=2, ensure_ascii=False))
    else:
        print('配置文件验证: 不存在')

if __name__ == '__main__':
    create_config()