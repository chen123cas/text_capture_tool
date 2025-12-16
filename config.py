#!/usr/bin/env python3
"""配置文件

存储软件的配置信息
"""

import os
from datetime import datetime

class Config:
    """配置类"""
    
    def __init__(self):
        """初始化配置"""
        # 获取程序所在目录
        app_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 默认配置
        self.default_config = {
            'docx_path': os.path.join(app_dir, 'text_capture.docx'),
            'capture_interval': 1.0,  # 文本捕获间隔（秒）
            'min_text_length': 1,  # 最小捕获文本长度
            'max_text_length': 10000,  # 最大捕获文本长度
            'enable_auto_save': True,  # 启用自动保存
            'show_notifications': True,  # 显示通知
            'text_source_tags': {
                'chrome.exe': '[网页]',
                'msedge.exe': '[网页]',
                'firefox.exe': '[网页]',
                'wechat.exe': '[微信]',
                'QQ.exe': '[QQ]',
                'winword.exe': '[Word]',
                'excel.exe': '[Excel]',
                'powerpnt.exe': '[PowerPoint]',
                'notepad.exe': '[记事本]',
                'notepad++.exe': '[Notepad++]',
                'code.exe': '[VS Code]',
                'idea.exe': '[IntelliJ]',
                'pycharm.exe': '[PyCharm]',
                'chrome.exe': '[Google Chrome]',
                'msedge.exe': '[Microsoft Edge]',
                'firefox.exe': '[Mozilla Firefox]',
                'opera.exe': '[Opera]',
                'safari.exe': '[Safari]',
                'thunderbird.exe': '[Thunderbird]',
                'outlook.exe': '[Outlook]',
                'teams.exe': '[Microsoft Teams]',
                'zoom.exe': '[Zoom]',
                'skype.exe': '[Skype]',
            }
        }
        
        # 当前配置
        self.config = self.default_config.copy()
        
        # 加载配置文件
        self.load_config()
        
    def load_config(self):
        """加载配置文件"""
        try:
            config_file = self.get_config_file_path()
            
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    import json
                    loaded_config = json.load(f)
                    self.config.update(loaded_config)
                print(f"配置文件已加载：{config_file}")
            else:
                print(f"配置文件不存在，使用默认配置：{config_file}")
                self.save_config()  # 创建默认配置文件
                
        except Exception as e:
            print(f"加载配置文件失败：{e}")
            print("使用默认配置")
            self.config = self.default_config.copy()
            
    def save_config(self):
        """保存配置文件"""
        try:
            config_file = self.get_config_file_path()
            config_dir = os.path.dirname(config_file)
            
            # 确保配置目录存在
            if not os.path.exists(config_dir):
                os.makedirs(config_dir)
                print(f"配置目录已创建：{config_dir}")
                
            # 保存配置文件
            with open(config_file, 'w', encoding='utf-8') as f:
                import json
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            print(f"配置文件已保存：{config_file}")
            
        except Exception as e:
            print(f"保存配置文件失败：{e}")
            
    def get_config_file_path(self):
        """获取配置文件路径"""
        app_data_dir = os.path.join(os.getenv('APPDATA', os.path.expanduser('~')), 'TextCaptureTool')
        return os.path.join(app_data_dir, 'config.json')
        
    def get(self, key, default=None):
        """获取配置值"""
        return self.config.get(key, default)
        
    def set(self, key, value):
        """设置配置值"""
        self.config[key] = value
        self.save_config()
        

        
    def get_docx_path(self):
        """获取Word文档路径"""
        return self.get('docx_path', self.default_config['docx_path'])
        
    def set_docx_path(self, path):
        """设置Word文档路径"""
        self.set('docx_path', path)
        
    def get_capture_interval(self):
        """获取文本捕获间隔"""
        return self.get('capture_interval', self.default_config['capture_interval'])
        
    def set_capture_interval(self, interval):
        """设置文本捕获间隔"""
        self.set('capture_interval', interval)
        
    def get_min_text_length(self):
        """获取最小捕获文本长度"""
        return self.get('min_text_length', self.default_config['min_text_length'])
        
    def set_min_text_length(self, length):
        """设置最小捕获文本长度"""
        self.set('min_text_length', length)
        
    def get_max_text_length(self):
        """获取最大捕获文本长度"""
        return self.get('max_text_length', self.default_config['max_text_length'])
        
    def set_max_text_length(self, length):
        """设置最大捕获文本长度"""
        self.set('max_text_length', length)
        

        
    def is_auto_save_enabled(self):
        """检查是否启用自动保存"""
        return self.get('enable_auto_save', self.default_config['enable_auto_save'])
        
    def is_notifications_enabled(self):
        """检查是否显示通知"""
        return self.get('show_notifications', self.default_config['show_notifications'])
        
    def get_text_source_tag(self, process_name):
        """获取文本来源标签"""
        tags = self.get('text_source_tags', self.default_config['text_source_tags'])
        return tags.get(process_name, f'[{process_name}]')
        
    def add_text_source_tag(self, process_name, tag):
        """添加文本来源标签"""
        tags = self.get('text_source_tags', self.default_config['text_source_tags'])
        tags[process_name] = tag
        self.set('text_source_tags', tags)
        
    def reset_to_default(self):
        """重置为默认配置"""
        self.config = self.default_config.copy()
        self.save_config()
        print("配置已重置为默认值")
        

# 创建全局配置对象
config = Config()