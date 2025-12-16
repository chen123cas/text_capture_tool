#!/usr/bin/env python3
"""安装脚本

用于打包和分发软件
"""

import os
import sys
from setuptools import setup, find_packages

# 读取requirements.txt文件
def read_requirements():
    requirements_file = os.path.join(os.path.dirname(__file__), 'requirements.txt')
    with open(requirements_file, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f if line.strip() and not line.startswith('#')]

# 读取README.md文件作为long_description
def read_readme():
    readme_file = os.path.join(os.path.dirname(__file__), 'README.md')
    if os.path.exists(readme_file):
        with open(readme_file, 'r', encoding='utf-8') as f:
            return f.read()
    return ""

setup(
    name="text-capture-tool",
    version="1.0.0",
    description="全局文本捕获+自动写入DOCX的桌面软件",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    author="Your Name",
    author_email="your-email@example.com",
    url="https://github.com/your-username/text-capture-tool",
    packages=find_packages(),
    install_requires=read_requirements(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Desktop Environment :: File Managers",
        "Topic :: Multimedia :: Graphics :: Capture :: Screen Capture",
        "Topic :: Text Processing :: Indexing",
        "Topic :: Text Processing :: Linguistic",
        "Topic :: Utilities",
    ],
    keywords=[
        "text-capture",
        "docx",
        "desktop-application",
        "pyqt5",
    ],
    entry_points={
        "console_scripts": [
            "text-capture-tool = main:main",
        ],
    },
    python_requires=">=3.8",
    include_package_data=True,
    zip_safe=False,
)

if __name__ == '__main__':
    setup()