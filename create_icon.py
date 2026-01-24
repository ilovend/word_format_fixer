#!/usr/bin/env python3
"""
创建Electron应用图标
"""

import os
from PIL import Image, ImageDraw, ImageFont

# 创建图标目录
icon_dir = "electron/build"
os.makedirs(icon_dir, exist_ok=True)

# 创建128x128像素的图标
icon_size = (128, 128)
icon = Image.new("RGB", icon_size, color="#2B579A")
draw = ImageDraw.Draw(icon)

# 绘制简单的图标
# 绘制文字
text = "WF"
font_size = 64
try:
    # 尝试使用系统字体
    font = ImageFont.truetype("arial.ttf", font_size)
except IOError:
    # 如果没有找到字体，使用默认字体
    font = ImageFont.load_default()
    font_size = 48

# 计算文字位置
text_bbox = draw.textbbox((0, 0), text, font=font)
text_width = text_bbox[2] - text_bbox[0]
text_height = text_bbox[3] - text_bbox[1]
x = (icon_size[0] - text_width) // 2
y = (icon_size[1] - text_height) // 2

# 绘制白色文字
draw.text((x, y), text, fill="white", font=font)

# 保存为PNG格式
png_path = os.path.join(icon_dir, "icon.png")
icon.save(png_path)
print(f"PNG图标已创建: {png_path}")

# 尝试保存为ICO格式（需要安装PIL的ICO插件）
try:
    ico_path = os.path.join(icon_dir, "icon.ico")
    # 创建不同尺寸的图标
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128)]
    icon.save(ico_path, format="ICO", sizes=icon_sizes)
    print(f"ICO图标已创建: {ico_path}")
except Exception as e:
    print(f"无法创建ICO图标: {e}")
    print("请手动创建ICO图标或使用在线转换工具将PNG转换为ICO")
