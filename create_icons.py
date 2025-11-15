#!/usr/bin/env python3
"""
Outlook Send Guard アイコン生成スクリプト
メールアイコン + 赤い×マークのアイコンを生成
"""

from PIL import Image, ImageDraw

def create_icon(size):
    """指定サイズのアイコンを作成"""
    # 背景色（白）
    img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    # スケーリング係数
    scale = size / 80

    # メール封筒を描画（青色）
    envelope_color = (0, 120, 212)  # Microsoft Blue

    # 封筒の本体（長方形）
    envelope_margin = int(10 * scale)
    envelope_rect = [
        envelope_margin,
        envelope_margin + int(10 * scale),
        size - envelope_margin,
        size - envelope_margin - int(5 * scale)
    ]
    draw.rectangle(envelope_rect, fill=envelope_color, outline=envelope_color)

    # 封筒のフラップ（三角形）
    flap_points = [
        (envelope_margin, envelope_margin + int(10 * scale)),  # 左下
        (size // 2, envelope_margin + int(30 * scale)),  # 中央上
        (size - envelope_margin, envelope_margin + int(10 * scale))  # 右下
    ]
    draw.polygon(flap_points, fill=(100, 150, 220), outline=(100, 150, 220))

    # 封筒の下部線（立体感）
    draw.line([
        (envelope_margin, envelope_margin + int(10 * scale)),
        (size // 2, envelope_margin + int(30 * scale))
    ], fill=(50, 90, 150), width=max(1, int(2 * scale)))

    draw.line([
        (size // 2, envelope_margin + int(30 * scale)),
        (size - envelope_margin, envelope_margin + int(10 * scale))
    ], fill=(50, 90, 150), width=max(1, int(2 * scale)))

    # 赤い×マークを描画（右上に配置）
    x_size = int(30 * scale)
    x_offset = size - x_size - int(2 * scale)
    y_offset = int(2 * scale)

    # 赤い円の背景
    circle_bbox = [
        x_offset,
        y_offset,
        x_offset + x_size,
        y_offset + x_size
    ]
    draw.ellipse(circle_bbox, fill=(211, 47, 47))  # 赤色

    # ×マークの線
    x_margin = int(8 * scale)
    x_width = max(2, int(3 * scale))

    # 左上から右下への線
    draw.line([
        (x_offset + x_margin, y_offset + x_margin),
        (x_offset + x_size - x_margin, y_offset + x_size - x_margin)
    ], fill='white', width=x_width)

    # 右上から左下への線
    draw.line([
        (x_offset + x_size - x_margin, y_offset + x_margin),
        (x_offset + x_margin, y_offset + x_size - x_margin)
    ], fill='white', width=x_width)

    return img

# 各サイズのアイコンを生成
sizes = [16, 32, 64, 80]

for size in sizes:
    icon = create_icon(size)
    filename = f'assets/icon-{size}.png'
    icon.save(filename, 'PNG')
    print(f'✓ {filename} を作成しました')

print('\nすべてのアイコンの作成が完了しました！')
