import os
import sys
import re
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PIL import Image

ROOT_DIR = "图片主文件夹路径"
OUTPUT_PDF = os.path.join(ROOT_DIR, "合并图表.pdf")

# 页边距与间距设置（单位：厘米）
MARGIN_CM = 0.7
GAP_MM = 5

PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN_PT = MARGIN_CM * 72 / 2.54
USABLE_WIDTH = PAGE_WIDTH - 2 * MARGIN_PT
USABLE_HEIGHT = PAGE_HEIGHT - 2 * MARGIN_PT
GAP_PT = GAP_MM * 72 / 25.4

try:
    RESAMPLE = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE = Image.ANTIALIAS


def classify_image(filename):
    """根据文件名关键词返回图片类别，用于决定排版方式"""
    if "产品收益业绩统计" in filename or "资产配置情况" in filename:
        return "no_shrink"
    elif any(key in filename for key in ["产品持仓成分股分布", "持仓期货品种分布"]):
        return "holding_layout"
    else:
        return "expandable"


def load_images(folder_path):
    """加载文件夹内所有PNG图片，并按类别分好"""
    paths = sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".png")
    ])
    categories = {"no_shrink": [], "holding_layout": [], "expandable": []}
    images = {}
    for path in paths:
        try:
            img = Image.open(path).convert("RGB")
            cat = classify_image(os.path.basename(path))
            categories[cat].append(path)
            images[path] = img
        except Exception as e:
            print(f"跳过 {path}: {e}")
    return categories, images


def layout_holding_row(holding_paths, y, c, images):
    stock_path = next((p for p in holding_paths if "产品持仓成分股分布" in os.path.basename(p)), None)
    futures_path = next((p for p in holding_paths if "持仓期货品种分布" in os.path.basename(p)), None)

    MAX_FUTURES_WIDTH = 8 * 72 / 2.54   # 期货图最大宽度8cm
    MAX_STOCK_HEIGHT = 3 * 72 / 2.54    # 股票图最大高度3cm

    total_w = USABLE_WIDTH
    stock_w_target = total_w * 4 / 7
    futures_w_target = min(total_w * 3 / 7, MAX_FUTURES_WIDTH)

    if stock_path and futures_path:
        stock_img = images[stock_path]
        futures_img = images[futures_path]
        stock_scale = stock_w_target / stock_img.size[0]
        futures_scale = futures_w_target / futures_img.size[0]
        stock_w, stock_h = stock_img.size[0] * stock_scale, stock_img.size[1] * stock_scale
        futures_w, futures_h = futures_img.size[0] * futures_scale, futures_img.size[1] * futures_scale
        row_h = max(stock_h, futures_h)
        y -= row_h
        x = MARGIN_PT
        c.drawImage(stock_path, x, y, width=stock_w, height=stock_h, preserveAspectRatio=True, mask='auto')
        x += stock_w + GAP_PT
        c.drawImage(futures_path, x, y, width=futures_w, height=futures_h, preserveAspectRatio=True, mask='auto')
    elif stock_path:
        img = images[stock_path]
        w, h = img.size
        scale = USABLE_WIDTH / w
        h_scaled = h * scale
        if h_scaled > MAX_STOCK_HEIGHT:
            scale = MAX_STOCK_HEIGHT / h
            h_scaled = MAX_STOCK_HEIGHT
        w_scaled = w * scale
        y -= h_scaled
        c.drawImage(stock_path, MARGIN_PT, y, width=w_scaled, height=h_scaled, preserveAspectRatio=True, mask='auto')
    elif futures_path:
        img = images[futures_path]
        w, h = img.size
        w_scaled = min(USABLE_WIDTH * 2 / 7, MAX_FUTURES_WIDTH)
        scale = w_scaled / w
        h_scaled = h * scale
        y -= h_scaled
        c.drawImage(futures_path, MARGIN_PT, y, width=w_scaled, height=h_scaled, preserveAspectRatio=True, mask='auto')
    return y


def create_page(c, groups, images):
    """为一个子文件夹的所有图片生成一个PDF页面"""
    y = PAGE_HEIGHT - MARGIN_PT
    spacing = 0
    remain_paths = groups["expandable"][:]

    stat_path = next((p for p in groups["no_shrink"] if "产品收益业绩统计" in os.path.basename(p)), None)
    alloc_path = next((p for p in groups["no_shrink"] if "资产配置情况" in os.path.basename(p)), None)

    # 第一行：收益统计（自适应宽度，高度不限）
    if stat_path:
        img = images[stat_path]
        iw, ih = img.size
        scale = USABLE_WIDTH / iw
        h = ih * scale
        y -= h
        c.drawImage(stat_path, MARGIN_PT, y, width=USABLE_WIDTH, height=h, preserveAspectRatio=True, mask='auto')
        y -= spacing

    # 第二行：资产配置（最大高度6cm）
    if alloc_path:
        MAX_ALLOC_HEIGHT = 6 * 72 / 2.54
        img = images[alloc_path]
        iw, ih = img.size
        scale = USABLE_WIDTH / iw
        h = ih * scale
        if h > MAX_ALLOC_HEIGHT:
            scale = MAX_ALLOC_HEIGHT / ih
            h = MAX_ALLOC_HEIGHT
        draw_w = iw * scale
        y -= h
        c.drawImage(alloc_path, MARGIN_PT, y, width=draw_w, height=h, preserveAspectRatio=True, mask='auto')
        y -= spacing

    # 第三行：持仓特殊布局（两张图并排）
    if groups["holding_layout"]:
        y = layout_holding_row(groups["holding_layout"], y, c, images)

    # 其余图片：均匀分割剩余高度
    n = len(remain_paths)
    if n and y > MARGIN_PT:
        unit_h = (y - MARGIN_PT) / n
        for p in remain_paths:
            img = images[p]
            iw, ih = img.size
            scale = USABLE_WIDTH / iw
            draw_h = ih * scale
            if draw_h > unit_h:
                scale = unit_h / ih
                draw_h = unit_h
            draw_w = iw * scale
            y -= draw_h
            c.drawImage(p, MARGIN_PT, y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask='auto')
            y -= spacing


def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]


def main():
    if ROOT_DIR == "图片主文件夹路径" or not os.path.exists(ROOT_DIR):
        print("错误：请先修改脚本顶部的 ROOT_DIR 为实际存在的图片文件夹路径。")
        sys.exit(1)

    c = canvas.Canvas(OUTPUT_PDF, pagesize=A4)
    subfolders = [f for f in os.listdir(ROOT_DIR) if os.path.isdir(os.path.join(ROOT_DIR, f))]
    subfolders.sort(key=natural_sort_key)

    for folder_name in subfolders:
        folder_path = os.path.join(ROOT_DIR, folder_name)
        groups, images = load_images(folder_path)
        total = sum(len(v) for v in groups.values())
        if total == 0:
            continue
        create_page(c, groups, images)
        c.showPage()

    c.save()
    print(f"PDF已生成：{OUTPUT_PDF}")


if __name__ == "__main__":
    main()