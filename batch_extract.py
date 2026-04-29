import os
import zipfile
import shutil
import tempfile
import sys
import re
import subprocess

# 压缩文件源目录
dir_path = os.path.expanduser("~/Desktop/source_archives")
# 解压后文件目录
target_dir = os.path.expanduser("~/Desktop/extracted")
# 默认解压到用户主目录下的extracted_files文件夹
# flatten False = 保持层级结构（多层），True = 单层（所有文件都放在 target_dir）
flatten = True
# deleteArchive 删除原来的压缩文档，True=删除，False=不删除
deleteArchive = False
# overwrite 是否覆盖同名文件：True=覆盖，False=跳过
overwrite = True
# add_folder 是否给解压的压缩包增加文件夹
add_folder = True
# use_7zip 是否使用7-Zip解压
use_7zip = True
# 放置7zip的地方
seven_zip_path = r"C:\Program Files\7-Zip\7z.exe"


def clean_special_chars(name):
    """清理文件名中的非法字符"""
    invalid_chars = r'[\\/:*?"<>|]'
    cleaned = re.sub(invalid_chars, "_", name)
    if not cleaned.strip():
        cleaned = "unknown_file"
    return cleaned


def get_unique_path(dst_path):
    """生成唯一路径，避免覆盖"""
    if not os.path.exists(dst_path):
        return dst_path
    base, ext = os.path.splitext(dst_path)
    i = 1
    new_path = f"{base}({i}){ext}"
    while os.path.exists(new_path):
        i += 1
        new_path = f"{base}({i}){ext}"
    return new_path


def extract_with_7zip(archive_path, target_dir):
    if not os.path.exists(seven_zip_path):
        raise Exception(f"未找到7-Zip: {seven_zip_path}")

    os.makedirs(target_dir, exist_ok=True)

    cmd = [seven_zip_path, "x", archive_path, f"-o{target_dir}", "-y", "-scsUTF-8"]

    # 尝试UTF-8编码
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="gbk", errors="ignore")
    if result.returncode != 0:
        # UTF-8失败，尝试GBK
        cmd[-1] = "-scsGBK"
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="gbk", errors="ignore")
        if result.returncode != 0:
            # 尝试自动检测编码
            cmd[-1] = "-scsAuto"
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="gbk", errors="ignore")
            if result.returncode != 0:
                raise Exception(f"7-Zip解压失败: {result.stderr}")


def safe_extract_with_zipfile(zip_path, temp_dir):
    """用标准库解压ZIP文件（仅尝试UTF-8和GBK）"""
    with zipfile.ZipFile(zip_path, "r") as zf:
        try:
            zf.extractall(temp_dir)
        except:
            # 尝试不同编码解码文件名
            for info in zf.infolist():
                try:
                    info.filename = info.filename.encode("cp437").decode("gbk", errors="ignore")
                except:
                    info.filename = clean_special_chars(info.filename)
            zf.extractall(temp_dir)


def extract_archive(archive_path, target_dir, flatten, overwrite):
    """解压压缩文件（支持ZIP和RAR）"""
    # 获取压缩文件名作为文件夹名称
    archive_filename = os.path.splitext(os.path.basename(archive_path))[0]
    cleaned_archive_name = clean_special_chars(archive_filename)

    # 确定最终目标目录，如果add_folder为True则创建以压缩文件命名的子文件夹
    final_target_dir = os.path.join(target_dir, cleaned_archive_name) if add_folder else target_dir
    os.makedirs(final_target_dir, exist_ok=True)

    with tempfile.TemporaryDirectory() as temp_dir:
        # 判断文件类型并选择合适的解压方法
        if use_7zip:
            extract_with_7zip(archive_path, temp_dir)
        else:
            # 非7-Zip模式下，仅支持ZIP
            if archive_path.lower().endswith(".zip"):
                safe_extract_with_zipfile(archive_path, temp_dir)
            else:
                raise Exception(f"不支持的文件格式，非7-Zip模式下仅支持ZIP: {archive_path}")

        # 处理解压后的文件
        for root, _, files in os.walk(temp_dir):
            for file in files:
                safe_file = clean_special_chars(file)
                src_path = os.path.join(root, file)

                if flatten:
                    dst_path = os.path.join(final_target_dir, safe_file)
                else:
                    rel_path = os.path.relpath(src_path, temp_dir)
                    safe_rel_path = clean_special_chars(rel_path)
                    dst_path = os.path.join(final_target_dir, safe_rel_path)
                    os.makedirs(os.path.dirname(dst_path), exist_ok=True)

                # 处理文件覆盖
                if os.path.exists(dst_path):
                    if overwrite:
                        os.remove(dst_path)
                    else:
                        dst_path = get_unique_path(dst_path)

                shutil.move(src_path, dst_path)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")

    print(f"压缩文件源目录: {os.path.abspath(dir_path)}")
    print(f"解压目标目录: {os.path.abspath(target_dir)}")
    print(f"使用7-Zip解压: {'是' if use_7zip else '否'}")
    print(f"解压模式: {'单层' if flatten else '多层'}")
    print(f"是否创建文件夹: {'是' if add_folder else '否'}")
    print(f"同名处理: {'覆盖' if overwrite else '自动重命名'}\n")

    if not os.path.exists(dir_path):
        print(f"错误: 源目录不存在 {dir_path}")
        sys.exit(1)

    os.makedirs(target_dir, exist_ok=True)

    archive_count = 0
    success_count = 0

    # 支持的压缩文件扩展名
    supported_extensions = (".zip", ".rar")

    for file_name in sorted(os.listdir(dir_path)):
        # 检查文件是否为支持的压缩格式
        if not file_name.lower().endswith(supported_extensions):
            continue

        archive_count += 1
        archive_path = os.path.join(dir_path, file_name)
        archive_ext = os.path.splitext(file_name)[1].lower()

        try:
            print(f"开始处理: {file_name} ({archive_ext[1:].upper()}格式)")
            extract_archive(archive_path, target_dir, flatten, overwrite)
            print(f"  解压成功: {file_name}")
            success_count += 1

            if deleteArchive:
                os.remove(archive_path)
                print(f"  已删除压缩文件: {file_name}")
        except Exception as e:
            print(f"  解压失败: {file_name}, 错误: {str(e)}")

    print(f"\n压缩文件总数: {archive_count}")
    print(f"成功解压: {success_count}")
    print(f"失败数量: {archive_count - success_count}")
