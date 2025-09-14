# -*- coding: utf-8 -*-
"""
💾 桌面信息扫描工具 v2.4
功能：扫描用户桌面信息，定时拍照，每10秒一次，最多60秒
优化：聚焦桌面目录、快捷方式解析、最近修改文件、JSON 输出、用户配置
"""

import os
import cv2
import string
import shutil
import platform
import winreg
import win32com.client
import win32api
import win32con
import json
import argparse
from datetime import datetime, timedelta
import time
from zipfile import ZipFile
from tqdm import tqdm

# ==================== 配置参数 ====================
OUTPUT_IMAGE_PREFIX = "photo"
INFO_TXT_FILENAME = "desktop_info.txt"
JSON_OUTPUT_FILENAME = "desktop_info.json"
COMPRESSED_ZIP_FILENAME = "desktop_scan.zip"
MAX_FILES = 10000
PHOTO_INTERVAL = 10
MAX_PHOTO_TIME = 60
MAX_DEPTH = 3
RECENT_DAYS = 30
EXCLUDE_DIRS = ['AppData', '$RECYCLE.BIN', 'System Volume Information', 'Temp']
TARGET_EXTENSIONS = ['.txt', '.docx', '.pdf', '.lnk', '.xlsx', '.pptx']
DEBUG = True

_sys_info_cache = None

def parse_args():
    parser = argparse.ArgumentParser(description="桌面信息扫描工具")
    parser.add_argument("--max-files", type=int, default=MAX_FILES, help="最大扫描文件数")
    parser.add_argument("--max-depth", type=int, default=MAX_DEPTH, help="最大扫描深度")
    parser.add_argument("--recent-days", type=int, default=RECENT_DAYS, help="只扫描最近N天的文件")
    parser.add_argument("--file-types", type=str, default=",".join(TARGET_EXTENSIONS), help="目标文件扩展名，逗号分隔")
    return parser.parse_args()

def get_system_info():
    global _sys_info_cache
    if _sys_info_cache is not None:
        return _sys_info_cache
    try:
        _sys_info_cache = {
            'OS': platform.system(),
            'Version': platform.version(),
            'Release': platform.release(),
            'Machine': platform.machine(),
            'Processor': platform.processor(),
            'Hostname': platform.node(),
            'Architecture': platform.architecture()[0],
            'Python Version': platform.python_version()
        }
        try:
            _sys_info_cache['Screen Resolution'] = f"{win32api.GetSystemMetrics(0)}x{win32api.GetSystemMetrics(1)}"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop")
            wallpaper, _ = winreg.QueryValueEx(key, "WallPaper")
            _sys_info_cache['Desktop Wallpaper'] = wallpaper
            winreg.CloseKey(key)
        except Exception as e:
            _sys_info_cache['Desktop Info'] = f"无法获取: {e}"
        return _sys_info_cache
    except Exception as e:
        print(f"❌ 获取系统信息失败: {e}")
        return {}

def get_critical_paths():
    drives = []
    for letter in string.ascii_uppercase:
        drive = f"{letter}:\\"
        if os.path.exists(drive):
            try:
                os.listdir(drive)
                user_path = os.path.join(drive, "Users")
                if os.path.exists(user_path):
                    for user in os.listdir(user_path):
                        user_dir = os.path.join(user_path, user)
                        if os.path.isdir(user_dir):
                            desktop_path = os.path.join(user_dir, "Desktop")
                            if os.path.exists(desktop_path):
                                drives.append(desktop_path)
                                print(f"✅ 发现桌面目录: {desktop_path}")
            except PermissionError:
                print(f"⚠️  磁盘 {drive} 存在但无权限访问")
            except Exception as e:
                print(f"❌ 无法访问磁盘 {drive}: {e}")
    return drives

def get_lnk_target(file_path):
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(file_path)
        return shortcut.TargetPath
    except Exception as e:
        return f"无法解析: {e}"

def scan_and_write_paths(root_dirs, max_files=MAX_FILES, max_depth=MAX_DEPTH, recent_days=RECENT_DAYS, filename=INFO_TXT_FILENAME):
    count = 0
    cutoff_time = (datetime.now() - timedelta(days=recent_days)).timestamp()
    json_data = []
    with open(filename, 'a', encoding='utf-8') as f:
        for root_dir in root_dirs:
            print(f"\n🔍 正在扫描桌面路径: {root_dir}")
            try:
                for dirpath, dirnames, filenames in tqdm(os.walk(root_dir), desc="桌面扫描进度"):
                    dirnames[:] = [d for d in dirnames if d not in EXCLUDE_DIRS]
                    curr_depth = dirpath[len(root_dir):].count(os.sep)
                    if curr_depth >= max_depth:
                        dirnames[:] = []
                        continue
                    for name in dirnames:
                        dir_path = os.path.join(dirpath, name)
                        try:
                            stats = os.stat(dir_path)
                            if stats.st_mtime >= cutoff_time:
                                f.write(f"[DIR] {name}\n")
                                f.write(f"路径: {dir_path}\n")
                                f.write(f"创建时间: {datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write(f"修改时间: {datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write("-" * 60 + "\n")
                                f.flush()
                                json_data.append({
                                    'type': 'DIR',
                                    'name': name,
                                    'path': dir_path,
                                    'created': datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                                    'modified': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                                })
                                count += 1
                        except (PermissionError, OSError) as e:
                            if DEBUG:
                                print(f"  ⚠️ 跳过目录 {dir_path}: {e}")
                    for name in filenames:
                        if TARGET_EXTENSIONS and not any(name.lower().endswith(ext) for ext in TARGET_EXTENSIONS):
                            continue
                        file_path = os.path.join(dirpath, name)
                        try:
                            stats = os.stat(file_path)
                            if stats.st_mtime >= cutoff_time:
                                lnk_target = get_lnk_target(file_path) if name.lower().endswith('.lnk') else None
                                f.write(f"[FILE] {name}\n")
                                f.write(f"路径: {file_path}\n")
                                f.write(f"大小: {stats.st_size} 字节 ({stats.st_size / 1024 / 1024:.2f} MB)\n")
                                f.write(f"创建时间: {datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write(f"修改时间: {datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}\n")
                                if lnk_target:
                                    f.write(f"快捷方式目标: {lnk_target}\n")
                                f.write("-" * 60 + "\n")
                                f.flush()
                                json_entry = {
                                    'type': 'FILE',
                                    'name': name,
                                    'path': file_path,
                                    'size': stats.st_size,
                                    'created': datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                                    'modified': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                                }
                                if lnk_target:
                                    json_entry['lnk_target'] = lnk_target
                                json_data.append(json_entry)
                                count += 1
                        except (PermissionError, OSError) as e:
                            if DEBUG:
                                print(f"  ⚠️ 跳过文件 {file_path}: {e}")
                        if count >= max_files:
                            print(f"❗ 已达到最大文件数限制 ({max_files})，停止扫描。")
                            return count, json_data
            except PermissionError:
                print(f"❌ 无权限访问路径: {root_dir}")
            except Exception as e:
                print(f"❌ 扫描路径 {root_dir} 时出错: {e}")
    print(f"\n✅ 桌面扫描完成！共收集 {count} 个条目。")
    return count, json_data

def take_photos():
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("❌ 摄像头打开失败！请检查设备或权限。")
        return []
    start_time = datetime.now()
    photo_files = []
    photo_count = 0
    print("📸 开始拍照（每10秒一次，最多60秒）...")
    while (datetime.now() - start_time).total_seconds() <= MAX_PHOTO_TIME:
        ret, frame = cap.read()
        if ret:
            filename = f"{OUTPUT_IMAGE_PREFIX}_{photo_count}.jpg"
            cv2.imwrite(filename, frame)
            print(f"📸 照片已保存: {os.path.abspath(filename)}")
            photo_files.append(filename)
            photo_count += 1
        else:
            print(f"❌ 拍照失败（尝试 {photo_count + 1}）")
        time.sleep(PHOTO_INTERVAL)
        if (datetime.now() - start_time).total_seconds() > MAX_PHOTO_TIME:
            break
    cap.release()
    if not photo_files:
        print("❌ 未成功拍摄任何照片！")
    return photo_files

def save_scan_results(sys_info, count, json_data, filename=INFO_TXT_FILENAME, json_filename=JSON_OUTPUT_FILENAME):
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"📋 桌面信息扫描报告 - 时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            f.write("🖥️ 系统信息\n")
            for key, value in sys_info.items():
                f.write(f"{key}: {value}\n")
            f.write("-" * 60 + "\n\n")
            f.write(f"📁 桌面路径 - 共 {count} 个条目\n")
            f.write("=" * 80 + "\n\n")
        json_output = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'system_info': sys_info,
            'file_count': count,
            'desktop_items': sorted(json_data, key=lambda x: x.get('modified', ''), reverse=True)
        }
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(json_output, f, ensure_ascii=False, indent=2)
        print(f"📄 扫描信息已保存: {os.path.abspath(filename)}")
        print(f"📄 JSON 输出已保存: {os.path.abspath(json_filename)}")
        return True
    except Exception as e:
        print(f"❌ 保存扫描结果失败: {e}")
        return False

def compress_all(photo_files):
    if os.path.exists(COMPRESSED_ZIP_FILENAME):
        print(f"⚠️ 压缩包已存在，跳过压缩: {COMPRESSED_ZIP_FILENAME}")
        return True
    files_to_zip = [INFO_TXT_FILENAME, JSON_OUTPUT_FILENAME] + photo_files
    try:
        with ZipFile(COMPRESSED_ZIP_FILENAME, 'w', compression=ZipFile.ZIP_DEFLATED, compresslevel=9) as zipf:
            for file in files_to_zip:
                if os.path.exists(file):
                    zipf.write(file, arcname=os.path.basename(file))
                    print(f"📦 添加到压缩包: {file}")
                else:
                    print(f"⚠️ 文件不存在，跳过: {file}")
        print(f"✅ 压缩包已创建: {os.path.abspath(COMPRESSED_ZIP_FILENAME)}")
        return True
    except Exception as e:
        print(f"❌ 压缩失败: {e}")
        return False

def main():
    args = parse_args()
    global MAX_FILES, MAX_DEPTH, RECENT_DAYS, TARGET_EXTENSIONS
    MAX_FILES = args.max_files
    MAX_DEPTH = args.max_depth
    RECENT_DAYS = args.recent_days
    TARGET_EXTENSIONS = [ext.strip() for ext in args.file_types.split(",")]

    print("💾 欢迎使用 桌面信息扫描工具 v2.4")
    print("============================================================")

    print("\n🔍 正在收集系统信息...")
    sys_info = get_system_info()
    if sys_info:
        print("✅ 系统信息收集完成")
    else:
        print("❌ 系统信息收集失败")

    critical_paths = get_critical_paths()
    if not critical_paths:
        print("❌ 未发现任何桌面目录！")
        return

    if not save_scan_results(sys_info, 0, []):
        return

    print(f"\n🚀 开始扫描 {len(critical_paths)} 个桌面路径...")
    count, json_data = scan_and_write_paths(critical_paths, max_files=MAX_FILES, max_depth=MAX_DEPTH, recent_days=RECENT_DAYS, filename=INFO_TXT_FILENAME)
    if count == 0:
        print("❌ 扫描结果为空，可能是权限问题或路径为空。")
        return

    if not save_scan_results(sys_info, count, json_data):
        return

    photo_files = take_photos()
    if not photo_files:
        print("❌ 拍照失败，但仍继续打包...")

    compress_all(photo_files)

    print("\n🎉 所有任务完成！")
    print(f"📄 详细信息: {INFO_TXT_FILENAME}")
    print(f"📄 JSON 输出: {JSON_OUTPUT_FILENAME}")
    print(f"📸 照片: {', '.join(photo_files) if photo_files else '无'}")
    print(f"📦 压缩包: {COMPRESSED_ZIP_FILENAME}")
    print("👋 程序结束。")

if __name__ == "__main__":
    main()