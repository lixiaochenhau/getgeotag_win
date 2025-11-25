import os
import sys
import subprocess
import csv
import datetime
import threading
import json
import tempfile
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext
from pathlib import Path

# 支持的文件扩展名配置
TRACK_EXTENSIONS = {'.gpx', '.kml', '.nmea', '.geo'}
PHOTO_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tiff', '.tif', '.cr2', '.nef', '.arw', '.heic', '.dng'}

class App:
    def __init__(self, root):
        """初始化主应用程序窗口及UI组件"""
        self.root = root
        self.root.title("Getgeotag")
        
        # 窗口尺寸及背景配置
        self.root.geometry("680x480")
        self.root.configure(bg="#ECECEC")

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.last_opened_dir = None

        # 主布局容器
        frame_main = tk.Frame(root, padx=15, pady=15, bg="#ECECEC")
        frame_main.pack(fill=tk.BOTH, expand=True)

        # 顶部操作说明标签
        self.lbl_instruction = tk.Label(
            frame_main,
            text="Select a single or parent folder containing photos and tracks to process.",
            font=('Helvetica', 12),
            fg="#555555",
            bg="#ECECEC",
            anchor="center"
        )
        self.lbl_instruction.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        # 文件夹选择按钮
        self.btn_select = tk.Button(
            frame_main, 
            text="Select Photo(s) and Track(s) Folder to Start", 
            command=self.select_folder, 
            height=1, 
            font=('Helvetica', 14, 'bold'),
            cursor="hand2", 
            highlightbackground="#ECECEC"
        )
        self.btn_select.pack(side=tk.TOP, fill=tk.X, pady=(0, 15))

        # 底部信息
        self.lbl_footer = tk.Label(
            frame_main,
            text="饮水机管理员 lixiaochen xiaochenensis@gmail.com",
            font=('Helvetica', 10),
            fg="#999999",
            bg="#ECECEC",
            anchor="w"
        )
        self.lbl_footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))

        # 日志显示区域
        self.text_log = scrolledtext.ScrolledText(
            frame_main, 
            state='disabled', 
            font=('Menlo', 12), 
            padx=10, 
            pady=10,
            bg="white",       
            fg="#333333",     
            relief=tk.FLAT,   
            borderwidth=0,
            highlightthickness=0 
        )
        self.text_log.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.log(">>> Application ready.")

        # 处理启动参数
        if len(sys.argv) > 1:
            path_arg = sys.argv[1]
            self.log(f"Argument detected: {path_arg}")
            threading.Thread(target=self.batch_process_entry, args=(path_arg,), daemon=True).start()

    def on_closing(self):
        """强制退出程序"""
        try:
            self.root.destroy()
        except:
            pass
        os._exit(0)

    def log(self, message):
        """线程安全的日志更新"""
        try:
            self.root.after(0, self._log_ui, message)
        except:
            pass

    def _log_ui(self, message):
        """更新日志UI组件"""
        try:
            self.text_log.config(state='normal')
            self.text_log.insert(tk.END, message + "\n")
            self.text_log.see(tk.END)
            self.text_log.config(state='disabled')
        except:
            pass

    def finish_task(self, summary=None):
        """任务结束处理"""
        if summary:
            self.log(summary)
        
        self.btn_select.config(state='normal', text="Select Photo(s) and Track(s) Folder to Start")
        self.log(">>> Waiting for next task...")

    def select_folder(self):
        """处理文件夹选择逻辑"""
        if self.last_opened_dir and os.path.isdir(self.last_opened_dir):
            start_dir = self.last_opened_dir
        else:
            start_dir = os.path.expanduser("~/Downloads")
            if not os.path.isdir(start_dir):
                start_dir = os.path.expanduser("~")

        folder_selected = filedialog.askdirectory(initialdir=start_dir)
        
        if folder_selected:
            self.last_opened_dir = folder_selected
            self.btn_select.config(state='disabled', text="Batch Processing...")
            threading.Thread(target=self.batch_process_entry, args=(folder_selected,), daemon=True).start()

    def get_resource_path(self, relative_path):
        """获取资源绝对路径，适配 PyInstaller"""
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    def run_exiftool(self, tool_path, args, file_list=None):
        """
        调用 ExifTool 子进程。
        使用临时文件传递参数以规避命令行长度限制。
        """
        temp_arg_file = None
        cmd = [tool_path] + args
        
        # 允许捕获标准输出和错误输出
        run_kwargs = {
            "capture_output": True,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "stdin": subprocess.DEVNULL
        }
        
        # Windows 平台隐藏控制台窗口
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            run_kwargs["startupinfo"] = startupinfo

        try:
            if file_list:
                fd, temp_arg_path = tempfile.mkstemp(text=True)
                with os.fdopen(fd, 'w', encoding='utf-8') as f:
                    for path in file_list:
                        f.write(path + "\n")
                cmd.extend(["-@", temp_arg_path])
                temp_arg_file = temp_arg_path

            process = subprocess.run(cmd, **run_kwargs)
            return process

        finally:
            if temp_arg_file and os.path.exists(temp_arg_file):
                try: os.remove(temp_arg_file)
                except: pass

    def scan_folder_content(self, target_dir):
        """扫描目录下的轨迹文件和照片文件"""
        track_files = []
        photo_files = []
        try:
            for file_path in target_dir.iterdir():
                if file_path.is_file() and not file_path.name.startswith('._'):
                    ext = file_path.suffix.lower()
                    if ext in TRACK_EXTENSIONS:
                        track_files.append(str(file_path.resolve()))
                    elif ext in PHOTO_EXTENSIONS:
                        photo_files.append(str(file_path.resolve()))
        except Exception:
            pass
        return track_files, photo_files

    def process_single_folder(self, target_dir, exiftool_path):
        """处理单个文件夹：写入GPS并生成CSV，包含详细错误日志"""
        track_files, photo_files = self.scan_folder_content(target_dir)

        if not track_files or not photo_files:
            self.log(f"  - Skipped: Missing tracks or photos.")
            return 0, None

        self.log(f"  - Found {len(track_files)} tracks, {len(photo_files)} photos.")

        # 执行 GPS 写入操作
        self.log(f"  - Writing GPS data...")
        # 移除 -q 参数以便在出错时能捕获更多信息，或者保留但捕获输出
        write_args = ['-overwrite_original', '-P']
        for track in track_files:
            write_args.extend(['-geotag', track])
        
        proc_write = self.run_exiftool(exiftool_path, write_args, list(photo_files))
        
        # 详细的错误处理逻辑
        if proc_write.returncode != 0:
            self.log(f"  - [Warning] ExifTool reported issues (Code {proc_write.returncode}):")
            # 打印 ExifTool 的具体报错信息，方便调试
            if proc_write.stdout:
                self.log(f"    [STDOUT]: {proc_write.stdout.strip()[:600]}") # 截断防止过长
            if proc_write.stderr:
                self.log(f"    [STDERR]: {proc_write.stderr.strip()[:600]}")
        else:
            # 即使 returncode 为 0，也可能没有更新图片（例如时间匹配不上）
            if "0 image files updated" in proc_write.stdout:
                self.log(f"  - [Warning] No images were updated. Possible time mismatch.")
                self.log(f"    Details: {proc_write.stdout.strip()[:300]}")

        # 生成 CSV 报告
        self.log(f"  - Generating CSV report...")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_csv = target_dir / f"photo_info_{timestamp}.csv"
        
        read_args = [
            '-j', '-q', '-charset', 'UTF8',
            '-FileName', '-GPSLatitude', '-n', 
            '-GPSLongitude', '-n', '-GPSAltitude', '-n', 
            '-DateTimeOriginal'
        ]
        
        proc_read = self.run_exiftool(exiftool_path, read_args, list(photo_files))
        
        count = 0
        try:
            meta_data = json.loads(proc_read.stdout)
            if meta_data:
                with open(output_csv, 'w', newline='', encoding='gbk') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Path", "Name", "Lat", "Lon", "Alt", "Date", "Time"])
                    for item in meta_data:
                        dt = item.get('DateTimeOriginal', '')
                        d_val, t_val = "", ""
                        if dt and " " in dt:
                            parts = dt.split(" ", 1)
                            d_val = parts[0].replace(":", "-")
                            t_val = parts[1]

                        writer.writerow([
                            item.get('SourceFile', ''),
                            item.get('FileName', ''),
                            item.get('GPSLatitude', ''),
                            item.get('GPSLongitude', ''),
                            item.get('GPSAltitude', ''),
                            d_val, t_val
                        ])
                        count += 1
                self.log(f"  - [Success] Generated: {output_csv.name}")
                return count, output_csv.name
        except:
            self.log(f"  - [Error] Failed to parse metadata.")
            if proc_read.stderr:
                self.log(f"    Debug: {proc_read.stderr.strip()}")
        
        return 0, None

    def batch_process_entry(self, root_dir_str):
        """批量处理入口：分析目录结构并顺序执行"""
        start_time = time.time()
        
        try:
            self.log("=" * 40)
            self.log("Analyzing structure...")
            
            root_dir = Path(root_dir_str).resolve()
            
            # 配置 ExifTool 路径
            exiftool_path = self.get_resource_path("exiftool.exe")
            
            # 本地调试时的回退机制
            if not os.path.exists(exiftool_path):
                # 尝试查找同目录下的 exiftool
                local_tool = os.path.join(root_dir, "exiftool.exe")
                if os.path.exists(local_tool):
                    exiftool_path = local_tool
                elif os.path.exists("exiftool.exe"):
                    exiftool_path = os.path.abspath("exiftool.exe")
                else:
                    self.log("Error: ExifTool binary not found.")
                    self.log(f"Searched at: {exiftool_path}")
                    self.root.after(0, self.finish_task)
                    return

            # 识别待处理文件夹
            task_folders = []
            
            # 1. 检查根目录
            t_files, p_files = self.scan_folder_content(root_dir)
            if t_files and p_files:
                task_folders.append(root_dir)
            
            # 2. 检查子目录 (深度为1)
            subdirs = sorted([p for p in root_dir.iterdir() if p.is_dir()], key=lambda p: p.name)
            for sub in subdirs:
                t_sub, p_sub = self.scan_folder_content(sub)
                if t_sub and p_sub:
                    task_folders.append(sub)
            
            # 去重
            seen = set()
            unique_tasks = []
            for f in task_folders:
                if str(f) not in seen:
                    unique_tasks.append(f)
                    seen.add(str(f))
            task_folders = unique_tasks

            if not task_folders:
                self.log("No folders found containing BOTH photos and tracks.")
                self.root.after(0, self.finish_task)
                return

            # 列出任务队列
            total_tasks = len(task_folders)
            self.log(f"Queue: {total_tasks} folder(s) found.")
            for i, folder in enumerate(task_folders, 1):
                self.log(f"  {i}. {folder.name}")
            
            self.log("-" * 40)

            # 顺序执行处理循环
            total_processed_photos = 0
            completed_folders = 0
            
            for index, folder in enumerate(task_folders, 1):
                self.log(f"Processing [{index}/{total_tasks}]: {folder.name}")
                
                p_count, _ = self.process_single_folder(folder, exiftool_path)
                
                if p_count > 0:
                    total_processed_photos += p_count
                    completed_folders += 1
                
                self.log("-" * 40)

            # 计算总耗时
            end_time = time.time()
            duration = end_time - start_time
            minutes = int(duration // 60)
            seconds = int(duration % 60)

            # 输出最终汇总报告
            summary = (
                f"\n[BATCH COMPLETED]\n"
                f"Folders: {completed_folders} / {total_tasks}\n"
                f"Photos:  {total_processed_photos}\n"
                f"Time:    {minutes}m {seconds}s\n"
                f"{'='*40}"
            )
            self.root.after(0, lambda: self.finish_task(summary))

        except Exception as e:
            self.log(f"Critical Error: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.root.after(0, self.finish_task)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()