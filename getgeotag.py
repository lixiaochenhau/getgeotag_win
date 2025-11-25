import os
import sys
import subprocess
import csv
import datetime
import threading
import json
import tempfile
import time
import re
import tkinter as tk
from tkinter import filedialog, scrolledtext
from pathlib import Path

# 支持的文件扩展名配置
TRACK_EXTENSIONS = {'.gpx', '.kml', '.nmea', '.geo'}
PHOTO_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tiff', '.tif', '.cr2', '.nef', '.arw', '.heic', '.dng'}

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Getgeotag")
        
        self.root.geometry("800x600")
        self.root.configure(bg="#ECECEC")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.last_opened_dir = None

        frame_main = tk.Frame(root, padx=15, pady=15, bg="#ECECEC")
        frame_main.pack(fill=tk.BOTH, expand=True)

        self.lbl_instruction = tk.Label(
            frame_main,
            text="Select a folder containing photos and tracks (External Drives Supported).",
            font=('Helvetica', 12), fg="#555555", bg="#ECECEC", anchor="center"
        )
        self.lbl_instruction.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        self.btn_select = tk.Button(
            frame_main, 
            text="Select Photo(s) and Track(s) Folder to Start", 
            command=self.select_folder, 
            height=1, font=('Helvetica', 14, 'bold'), cursor="hand2", highlightbackground="#ECECEC"
        )
        self.btn_select.pack(side=tk.TOP, fill=tk.X, pady=(0, 15))

        self.lbl_footer = tk.Label(
            frame_main,
            text="饮水机管理员 lixiaochen xiaochenensis@gmail.com",
            font=('Helvetica', 10), fg="#999999", bg="#ECECEC", anchor="w"
        )
        self.lbl_footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))

        self.text_log = scrolledtext.ScrolledText(
            frame_main, state='disabled', font=('Menlo', 11), 
            padx=10, pady=10, bg="white", fg="#333333", relief=tk.FLAT, borderwidth=0
        )
        self.text_log.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.log(">>> Application ready.")

        if len(sys.argv) > 1:
            path_arg = sys.argv[1]
            self.log(f"Argument detected: {path_arg}")
            threading.Thread(target=self.batch_process_entry, args=(path_arg,), daemon=True).start()

    def on_closing(self):
        try: self.root.destroy()
        except: pass
        os._exit(0)

    def log(self, message):
        try: self.root.after(0, self._log_ui, message)
        except: pass

    def _log_ui(self, message):
        try:
            self.text_log.config(state='normal')
            self.text_log.insert(tk.END, message + "\n")
            self.text_log.see(tk.END)
            self.text_log.config(state='disabled')
        except: pass

    def finish_task(self, summary=None):
        if summary: self.log(summary)
        self.btn_select.config(state='normal', text="Select Photo(s) and Track(s) Folder to Start")
        self.log(">>> Waiting for next task...")

    def select_folder(self):
        if self.last_opened_dir and os.path.isdir(self.last_opened_dir):
            start_dir = self.last_opened_dir
        else:
            start_dir = os.path.expanduser("~/Downloads")
            if not os.path.isdir(start_dir): start_dir = os.path.expanduser("~")

        folder_selected = filedialog.askdirectory(initialdir=start_dir)
        if folder_selected:
            self.last_opened_dir = folder_selected
            self.btn_select.config(state='disabled', text="Batch Processing...")
            threading.Thread(target=self.batch_process_entry, args=(folder_selected,), daemon=True).start()

    def get_resource_path(self, relative_path):
        if hasattr(sys, '_MEIPASS'): return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    def run_exiftool(self, tool_path, args, file_list=None):
        temp_arg_file = None
        cmd = [tool_path] + args
        run_kwargs = {"capture_output": True, "text": True, "encoding": "utf-8", "errors": "replace", "stdin": subprocess.DEVNULL}
        
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            run_kwargs["startupinfo"] = startupinfo

        try:
            if file_list:
                fd, temp_arg_path = tempfile.mkstemp(text=True)
                with os.fdopen(fd, 'w', encoding='utf-8') as f:
                    for path in file_list:
                        # Normalize path separators
                        clean_path = str(path).replace("\\", "/")
                        f.write(clean_path + "\n")
                cmd.extend(["-@", temp_arg_path])
                temp_arg_file = temp_arg_path

            process = subprocess.run(cmd, **run_kwargs)
            return process

        finally:
            if temp_arg_file and os.path.exists(temp_arg_file):
                try: os.remove(temp_arg_file)
                except: pass

    def clean_log_output(self, output_str):
        if not output_str: return ""
        lines = output_str.split('\n')
        clean_lines = []
        for line in lines:
            if line.count('') > 3 or line.count('◆') > 3: continue
            if line.strip(): clean_lines.append(line.strip())
        return "\n".join(clean_lines)

    def process_single_folder(self, target_dir, exiftool_path):
        track_files, photo_files = self.scan_folder_content(target_dir)

        if not track_files or not photo_files:
            self.log(f"  - Skipped: Missing tracks or photos.")
            return 0, None

        self.log(f"  - Found {len(track_files)} tracks, {len(photo_files)} photos.")

        # =======================================================
        # 1. Write GPS Data
        # =======================================================
        self.log(f"  - Attempting to write GPS data...")
        
        write_args = [
            '-overwrite_original', 
            '-P',
            '-charset', 'filename=UTF8',
            
            # Set timezone offset to UTC+8
            '-Geotime<${DateTimeOriginal}+08:00',
            
            # Set time tolerance to 120 seconds (2 minutes)
            '-api', 'GeoMaxIntSecs=120'
        ]
        
        for track in track_files:
            track_clean = str(track).replace("\\", "/")
            write_args.extend(['-geotag', track_clean])
        
        proc_write = self.run_exiftool(exiftool_path, write_args, list(photo_files))
        
        files_updated = 0
        if proc_write.returncode != 0:
            self.log(f"  - [Error] ExifTool failed (Code {proc_write.returncode}).")
            clean_stderr = self.clean_log_output(proc_write.stderr)
            if clean_stderr: self.log(f"    Details: {clean_stderr[:500]}")
        else:
            match = re.search(r'(\d+)\s+image files updated', proc_write.stdout)
            if match: files_updated = int(match.group(1))
            
            if files_updated == 0:
                self.log("  - [Warning] 0 images were updated.")
                self.log("    (Cause: No GPS points found within 2 minutes of photo time).")
            else:
                self.log(f"  - [Success] Successfully geotagged {files_updated} images.")

        # =======================================================
        # 2. Generate CSV Report
        # =======================================================
        
        if files_updated > 0:
            self.log(f"  - Generating CSV report (Based on NEWLY WRITTEN data)...")
        else:
            self.log(f"  - [Notice] Generating CSV report based on PRE-EXISTING metadata.")

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_csv = target_dir / f"photo_info_{timestamp}.csv"
        
        read_args = [
            '-j', '-q', 
            '-charset', 'UTF8',           
            '-charset', 'filename=UTF8',
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
                self.log(f"  - [Report] Saved to: {output_csv.name}")
                return count, output_csv.name
        except json.JSONDecodeError:
            self.log(f"  - [Error] Failed to parse metadata JSON.")
        except Exception as e:
             self.log(f"  - [Error] CSV Generation failed: {e}")
        
        return 0, None

    def scan_folder_content(self, target_dir):
        track_files = []
        photo_files = []
        try:
            for file_path in target_dir.iterdir():
                if file_path.is_file() and not file_path.name.startswith('._'):
                    ext = file_path.suffix.lower()
                    if ext in TRACK_EXTENSIONS: track_files.append(str(file_path.resolve()))
                    elif ext in PHOTO_EXTENSIONS: photo_files.append(str(file_path.resolve()))
        except Exception: pass
        return track_files, photo_files

    def batch_process_entry(self, root_dir_str):
        start_time = time.time()
        try:
            self.log("=" * 40)
            self.log("Analyzing structure...")
            root_dir = Path(root_dir_str).resolve()
            exiftool_path = self.get_resource_path("exiftool.exe")
            
            if not os.path.exists(exiftool_path):
                if os.path.exists("exiftool.exe"): exiftool_path = os.path.abspath("exiftool.exe")
                else:
                    self.log("Error: ExifTool binary not found.")
                    self.root.after(0, self.finish_task)
                    return

            task_folders = []
            t_files, p_files = self.scan_folder_content(root_dir)
            if t_files and p_files: task_folders.append(root_dir)
            
            subdirs = sorted([p for p in root_dir.iterdir() if p.is_dir()], key=lambda p: p.name)
            for sub in subdirs:
                t_sub, p_sub = self.scan_folder_content(sub)
                if t_sub and p_sub: task_folders.append(sub)
            
            seen = set(); unique_tasks = []
            for f in task_folders:
                if str(f) not in seen: unique_tasks.append(f); seen.add(str(f))
            task_folders = unique_tasks

            if not task_folders:
                self.log("No folders found containing BOTH photos and tracks.")
                self.root.after(0, self.finish_task)
                return

            total_tasks = len(task_folders)
            self.log(f"Queue: {total_tasks} folder(s) found.")
            for i, folder in enumerate(task_folders, 1): self.log(f"  {i}. {folder.name}")
            self.log("-" * 40)

            total_processed_photos = 0; completed_folders = 0
            for index, folder in enumerate(task_folders, 1):
                self.log(f"Processing [{index}/{total_tasks}]: {folder.name}")
                p_count, _ = self.process_single_folder(folder, exiftool_path)
                if p_count > 0: total_processed_photos += p_count; completed_folders += 1
                self.log("-" * 40)

            end_time = time.time()
            duration = end_time - start_time
            minutes = int(duration // 60); seconds = int(duration % 60)

            summary = (f"\n[BATCH COMPLETED]\nFolders: {completed_folders} / {total_tasks}\n"
                       f"Photos:  {total_processed_photos}\nTime:    {minutes}m {seconds}s\n{'='*40}")
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