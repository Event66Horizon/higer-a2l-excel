# 图形界面部分（tkinter）
import shutil
import re
import sys
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

def is_chinese(ch: str) -> bool:
    return '\u4e00' <= ch <= '\u9fff'

def make_safe_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def ensure_dir(path: Path):
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)

def get_files_to_process(src_root: Path):
    files = []
    for path in src_root.rglob('*'):
        if path.is_file() and path.name.lower().startswith('y'):
            files.append(path)
    return files

def process_files(src_root: Path, dst_root: Path, files, log_callback=None):
    ensure_dir(dst_root)
    for path in files:
        old_name = path.name
        new_name = old_name
        if new_name[0] != 'Y':
            new_name = 'Y' + new_name[1:]
            if log_callback:
                log_callback(f'首字母改成大写: {new_name}')
        first_chinese_pos = -1
        for idx, ch in enumerate(new_name):
            if is_chinese(ch):
                first_chinese_pos = idx
                break
        if first_chinese_pos != -1:
            english_part = new_name[:first_chinese_pos]
            english_part = re.sub(r'[^A-Za-z0-9]+$', '', english_part)
            chinese_part = new_name[first_chinese_pos:]
            new_name = english_part + '&' + chinese_part
            if log_callback:
                log_callback(f'英文和中文之间只保留一个&: {new_name}')
        if new_name != path.name:
            new_path = path.with_name(make_safe_filename(new_name))
            if new_path.exists():
                if log_callback:
                    log_callback(f'目标重命名文件已存在, 跳过: {new_path}')
                continue
            path.rename(new_path)
            path = new_path
            if log_callback:
                log_callback(f'重命名完成: {path}')
        dst_file = dst_root / path.name
        if dst_file.exists():
            if log_callback:
                log_callback(f'目标移动文件已存在, 跳过: {dst_file}')
            continue
        shutil.move(str(path), str(dst_file))
        if log_callback:
            log_callback(f'已移动到: {dst_file}')


class RenameApp:

    def __init__(self, root):
        self.root = root
        self.root.title('文件重命名工具')
        if getattr(sys, 'frozen', False):
            self.src_dir = Path(sys.executable).parent.resolve()
        else:
            self.src_dir = Path(__file__).parent.resolve()
        self.dst_dir = self.src_dir / 'output'
        self.dst_dir = self.src_dir / 'output'
        self.files = get_files_to_process(self.src_dir)
        self.invalid_files = self.get_invalid_files(self.files)
        self.label_files = tk.Label(root, text=f"检测到需要处理的文件：{len(self.files)} 个", font=("微软雅黑", 12))
        self.label_files.pack(pady=10)
        self.label_invalid = tk.Label(root, text=f"不符合规范的文件：{len(self.invalid_files)} 个", font=("微软雅黑", 12), fg='red')
        self.label_invalid.pack(pady=5)
        self.button = tk.Button(root, text="开始处理", font=("微软雅黑", 12), command=self.on_process)
        self.button.pack(pady=10)

    def get_invalid_files(self, files):
        invalid = []
        for path in files:
            if not self.is_valid_name(path.name):
                invalid.append(path)
        return invalid

    def is_valid_name(self, name: str):
        # 规范：Y开头，英文和中文之间只有一个&
        if not name.startswith('Y'):
            return False
        first_chinese_pos = -1
        for idx, ch in enumerate(name):
            if is_chinese(ch):
                first_chinese_pos = idx
                break
        if first_chinese_pos == -1:
            return True  # 没有中文部分
        english_part = name[:first_chinese_pos]
        chinese_part = name[first_chinese_pos:]
        return english_part.endswith('&') and not chinese_part.startswith('&')

    def on_process(self):
        if not self.files:
            messagebox.showinfo("提示", "没有需要处理的文件！")
            return
        process_files(self.src_dir, self.dst_dir, self.files)
        messagebox.showinfo("处理完成", f"全部处理完成！\n结果已输出到: {self.dst_dir}")
        # 处理完成后立即重新检测并刷新界面
        self.files = get_files_to_process(self.src_dir)
        self.invalid_files = self.get_invalid_files(self.files)
        self.label_files.config(text=f"检测到需要处理的文件：{len(self.files)} 个")
        self.label_invalid.config(text=f"不符合规范的文件：{len(self.invalid_files)} 个")

if __name__ == '__main__':
    root = tk.Tk()
    app = RenameApp(root)
    root.geometry('350x150')
    root.mainloop()