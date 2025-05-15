import speech_recognition as sr
import os
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, simpledialog
import json
import webbrowser
import pythoncom
import win32com.client
import difflib

CONFIG_FILE = "command_map.json"

# -----------------------------
# 初始化语音合成

def speak(text):
    result_box.insert(tk.END, f"{text}\n")
    result_box.see(tk.END)

# -----------------------------
# 配置读取与保存

def load_commands():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_commands():
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(custom_commands, f, ensure_ascii=False, indent=2)

custom_commands = load_commands()

# -----------------------------
# 语音转文字

def speech_to_text_once():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    try:
        with mic as source:
            status_var.set("开始语音识别...")
            result_box.insert(tk.END, "🎤 说话中...\n")
            result_box.see(tk.END)
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
            text = recognizer.recognize_google(audio, language='zh-CN')
            result_box.insert(tk.END, f"📝 识别结果: {text}\n")
            result_box.see(tk.END)
            status_var.set("识别完成")
    except sr.WaitTimeoutError:
        result_box.insert(tk.END, "⏰ 没有检测到语音\n")
    except sr.UnknownValueError:
        result_box.insert(tk.END, "❓ 无法识别内容\n")
    except sr.RequestError:
        result_box.insert(tk.END, "⚠️ 语音识别服务出错\n")

# -----------------------------
# 指令执行器

def ask_user_to_select(options):
    def on_select():
        selected = listbox.get(tk.ACTIVE)
        dialog.destroy()
        nonlocal choice
        choice = selected

    choice = None
    dialog = tk.Toplevel(window)
    dialog.title("请选择要执行的命令")
    dialog.geometry("300x200")
    tk.Label(dialog, text="未能精准识别指令，请选择匹配项：").pack(pady=5)

    listbox = tk.Listbox(dialog)
    for opt in options:
        listbox.insert(tk.END, opt)
    listbox.pack(pady=10)

    tk.Button(dialog, text="确定", command=on_select).pack(pady=5)

    dialog.transient(window)
    dialog.grab_set()
    window.wait_window(dialog)

    return choice

def handle_command(command):
    def find_best_match(cmd):
        cmd = cmd.lower()
        for keyword in custom_commands:
            key_lower = keyword.lower()
            if key_lower in cmd or f"我的{key_lower}" in cmd or cmd.strip() == key_lower:
                return keyword

        possible_matches = difflib.get_close_matches(cmd, custom_commands.keys(), n=3, cutoff=0.4)
        if possible_matches:
            return ask_user_to_select(possible_matches)
        return None

    keyword = find_best_match(command)
    if not keyword:
        speak("我还不懂这个指令")
        return

    target = custom_commands[keyword]
    try:
        if isinstance(target, dict):
            if "url" in target:
                speak(f"正在打开 {keyword}")
                webbrowser.open(target["url"])
            elif "folder" in target:
                speak(f"正在打开 {keyword} 文件夹")
                os.startfile(target["folder"])
            elif "file" in target:
                speak(f"正在打开文件 {keyword}")
                os.startfile(target["file"])
        elif isinstance(target, str):
            speak(f"正在打开 {keyword}")
            subprocess.Popen([target])
        else:
            speak("指令匹配失败")
    except Exception as e:
        speak(f"执行失败：{str(e)}")

# -----------------------------
# 语音识别监听

def recognize_speech_background():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)

    def listen_loop():
        with mic as source:
            while listening:
                try:
                    status_var.set("监听中...")
                    audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
                    text = recognizer.recognize_google(audio, language='zh-CN')
                    command_var.set(text)
                    status_var.set("识别成功")
                    handle_command(text)
                except sr.WaitTimeoutError:
                    continue
                except sr.UnknownValueError:
                    status_var.set("无法识别语音")
                except sr.RequestError:
                    status_var.set("识别服务出错")

    threading.Thread(target=listen_loop, daemon=True).start()

def start_listening():
    global listening
    listening = True
    speak("开始监听语音指令")
    recognize_speech_background()
    listen_btn.config(state=tk.DISABLED)
    stop_btn.config(state=tk.NORMAL)

def stop_listening():
    global listening
    listening = False
    status_var.set("监听已停止")
    speak("已停止监听")
    listen_btn.config(state=tk.NORMAL)
    stop_btn.config(state=tk.DISABLED)

# -----------------------------
# 设置界面

def open_settings():
    settings = tk.Toplevel(window)
    settings.title("设置 - 添加/删除/重命名命令")
    settings.geometry("600x600")
    settings.minsize(600, 600)
    settings.configure(bg="#f8f9fa")

    def adjust_transparency():
        val = simpledialog.askfloat("设置透明度", "请输入透明度（0.1 - 1.0）:", minvalue=0.1, maxvalue=1.0)
        if val:
            window.attributes("-alpha", val)

    def toggle_topmost():
        current = window.attributes("-topmost")
        window.attributes("-topmost", not current)
        status = "已置顶" if not current else "取消置顶"
        speak(status)

    # 输入区域
    tk.Label(settings, text="关键词:", bg="#f8f9fa").pack()
    keyword_entry = tk.Entry(settings, width=30)
    keyword_entry.pack(pady=3)

    exe_path_var = tk.StringVar()
    url_var = tk.StringVar()
    folder_var = tk.StringVar()
    file_var = tk.StringVar()

    tk.Label(settings, text="程序路径（可选）:", bg="#f8f9fa").pack()
    tk.Entry(settings, textvariable=exe_path_var, width=35).pack()
    tk.Button(settings, text="选择程序", command=lambda: exe_path_var.set(filedialog.askopenfilename(filetypes=[("可执行程序", "*.exe")]))).pack(pady=2)

    tk.Label(settings, text="网址链接（可选）:", bg="#f8f9fa").pack()
    tk.Entry(settings, textvariable=url_var, width=35).pack()

    tk.Label(settings, text="文件夹路径（可选）:", bg="#f8f9fa").pack()
    tk.Entry(settings, textvariable=folder_var, width=35).pack()
    tk.Button(settings, text="选择文件夹", command=lambda: folder_var.set(filedialog.askdirectory())).pack(pady=2)

    tk.Label(settings, text="文件路径（可选）:", bg="#f8f9fa").pack()
    tk.Entry(settings, textvariable=file_var, width=35).pack()
    tk.Button(settings, text="选择文件", command=lambda: file_var.set(filedialog.askopenfilename())).pack(pady=2)

    def save_mapping():
        keyword = keyword_entry.get().strip()
        exe = exe_path_var.get().strip()
        url = url_var.get().strip()
        folder = folder_var.get().strip()
        file_path = file_var.get().strip()

        if not keyword:
            return messagebox.showerror("错误", "关键词不能为空")

        if exe and os.path.isfile(exe):
            custom_commands[keyword] = exe
        elif url.startswith("http"):
            custom_commands[keyword] = {"url": url}
        elif os.path.isdir(folder):
            custom_commands[keyword] = {"folder": folder}
        elif os.path.isfile(file_path):
            custom_commands[keyword] = {"file": file_path}
        else:
            return messagebox.showerror("错误", "请输入有效的路径或网址")

        save_commands()
        update_command_list()
        speak(f"已添加 {keyword}")

    tk.Button(settings, text="保存关键词", command=save_mapping).pack(pady=5)

    listbox = tk.Listbox(settings, height=10)
    listbox.pack(fill=tk.BOTH, expand=True, padx=10)

    def update_command_list():
        listbox.delete(0, tk.END)
        for key, val in custom_commands.items():
            if isinstance(val, dict):
                if "url" in val:
                    listbox.insert(tk.END, f"[网页] {key} => {val['url']}")
                elif "folder" in val:
                    listbox.insert(tk.END, f"[文件夹] {key} => {val['folder']}")
                elif "file" in val:
                    listbox.insert(tk.END, f"[文件] {key} => {val['file']}")
            else:
                listbox.insert(tk.END, f"[程序] {key} => {val}")

    def delete_selected():
        selection = listbox.curselection()
        if selection:
            key = listbox.get(selection[0]).split(" => ")[0].split("] ")[1]
            if key in custom_commands:
                del custom_commands[key]
                save_commands()
                update_command_list()
                speak(f"已删除 {key}")

    def rename_selected():
        selection = listbox.curselection()
        if selection:
            old_key = listbox.get(selection[0]).split(" => ")[0].split("] ")[1]
            new_key = simpledialog.askstring("重命名", f"将 \"{old_key}\" 重命名为:")
            if new_key and new_key not in custom_commands:
                custom_commands[new_key] = custom_commands.pop(old_key)
                save_commands()
                update_command_list()
                speak(f"已重命名为 {new_key}")
            else:
                messagebox.showerror("错误", "无效或重复的名称")

    tk.Button(settings, text="删除选中", command=delete_selected).pack(pady=3)
    tk.Button(settings, text="重命名选中", command=rename_selected).pack(pady=3)
    tk.Button(settings, text="设置透明度", command=adjust_transparency).pack(pady=2)
    tk.Button(settings, text="切换置顶", command=toggle_topmost).pack(pady=2)
    tk.Button(settings, text="语音转文字", command=speech_to_text_once).pack(pady=5)

    def scan_start_menu_programs():
        start_menu_paths = [
            os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"),
            r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
        ]
        shell = win32com.client.Dispatch("WScript.Shell")
        added = 0

        for menu_path in start_menu_paths:
            for root, dirs, files in os.walk(menu_path):
                for file in files:
                    if file.endswith(".lnk"):
                        lnk_path = os.path.join(root, file)
                        try:
                            shortcut = shell.CreateShortcut(lnk_path)
                            target = shortcut.Targetpath
                            if target and target.lower().endswith(".exe"):
                                keyword = os.path.splitext(file)[0]
                                if keyword not in custom_commands:
                                    custom_commands[keyword] = target
                                    added += 1
                        except Exception:
                            continue

        if added:
            save_commands()
            update_command_list()
            speak(f"已导入开始菜单中的 {added} 个程序")
        else:
            speak("没有发现新的程序可以添加")

    tk.Button(settings, text="从开始菜单导入程序", command=scan_start_menu_programs).pack(pady=10)
    update_command_list()

# -----------------------------
# 主窗口界面优化

window = tk.Tk()
window.title("🎙️ 语音助手")
window.geometry("600x520")
window.configure(bg="#f0f2f5")
window.attributes('-alpha', 0.3)
window.attributes('-topmost', True)

command_var = tk.StringVar()
status_var = tk.StringVar()

# 标题
tk.Label(window, text="语音助手", font=("微软雅黑", 20, "bold"), bg="#f0f2f5", fg="#333").pack(pady=10)

# 状态区域
status_frame = tk.Frame(window, bg="#ffffff", relief="groove", bd=1)
status_frame.pack(pady=5, padx=10, fill=tk.X)

tk.Label(status_frame, text="当前语音识别内容：", font=("微软雅黑", 12), bg="#ffffff").pack(side=tk.LEFT, padx=5)
tk.Entry(status_frame, textvariable=command_var, font=("微软雅黑", 12), width=35, state="readonly").pack(side=tk.LEFT, padx=5)
tk.Label(status_frame, textvariable=status_var, font=("微软雅黑", 12), fg="blue", bg="#ffffff").pack(side=tk.RIGHT, padx=5)

# 按钮区域
button_frame = tk.Frame(window, bg="#f0f2f5")
button_frame.pack(pady=10)

listen_btn = tk.Button(button_frame, text="🎧 开始监听", font=("微软雅黑", 12), command=start_listening, bg="#4CAF50", fg="white", width=12)
listen_btn.grid(row=0, column=0, padx=10, pady=5)

stop_btn = tk.Button(button_frame, text="⏹️ 停止监听", font=("微软雅黑", 12), command=stop_listening, bg="#f44336", fg="white", width=12, state=tk.DISABLED)
stop_btn.grid(row=0, column=1, padx=10, pady=5)

speech_btn = tk.Button(button_frame, text="🎤 语音转文字", font=("微软雅黑", 12), command=speech_to_text_once, bg="#2196F3", fg="white", width=12)
speech_btn.grid(row=1, column=0, padx=10, pady=5)

settings_btn = tk.Button(button_frame, text="⚙️ 打开设置", font=("微软雅黑", 12), command=open_settings, width=12)
settings_btn.grid(row=1, column=1, padx=10, pady=5)

# 输出框
tk.Label(window, text="识别结果输出：", font=("微软雅黑", 12), bg="#f0f2f5").pack()
result_box = scrolledtext.ScrolledText(window, width=70, height=10, font=("微软雅黑", 10))
result_box.pack(padx=10, pady=10)

# 退出按钮
tk.Button(window, text="退出程序", command=window.quit, font=("微软雅黑", 11), bg="#888", fg="white").pack(pady=5)

speak("语音助手已启动")
listening = False
window.mainloop()