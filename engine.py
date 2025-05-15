"""
Voice Assistant (Python + PyQt5 版)
==================================
功能一览
• 语音识别（Google Web API，含实时监听 & 单次识别）
• 自然语言指令：搜索、浏览器标签控制、关闭前台程序等
• 自定义关键词映射（程序/网址/文件夹/文件） + 开始菜单一键导入
• 系统托盘图标（显示/隐藏窗口、开始/停止监听、退出）
• 全局热键（默认 F8，可在设置里修改，实时生效）
"""

import sys, os, json, webbrowser, difflib, threading, ctypes, psutil
import pyautogui, speech_recognition as sr, pythoncom, win32com.client
import win32gui, win32process, keyboard, pystray
from pystray import MenuItem as TrayItem
from PIL import Image, ImageDraw
from PyQt5 import QtCore, QtGui, QtWidgets

CONFIG_FILE = "command_map.json"


# ────────────────────────────────────────────────────────────────────────────────
# 实用线程：后台语音监听
# ────────────────────────────────────────────────────────────────────────────────
class SpeechThread(QtCore.QThread):
    recognized = QtCore.pyqtSignal(str)
    status_msg = QtCore.pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self._running = True
        self.recognizer = sr.Recognizer()
        self.mic = sr.Microphone()

    def run(self):
        with self.mic as source:
            self.recognizer.adjust_for_ambient_noise(source)
            while self._running:
                try:
                    self.status_msg.emit("监听中…")
                    audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
                    text = self.recognizer.recognize_google(audio, language="zh-CN")
                    self.recognized.emit(text)
                    self.status_msg.emit("识别成功")
                except sr.WaitTimeoutError:
                    continue
                except sr.UnknownValueError:
                    self.status_msg.emit("无法识别语音")
                except sr.RequestError:
                    self.status_msg.emit("识别服务出错")

    def stop(self):
        self._running = False


# ────────────────────────────────────────────────────────────────────────────────
# 设置/指令管理对话框
# ────────────────────────────────────────────────────────────────────────────────
class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, command_map: dict, current_hotkey: str, save_cb, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置 / 指令管理")
        self.resize(420, 480)
        self.command_map = command_map
        self.current_hotkey = current_hotkey
        self.save_cb = save_cb
        self.init_ui()

    # —— UI 组件 ────────────────────────────────────────────────────────────────
    def init_ui(self):
        main = QtWidgets.QVBoxLayout(self)

        # 指令列表
        self.list_widget = QtWidgets.QListWidget()
        main.addWidget(self.list_widget, stretch=1)

        # 按钮条
        btn_bar = QtWidgets.QHBoxLayout()
        for text, slot in [
            ("添加", self.add_cmd),
            ("删除选中", self.del_cmd),
            ("重命名选中", self.rename_cmd),
            ("从开始菜单导入", self.import_start_menu),
        ]:
            b = QtWidgets.QPushButton(text)
            b.clicked.connect(slot)
            btn_bar.addWidget(b)
        main.addLayout(btn_bar)

        # 热键设置
        hotkey_box = QtWidgets.QHBoxLayout()
        hotkey_box.addWidget(QtWidgets.QLabel("全局热键："))
        self.hotkey_edit = QtWidgets.QLineEdit(self.current_hotkey)
        hotkey_box.addWidget(self.hotkey_edit)
        main.addLayout(hotkey_box)

        # OK / Cancel
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok
                                          | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        main.addWidget(btns)

        self.refresh_list()

    # —— 列表刷新 ────────────────────────────────────────────────────────────
    def refresh_list(self):
        self.list_widget.clear()
        for k, v in self.command_map.items():
            if isinstance(v, dict):
                if "url" in v:
                    show = f"[网页] {k}  →  {v['url']}"
                elif "folder" in v:
                    show = f"[文件夹] {k}  →  {v['folder']}"
                else:
                    show = f"[文件] {k}  →  {v['file']}"
            else:
                show = f"[程序] {k}  →  {v}"
            self.list_widget.addItem(show)

    # —— 添加指令 ────────────────────────────────────────────────────────────
    def add_cmd(self):
        kw, ok = QtWidgets.QInputDialog.getText(self, "关键词", "输入关键词：")
        if not ok or not kw:
            return

        typ, ok = QtWidgets.QInputDialog.getItem(
            self, "类型", "选择类型：", ["程序", "网址", "文件夹", "文件"], 0, False
        )
        if not ok:
            return

        result = None
        if typ == "程序":
            path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择程序", filter="*.exe")
            if path:
                result = path
        elif typ == "网址":
            url, ok = QtWidgets.QInputDialog.getText(self, "网址", "http(s)://")
            if ok and url.startswith("http"):
                result = {"url": url}
        elif typ == "文件夹":
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "选择文件夹")
            if folder:
                result = {"folder": folder}
        else:
            f, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择文件")
            if f:
                result = {"file": f}

        if result:
            self.command_map[kw] = result
            self.save_cb()
            self.refresh_list()

    # —— 删除 / 重命名 ──────────────────────────────────────────────────────
    def del_cmd(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        key = item.text().split("  →")[0].split("] ")[-1]
        if key in self.command_map:
            del self.command_map[key]
            self.save_cb()
            self.refresh_list()

    def rename_cmd(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        old = item.text().split("  →")[0].split("] ")[-1]
        new, ok = QtWidgets.QInputDialog.getText(self, "重命名", f"将“{old}”改为：")
        if ok and new and new not in self.command_map:
            self.command_map[new] = self.command_map.pop(old)
            self.save_cb()
            self.refresh_list()

    # —— 开始菜单导入 ───────────────────────────────────────────────────────
    def import_start_menu(self):
        added = 0
        shell = win32com.client.Dispatch("WScript.Shell")
        for base in [
            os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"),
            r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
        ]:
            for root, _, files in os.walk(base):
                for f in files:
                    if f.endswith(".lnk"):
                        lnk = os.path.join(root, f)
                        try:
                            sc = shell.CreateShortcut(lnk)
                            target = sc.Targetpath
                            if target.lower().endswith(".exe"):
                                name = os.path.splitext(f)[0]
                                if name not in self.command_map:
                                    self.command_map[name] = target
                                    added += 1
                        except Exception:
                            pass
        if added:
            QtWidgets.QMessageBox.information(self, "成功", f"导入 {added} 个程序")
            self.save_cb()
            self.refresh_list()
        else:
            QtWidgets.QMessageBox.information(self, "提示", "未发现新程序")

    # —— 热键获取 ───────────────────────────────────────────────────────────
    def get_hotkey(self) -> str:
        return self.hotkey_edit.text().strip()


# ────────────────────────────────────────────────────────────────────────────────
# 主窗口（含托盘 & 热键）
# ────────────────────────────────────────────────────────────────────────────────
class VoiceAssistant(QtWidgets.QMainWindow):
    # 跨线程信号（托盘 & 热键 调用）
    sig_start = QtCore.pyqtSignal()
    sig_stop = QtCore.pyqtSignal()
    sig_show = QtCore.pyqtSignal()
    sig_exit = QtCore.pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("🎙️ 语音助手 (Python + Qt)")
        self.resize(700, 520)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
        self.setWindowOpacity(0.92)

        # 状态变量
        self.listening = False
        self.speech_thread: SpeechThread | None = None
        self.custom_cmds = self.load_cmds()

        # 热键配置
        self.settings = QtCore.QSettings("VACompany", "VoiceAssistant")
        self.current_hotkey = self.settings.value("hotkey", "F8")

        # UI
        self.init_ui()

        # 托盘
        self.tray = SystemTray(self)
        self.tray.start()

        # 信号槽
        self.sig_start.connect(self.start_listen)
        self.sig_stop.connect(self.stop_listen)
        self.sig_show.connect(self.show_window)
        self.sig_exit.connect(self.force_exit)

        # 全局热键
        keyboard.add_hotkey(self.current_hotkey, lambda: self.hotkey_toggle())

        self.speak("语音助手已启动（按下 {} 开始监听）".format(self.current_hotkey))

    # —— UI ────────────────────────────────────────────────────────────────────
    def init_ui(self):
        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        vbox = QtWidgets.QVBoxLayout(central)

        # 状态行
        state_box = QtWidgets.QHBoxLayout()
        self.state_label = QtWidgets.QLabel("状态：空闲")
        self.input_line = QtWidgets.QLineEdit()
        self.input_line.setReadOnly(True)
        state_box.addWidget(self.state_label)
        state_box.addStretch()
        state_box.addWidget(QtWidgets.QLabel("最近识别："))
        state_box.addWidget(self.input_line)
        vbox.addLayout(state_box)

        # 按钮区
        btn_box = QtWidgets.QHBoxLayout()
        self.btn_start = QtWidgets.QPushButton("🎧 开始监听")
        self.btn_start.clicked.connect(self.start_listen)
        self.btn_stop = QtWidgets.QPushButton("⏹️ 停止监听")
        self.btn_stop.clicked.connect(self.stop_listen)
        self.btn_stop.setEnabled(False)
        btn_speech = QtWidgets.QPushButton("🎤 语音转文字")
        btn_speech.clicked.connect(self.speech_once)
        btn_settings = QtWidgets.QPushButton("⚙️ 设置 / 指令管理")
        btn_settings.clicked.connect(self.open_settings)
        btn_box.addWidget(self.btn_start)
        btn_box.addWidget(self.btn_stop)
        btn_box.addWidget(btn_speech)
        btn_box.addWidget(btn_settings)
        vbox.addLayout(btn_box)

        # 输出框
        vbox.addWidget(QtWidgets.QLabel("输出："))
        self.result_box = QtWidgets.QTextEdit()
        self.result_box.setReadOnly(True)
        vbox.addWidget(self.result_box, stretch=1)

    # —— 说话输出 ────────────────────────────────────────────────────────────
    def speak(self, text: str):
        self.result_box.append(text)
        self.result_box.moveCursor(QtGui.QTextCursor.End)

    # —— 配置读写 ────────────────────────────────────────────────────────────
    def load_cmds(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def save_cmds(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.custom_cmds, f, ensure_ascii=False, indent=2)

    # —— Settings ────────────────────────────────────────────────────────────
    def open_settings(self):
        dlg = SettingsDialog(self.custom_cmds, self.current_hotkey, self.save_cmds, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            # 新热键
            new_hotkey = dlg.get_hotkey()
            if new_hotkey and new_hotkey != self.current_hotkey:
                try:
                    keyboard.remove_hotkey(self.current_hotkey)
                except Exception:
                    keyboard.unhook_all_hotkeys()
                keyboard.add_hotkey(new_hotkey, lambda: self.hotkey_toggle())
                self.current_hotkey = new_hotkey
                self.settings.setValue("hotkey", new_hotkey)
                self.settings.sync()
                self.speak(f"已将热键改为 {new_hotkey}")
            self.tray.update_menu()  # 指令变动也可能影响菜单
    # —— 热键切换（启动 ⇄ 停止） ——————————————————————————
    def hotkey_toggle(self):
        if self.listening:
            self.sig_stop.emit()      # 已在听 → 停止
        else:
            self.sig_start.emit()     # 未在听 → 开始
    # —— 监听线程控制 ───────────────────────────────────────────────────────
    def start_listen(self):
        # 供外部线程/热键调用 (信号)
        self.sig_start.emit()

    def stop_listen(self):
        self.sig_stop.emit()

    def start_listen_core(self):
        if self.listening:
            return
        self.listening = True
        self.state_label.setText("状态：监听中…")
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)

        # 后台线程
        self.speech_thread = SpeechThread()
        self.speech_thread.recognized.connect(self.on_recognized)
        self.speech_thread.status_msg.connect(self.state_label.setText)
        self.speech_thread.start()
        self.tray.update_menu()

    def stop_listen_core(self):
        if not self.listening:
            return
        self.listening = False
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.state_label.setText("状态：空闲")
        if self.speech_thread:
            self.speech_thread.stop()
            self.speech_thread.wait()
            self.speech_thread = None
        self.tray.update_menu()

    # 信号到槽
    def start_listen(self):  # noqa: override
        self.start_listen_core()

    def stop_listen(self):  # noqa: override
        self.stop_listen_core()

    # —— 单次语音转文字 ────────────────────────────────────────────────────
    def speech_once(self):
        rec = sr.Recognizer()
        with sr.Microphone() as src:
            self.state_label.setText("开始说话…")
            rec.adjust_for_ambient_noise(src)
            try:
                audio = rec.listen(src, timeout=5, phrase_time_limit=5)
                text = rec.recognize_google(audio, language="zh-CN")
                self.state_label.setText("识别完成")
                self.speak(f"📝 {text}")
                self.input_line.setText(text)
                self.handle_cmd(text)
            except sr.WaitTimeoutError:
                self.speak("⏰ 未检测到语音")
            except sr.UnknownValueError:
                self.speak("❓ 无法识别")
            except sr.RequestError:
                self.speak("⚠️ 服务出错")
            finally:
                self.state_label.setText("状态：空闲")

    # —— 识别回调 ───────────────────────────────────────────────────────────
    @QtCore.pyqtSlot(str)
    def on_recognized(self, txt):
        self.input_line.setText(txt)
        self.speak(f"📝 {txt}")
        self.handle_cmd(txt)

    # —— 托盘 & 窗口控制 ────────────────────────────────────────────────────
    def show_window(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def force_exit(self):
        self.tray.stop()
        QtWidgets.QApplication.quit()

    # —— Qt 关闭事件：隐藏到托盘 ────────────────────────────────────────────
    def closeEvent(self, e: QtGui.QCloseEvent):
        e.ignore()
        self.hide()
        self.speak("窗口已隐藏，可在托盘恢复")

    # ────────────────────────────────────────────────────────────────────
    # 指令处理
    # ────────────────────────────────────────────────────────────────────
    def ask_select(self, options: list[str]) -> str | None:
        item, ok = QtWidgets.QInputDialog.getItem(
            self, "选择指令", "未能精准识别，请选择：", options, 0, False
        )
        return item if ok else None

    def handle_cmd(self, cmd: str):
        lower = cmd.lower()

        # —— 关键字逻辑指令 ----------------------------------------------------
        if lower.startswith("搜索"):
            q = cmd[2:].strip()
            if q:
                self.speak(f"🔍 AI 搜索: {q}")
                webbrowser.open(f"https://www.perplexity.ai/search?q={q}")
            else:
                self.speak("请给出搜索内容")
            return

        if lower.startswith(("谷歌搜索", "查找")):
            q = cmd.replace("谷歌搜索", "").replace("查找", "").strip()
            if q:
                self.speak(f"🔍 Google: {q}")
                webbrowser.open(f"https://www.google.com/search?q={q}")
            else:
                self.speak("请给出搜索内容")
            return

        if "关闭当前程序" in lower or "关闭程序" in lower:
            self.close_foreground()
            return

        browser_map = {
            ("关闭标签页", "关闭网页"): ("ctrl", "w"),
            ("新建标签页", "打开新标签页"): ("ctrl", "t"),
            ("刷新网页", "刷新页面"): ("f5",),
            ("后退", "返回上一页"): ("alt", "left"),
            ("前进", "下一页"): ("alt", "right"),
        }
        for keys, hot in browser_map.items():
            if any(k in lower for k in keys):
                self.browser_action(*hot)
                self.speak(f"已执行浏览器操作")
                return

        # —— 自定义关键词匹配 ---------------------------------------------------
        key = self.find_best_match(lower)
        if not key:
            self.speak("❌ 未识别此指令")
            return

        target = self.custom_cmds[key]
        try:
            if isinstance(target, dict):
                if "url" in target:
                    self.speak(f"🌐 打开 {key}")
                    webbrowser.open(target["url"])
                elif "folder" in target:
                    self.speak(f"📂 打开文件夹 {key}")
                    os.startfile(target["folder"])
                else:
                    self.speak(f"📄 打开文件 {key}")
                    os.startfile(target["file"])
            else:
                self.speak(f"🚀 运行 {key}")
                ctypes.windll.shell32.ShellExecuteW(None, "runas", target, None, None, 1)
        except Exception as e:
            self.speak(f"⚠️ 执行失败: {e}")

    # —— 工具：最佳匹配 --------------------------------------------------------
    def find_best_match(self, cmd: str):
        for k in self.custom_cmds:
            if k.lower() in cmd:
                return k
        matches = difflib.get_close_matches(cmd, self.custom_cmds.keys(), n=3, cutoff=0.4)
        if matches:
            return self.ask_select(matches)
        return None

    # —— 浏览器操作 ------------------------------------------------------------
    def browser_action(self, *keys):
        if len(keys) == 1:
            pyautogui.press(keys[0])
        else:
            pyautogui.hotkey(*keys)

    # —— 关闭前台程序 ----------------------------------------------------------
    def close_foreground(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            for p in psutil.process_iter(['pid', 'name']):
                if p.info['pid'] == pid:
                    self.speak(f"关闭 {p.info['name']}")
                    p.terminate()
                    return
            self.speak("未找到前台程序")
        except Exception as e:
            self.speak(f"关闭失败: {e}")


# ────────────────────────────────────────────────────────────────────────────────
# 系统托盘封装
# ────────────────────────────────────────────────────────────────────────────────
class SystemTray:
    def __init__(self, win: VoiceAssistant):
        self.win = win
        img = Image.new("RGB", (64, 64), (0, 128, 0))
        d = ImageDraw.Draw(img); d.text((15, 20), "VA", fill=(255, 255, 255))
        self.icon = pystray.Icon("VA", img, "语音助手", menu=self.menu())

    def menu(self):
        def gen_menu():
            yield TrayItem("显示窗口", lambda: self.win.sig_show.emit())
            if self.win.listening:
                yield TrayItem("停止监听", lambda: self.win.sig_stop.emit())
            else:
                yield TrayItem("开始监听", lambda: self.win.sig_start.emit())
            yield TrayItem("退出", lambda: self.win.sig_exit.emit())
        return pystray.Menu(gen_menu)

    def start(self):
        threading.Thread(target=self.icon.run, daemon=True).start()

    def stop(self):
        try:
            self.icon.stop()
        except Exception:
            pass

    def update_menu(self):
        try:
            self.icon.update_menu()
        except Exception:
            pass


# ────────────────────────────────────────────────────────────────────────────────
# 程序入口
# ────────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    app.setQuitOnLastWindowClosed(False)

    va = VoiceAssistant()
    va.show()
    sys.exit(app.exec_())
