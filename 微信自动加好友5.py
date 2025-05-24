import sys
import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
import pyautogui
import time
import uiautomation as auto
import subprocess
import random
from datetime import datetime
import os

# 强制导入 comtypes.stream，防止打包时被遗漏
import comtypes
import comtypes.stream


# 日志目录
LOG_DIR = "logs"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# 日志文件名（按运行时间生成）
log_file_path = os.path.join(LOG_DIR, f"wechat_automation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")


# 写入日志函数
def log_info(message, log_widget=None):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = f"[{timestamp}] {message}"
    print(msg)
    with open(log_file_path, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    if log_widget:
        log_widget.config(state='normal')
        log_widget.insert(tk.END, msg + "\n")
        log_widget.config(state='disabled')
        log_widget.see(tk.END)


# 图像识别点击函数
def click_image(img_path, confidence=0.8, retries=3, delay=2):
    for i in range(retries):
        location = pyautogui.locateOnScreen(img_path)
        if location:
            center = pyautogui.center(location)
            human_like_click(center.x, center.y)
            return True
        time.sleep(delay)
    return False


# 模拟人类鼠标移动轨迹
def human_like_move_to(x, y):
    current_x, current_y = pyautogui.position()
    steps = random.choice([2, 3, 4])
    for i in range(1, steps + 1):
        ratio = i / steps
        intermediate_x = int(current_x * (1 - ratio) + x * ratio)
        intermediate_y = int(current_y * (1 - ratio) + y * ratio)
        jitter_x = random.randint(-3, 3)
        jitter_y = random.randint(-3, 3)
        pyautogui.moveTo(intermediate_x + jitter_x, intermediate_y + jitter_y)
        time.sleep(random.uniform(0.1, 0.3))


# 模拟人类点击动作
def human_like_click(x=None, y=None):
    current_x, current_y = pyautogui.position()
    target_x = x if x is not None else current_x
    target_y = y if y is not None else current_y

    human_like_move_to(target_x, target_y)
    time.sleep(random.uniform(0.2, 1))

    final_x = target_x + random.randint(-8, 8)
    final_y = target_y + random.randint(-8, 8)
    pyautogui.click(final_x, final_y)


# 统一控件点击接口
def click_control(control):
    rect = control.BoundingRectangle
    center_x = (rect.left + rect.right) // 2
    center_y = (rect.top + rect.bottom) // 2
    human_like_click(center_x, center_y)


# 自动化主流程
def run_automation(config, log_widget, update_countdown_callback, start_button):
    from uiautomation import UIAutomationInitializerInThread
    with UIAutomationInitializerInThread():
        wechat_path = config['wechat_path']
        accounts = config['accounts']
        names = config['names']
        verify_message = config['verify_message']
        min_interval = config['min_interval']
        max_interval = config['max_interval']

    # 禁用开始按钮
    start_button.config(state=tk.DISABLED)

    wechat_process = subprocess.Popen([wechat_path])
    log_info("启动微信客户端...", log_widget)
    time.sleep(5)

    try:
        wechat_window = auto.WindowControl(
            searchDepth=1,
            ClassName='WeChatMainWndForPC',
            Name='微信',
            timeout=30
        )
        log_info("找到微信主窗口，正在激活...", log_widget)
        wechat_window.SetActive()
        wechat_window.MoveToCenter()
    except auto.FindControlNotFoundError:
        log_info("未找到微信主窗口", log_widget)
        start_button.config(state=tk.NORMAL)
        return

    for account, name in zip(accounts, names):
        try:
            log_info(f"开始添加好友：{account}（备注名：{name}）", log_widget)

            wechat_window.SetActive()
            wechat_window.MoveToCenter()
            time.sleep(random.uniform(1, 2))

            # 点击聊天 -> 通讯录 -> 添加朋友
            chat_button = wechat_window.ButtonControl(Name='聊天')
            click_control(chat_button)
            time.sleep(1)

            contacts_button = wechat_window.ButtonControl(Name='通讯录')
            click_control(contacts_button)
            time.sleep(1)

            add_friend_button = wechat_window.ButtonControl(Name="添加朋友")
            click_control(add_friend_button)
            time.sleep(1)

            # 输入微信号
            search_box = wechat_window.EditControl(Name='微信号/手机号')
            click_control(search_box)
            search_box.SendKeys('{Ctrl}a{Delete}')
            search_box.SendKeys(account)
            time.sleep(1)

            contact_result = wechat_window.ListItemControl(
                NameRegex=f"搜索[:：]\\s*{account}",
                searchDepth=10,
                timeout=15
            )
            click_control(contact_result)
            time.sleep(1)
            image_path = resource_path("add_to_contacts.png")
            if not click_image(image_path):
                log_info("添加到通讯录失败", log_widget)
                continue

            time.sleep(1)

            remark_window = None
            for window in wechat_window.GetChildren():
                rect = window.BoundingRectangle
                if rect.width() > 100 and rect.height() > 50:
                    remark_window = window
                    break

            if not remark_window:
                log_info("未找到窗口", log_widget)
                continue

            # 填写验证信息
            verify_box = remark_window.EditControl(
                NameRegex="验证信息|我是.*",
                searchDepth=10,
                timeout=10
            )
            click_control(verify_box)
            verify_box.SendKeys('{Ctrl}a{Delete}')
            verify_box.SendKeys(verify_message)
            time.sleep(1)

            # 填写备注
            remarks_box = remark_window.EditControl(Name="", timeout=5)
            click_control(remarks_box)
            remarks_box.SendKeys('{Ctrl}a{Delete}')
            remarks_box.SendKeys(name)
            time.sleep(1)

            # 点击确定
            confirm_button = remark_window.ButtonControl(Name="确定", timeout=5)
            click_control(confirm_button)

            wait_time = random.uniform(min_interval, max_interval)
            log_info(f"等待 {int(wait_time)} 秒（{int(wait_time // 60)}分{int(wait_time % 60)}秒）", log_widget)

            def countdown(count):
                while count >= 0:
                    mins, secs = divmod(count, 60)
                    update_countdown_callback(f"⏳ 剩余等待时间：{mins}分{secs}秒")
                    time.sleep(1)
                    count -= 1

            countdown_thread = threading.Thread(target=countdown, args=(int(wait_time),))
            countdown_thread.start()
            countdown_thread.join()

        except Exception as e:
            log_info(f"添加好友 {account} 失败 {e}", log_widget)

    log_info("自动化流程已完成", log_widget)
    update_countdown_callback("✅ 自动化流程完成")

    # 启用开始按钮
    start_button.config(state=tk.NORMAL)


# GUI 主程序
def start_gui():
    root = tk.Tk()
    root.title("豪哥666")
    root.geometry("300x600")
    root.resizable(False, False)

    tk.Label(root, text="微信安装路径").place(x=20, y=10)
    path_entry = tk.Entry(root, width=36)
    path_entry.place(x=20, y=30)

    tk.Label(root, text="账号+备注（每行一个，空格分隔）").place(x=20, y=60)
    account_text = tk.Text(root, height=10, width=36)
    account_text.place(x=20, y=80)

    tk.Label(root, text="验证信息").place(x=20, y=220)
    verify_entry = tk.Entry(root, width=36)
    verify_entry.place(x=20, y=240)

    tk.Label(root, text="添加间隔").place(x=60, y=270)
    tk.Label(root, text="秒").place(x=220, y=270)
    interval_frame = tk.Frame(root)
    interval_frame.place(x=120, y=270)

    min_interval_entry = tk.Entry(interval_frame, width=5)
    min_interval_entry.pack(side=tk.LEFT)
    tk.Label(interval_frame, text="~").pack(side=tk.LEFT)
    max_interval_entry = tk.Entry(interval_frame, width=5)
    max_interval_entry.pack(side=tk.LEFT)

    start_button = tk.Button(root, text="开始", width=15, bg="green", fg="white",
                             font=("微软雅黑", 10, "bold"), relief="raised", bd=3,
                             activebackground="dark green", command=lambda: on_start())

    start_button.place(x=90, y=310)

    # 倒计时标签容器（用于居中）
    countdown_frame = tk.Frame(root)
    countdown_frame.place(x=0, y=350, relwidth=1)  # 横向撑满，y控制垂直位置

    countdown_label = tk.Label(countdown_frame, text="⏳ 等待开始...", font=("微软雅黑", 12, "bold"), fg="blue")
    countdown_label.pack()

    log_frame = tk.Frame(root)
    log_frame.place(x=15, y=380, width=280, height=200)
    log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', height=10, width=40)
    log_text.place(x=0, y=0, relwidth=1, relheight=1)

    def update_countdown(text):
        countdown_label.config(text=text)

    def on_start():
        wechat_path = path_entry.get().strip()
        verify_message = verify_entry.get().strip()
        if not verify_message:
            messagebox.showwarning("警告", "验证信息不能为空！")
            return
        raw_input = account_text.get("1.0", tk.END).strip().splitlines()
        min_interval_str = min_interval_entry.get().strip()
        max_interval_str = max_interval_entry.get().strip()

        accounts = []
        names = []

        for line in raw_input:
            parts = line.strip().split()
            if len(parts) >= 2:
                accounts.append(parts[0])
                names.append(' '.join(parts[1:]))

        if not os.path.isfile(wechat_path):
            messagebox.showerror("错误", "微信路径无效，请重新输入！")
            return

        if len(accounts) != len(names):
            messagebox.showerror("错误", "账号与备注数量不匹配，请检查格式！")
            return

        try:
            min_interval = float(min_interval_str)
            max_interval = float(max_interval_str)
            if min_interval < 1 or max_interval < min_interval:
                raise ValueError("请输入有效的间隔范围。")
        except ValueError as ve:
            messagebox.showerror("错误", str(ve) + "请确保最小值至少为1且小于最大值。")
            return

        config = {
            'wechat_path': wechat_path,
            'accounts': accounts,
            'names': names,
            'verify_message': verify_message,
            'min_interval': min_interval,
            'max_interval': max_interval
        }

        thread = threading.Thread(target=run_automation, args=(config, log_text, update_countdown, start_button))
        thread.start()

    root.mainloop()

def resource_path(relative_path):
    """获取资源文件的绝对路径，兼容 PyInstaller 打包后的路径"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的临时目录
        base_path = sys._MEIPASS
    else:
        # 脚本当前目录
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    start_gui()
