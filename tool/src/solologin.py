import os
import sys
import tkinter as tk
from tkinter import messagebox
import subprocess
import time
import ctypes
import winreg

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ==================== 权限检查与提权 ====================

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        try:
            # 重新启动自身，并请求管理员权限
            script = os.path.abspath(sys.argv[0])
            params = ' '.join([script] + sys.argv[1:])
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, params, None, 1
            )
            sys.exit()
        except Exception as e:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("错误", f"无法获取管理员权限：\n{str(e)}\n\n请右键文件，选择“以管理员身份运行”。")
            root.destroy()
            sys.exit()

# ==================== 关闭并清除 Edge 数据 ====================

def close_edge():
    try:
        subprocess.run(['taskkill', '/f', '/im', 'msedge.exe'], capture_output=True, timeout=5)
        time.sleep(1)
    except:
        pass

def clear_edge_account():
    try:
        edge_user_data = os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Microsoft\Edge\User Data")

        # 删除关键文件
        account_files = [
            "Current Session", "Current Tabs", "Last Session", "Last Tabs",
            "Preferences", "Secure Preferences", "Web Data", "Top Sites", "Bookmarks"
        ]
        for file in account_files:
            file_path = os.path.join(edge_user_data, file)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                elif os.path.isdir(file_path):
                    import shutil
                    shutil.rmtree(file_path)
            except:
                pass

        # 清理各 Profile 的 Preferences（移除账户信息）
        if os.path.exists(edge_user_data):
            for item in os.listdir(edge_user_data):
                item_path = os.path.join(edge_user_data, item)
                if os.path.isdir(item_path) and (item == "Default" or item.startswith("Profile ")):
                    for prefs_name in ["Preferences", "Secure Preferences"]:
                        prefs_path = os.path.join(item_path, prefs_name)
                        if os.path.exists(prefs_path):
                            try:
                                with open(prefs_path, 'r+', encoding='utf-8') as f:
                                    content = f.read()
                                    # 清除账户标识
                                    content = content.replace('"account_id":"', '"account_id":"')
                                    content = content.replace('"gaia_id":"', '"gaia_id":"')
                                    content = content.replace('"email":"', '"email":"')
                                    content = content.replace('"sync"', '"sync_disabled"')
                                    f.seek(0)
                                    f.write(content)
                                    f.truncate()
                            except:
                                pass

        # 删除注册表项
        reg_paths = [
            r"Software\Microsoft\Edge\AccountManager",
            r"Software\Microsoft\Edge\BrowserDiagnostics",
            r"Software\Microsoft\Edge\Default Browser",
            r"Software\Microsoft\Edge\First Run",
            r"Software\Microsoft\Edge\SyncData",
            r"Software\Microsoft\Edge\Update",
            r"Software\Microsoft\Edge\WebView2"
        ]
        for path in reg_paths:
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
            except:
                pass

        # 删除 Windows 凭据
        try:
            result = subprocess.run(['cmdkey', '/list'], capture_output=True, text=True, encoding='gbk')
            for line in result.stdout.splitlines():
                if 'MicrosoftAccount' in line or 'Edge' in line:
                    parts = line.split('target=')
                    if len(parts) > 1:
                        target = parts[1].strip()
                        try:
                            subprocess.run(['cmdkey', '/delete:' + target], capture_output=True)
                        except:
                            pass
        except:
            pass

    except:
        pass

# ==================== 解绑微软账户 ====================

def unbind_microsoft_account():
    status_text.config(state=tk.NORMAL)
    status_text.delete('1.0', tk.END)
    status_text.insert(tk.END, "正在解绑微软账户...\n")
    status_text.update()

    close_edge()
    status_text.insert(tk.END, "✓ 关闭 Edge\n")
    status_text.update()

    clear_edge_account()
    status_text.insert(tk.END, "✓ 清除 Edge 账户数据\n")
    status_text.update()

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\IdentityCRL")
    except:
        pass
    status_text.insert(tk.END, "✓ 删除 IdentityCRL\n")
    status_text.update()

    try:
        result = subprocess.run(['cmdkey', '/list'], capture_output=True, text=True, encoding='gbk')
        for line in result.stdout.splitlines():
            if 'MicrosoftAccount' in line:
                parts = line.split('target=')
                if len(parts) > 1:
                    target = parts[1].strip()
                    try:
                        subprocess.run(['cmdkey', '/delete:' + target], capture_output=True)
                    except:
                        pass
    except:
        pass
    status_text.insert(tk.END, "✓ 删除系统凭证\n")
    status_text.insert(tk.END, "解绑完成，请使用本地账户登录。\n")
    status_text.config(state=tk.DISABLED)

# ==================== 创建本地账户（居中弹窗 + 左右对齐）====================

def create_local_account():
    input_window = tk.Toplevel(root)
    input_window.title("创建本地账户")
    input_window.geometry("400x200")
    input_window.resizable(False, False)
    input_window.transient(root)
    input_window.grab_set()

    # ✅ 居中显示弹窗
    root.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - (200)
    y = root.winfo_y() + (root.winfo_height() // 2) - (125)
    input_window.geometry(f"+{x}+{y}")

    # 主框架，使用 grid 布局
    frame = tk.Frame(input_window)
    frame.pack(pady=20, padx=30, fill="both", expand=True)

    # 设置列权重：左右留白，中间内容居中
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_columnconfigure(3, weight=1)

    # 用户名
    tk.Label(frame, text="用户名:", font=("微软雅黑", 10)).grid(row=0, column=1, sticky='e', padx=(0, 10), pady=12)
    username_entry = tk.Entry(frame, width=20, font=("微软雅黑", 10))
    username_entry.grid(row=0, column=2, sticky='w', pady=12)

    # 密码
    tk.Label(frame, text="密码（可为空）:", font=("微软雅黑", 10)).grid(row=1, column=1, sticky='e', padx=(0, 10), pady=12)
    password_entry = tk.Entry(frame, width=20, show="*", font=("微软雅黑", 10))
    password_entry.grid(row=1, column=2, sticky='w', pady=12)

    # 按钮区域：居中放置，间距加大
    btn_frame = tk.Frame(frame)
    btn_frame.grid(row=2, column=1, columnspan=2, pady=20)

    def on_confirm():
        username = username_entry.get().strip()
        password = password_entry.get()

        if not username:
            messagebox.showwarning("输入错误", "用户名不能为空！", parent=input_window)
            return

        # 检查用户是否存在
        try:
            result = subprocess.run(['net', 'user', username], capture_output=True, text=True, encoding='gbk')
            if result.returncode == 0:
                messagebox.showerror("创建失败", f"用户 '{username}' 已存在！", parent=input_window)
                return
        except Exception as e:
            messagebox.showerror("错误", f"检查用户时出错：\n{str(e)}", parent=input_window)
            return

        # 创建用户
        try:
            subprocess.run(['net', 'user', username, password, '/add'], check=True, encoding='gbk')
            subprocess.run(['net', 'localgroup', 'Administrators', username, '/add'], check=True, encoding='gbk')
            status_text.config(state=tk.NORMAL)
            status_text.delete('1.0', tk.END)
            status_text.insert(tk.END, f"本地账户 '{username}' 创建成功！\n")
            status_text.config(state=tk.DISABLED)
            input_window.destroy()
        except Exception as e:
            messagebox.showerror("失败", f"创建失败：\n{str(e)}", parent=input_window)

    # 按钮：加大间距，更好看
    tk.Button(btn_frame, text="创建", width=10, font=("微软雅黑", 9), command=on_confirm).pack(side=tk.LEFT, padx=25)
    tk.Button(btn_frame, text="取消", width=10, font=("微软雅黑", 9), command=input_window.destroy).pack(side=tk.RIGHT, padx=25)

    # 设置输入框自动获得焦点
    username_entry.focus()

# ==================== 主界面布局（紧凑 + 左右对齐）====================

if not is_admin():
    run_as_admin()
else:
    root = tk.Tk()
    root.title("账户管理工具")
    root.geometry("640x360")
    root.resizable(False, False)
    icon_path = resource_path("assets\\2395.ico")
    root.iconbitmap(icon_path)
    
    # 计算主界面居中位置
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - 700) // 2
    y = (screen_height - 500) // 2
    root.geometry(f"+{x}+{y}")

    # 标题
    title_label = tk.Label(root, text="账户管理工具", font=("微软雅黑", 16))
    title_label.pack(pady=20)

    # 按钮区域（左右对齐）
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=20)

    tk.Button(btn_frame, text="创建本地账户", width=15, height=2, font=("微软雅黑", 10), command=create_local_account).pack(side=tk.LEFT, padx=20)
    tk.Button(btn_frame, text="解绑微软账户", width=15, height=2, font=("微软雅黑", 10), command=unbind_microsoft_account).pack(side=tk.RIGHT, padx=20)

    # 状态框
    status_frame = tk.Frame(root)
    status_frame.pack(pady=20, fill="both", expand=True, padx=40)

    status_scrollbar = tk.Scrollbar(status_frame)
    status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    status_text = tk.Text(
        status_frame,
        height=10,
        state=tk.DISABLED,
        yscrollcommand=status_scrollbar.set,
        font=("Consolas", 9)
    )
    status_text.pack(fill="both", expand=True)
    status_scrollbar.config(command=status_text.yview)

    root.mainloop()