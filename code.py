import os
import sys
import ctypes
import subprocess
import requests
from threading import Thread
from tkinter import (Tk, Label, Button, Frame, filedialog, messagebox,
                     StringVar, SUNKEN, END, ttk)
from tkinter.scrolledtext import ScrolledText  # 修复ScrolledText导入


class WindowsReinstaller:
    def __init__(self, root):
        self.root = root
        self.root.title("EasyRestallZ")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        # 检查管理员权限
        self.is_admin = self.check_admin()
        if not self.is_admin:
            messagebox.showerror("权限不足", "请以管理员身份运行程序！")
            sys.exit(1)

        # 获取系统版本
        self.win_version = self.get_windows_version()

        # 创建UI
        self.create_ui()

        # 微软官方重装工具链接（支持Win10/11）
        self.mct_url = "https://go.microsoft.com/fwlink/?LinkId=691209"

    def check_admin(self):
        """检查是否以管理员身份运行"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def get_windows_version(self):
        """从注册表读取Windows版本信息"""
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
                0, winreg.KEY_READ
            )
            product_name = winreg.QueryValueEx(key, "ProductName")[0]
            build = winreg.QueryValueEx(key, "CurrentBuild")[0]
            ubr = winreg.QueryValueEx(key, "UBR")[0]
            return f"{product_name} (Build {build}.{ubr})"
        except Exception as e:
            return f"获取版本失败: {str(e)}"

    def create_ui(self):
        """构建图形界面"""
        font = ("SimHei", 10)

        # 版本信息区域
        version_frame = Frame(self.root)
        version_frame.pack(pady=10, padx=20, fill="x")
        Label(version_frame, text="当前系统版本:", font=font).pack(anchor="w")
        self.version_label = Label(
            version_frame, text=self.win_version,
            font=(font[0], font[1], "bold"), fg="#0066CC"
        )
        self.version_label.pack(anchor="w", pady=5)

        # 功能按钮区域
        btn_frame = Frame(self.root)
        btn_frame.pack(pady=20)
        self.btn_online = Button(
            btn_frame, text="官网一键重装", font=font, width=20, height=2,
            command=self.start_online_reinstall
        )
        self.btn_online.grid(row=0, column=0, padx=15)

        self.btn_local = Button(
            btn_frame, text="本地ISO安装", font=font, width=20, height=2,
            command=self.start_local_reinstall
        )
        self.btn_local.grid(row=0, column=1, padx=15)

        # 日志显示区域
        log_frame = Frame(self.root)
        log_frame.pack(pady=10, padx=20, fill="both", expand=True)
        Label(log_frame, text="操作日志:", font=font).pack(anchor="w")
        self.log_text = ScrolledText(
            log_frame, font=("Consolas", 9), wrap="word", state="disabled"
        )
        self.log_text.pack(fill="both", expand=True, pady=5)

        # 状态栏
        self.status_var = StringVar(value="就绪")
        status_bar = Label(
            self.root, textvariable=self.status_var, bd=1, relief=SUNKEN,
            anchor="w", font=font
        )
        status_bar.pack(side="bottom", fill="x")

    def log(self, message):
        """添加日志到界面"""
        self.log_text.config(state="normal")
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)
        self.log_text.config(state="disabled")

    def start_online_reinstall(self):
        """通过微软官网工具重装"""
        if messagebox.askyesno("警告", "此操作将清除C盘数据！\n建议先备份文件，是否继续？"):
            self.btn_online.config(state="disabled")
            self.btn_local.config(state="disabled")
            self.log("=== 开始官网重装流程 ===")
            Thread(target=self.download_mct, daemon=True).start()

    def start_local_reinstall(self):
        """从本地ISO文件安装"""
        iso_path = filedialog.askopenfilename(
            title="选择Windows ISO文件",
            filetypes=[("ISO镜像", "*.iso"), ("所有文件", "*.*")]
        )
        if iso_path and messagebox.askyesno("警告", f"将从 {iso_path} 安装系统，是否继续？"):
            self.btn_online.config(state="disabled")
            self.btn_local.config(state="disabled")
            self.log(f"=== 开始本地ISO安装: {iso_path} ===")
            Thread(target=self.mount_iso, args=(iso_path,), daemon=True).start()

    def download_mct(self):
        """下载微软MediaCreationTool并运行"""
        try:
            self.status_var.set("正在下载官方重装工具...")
            mct_path = os.path.join(os.environ["TEMP"], "MediaCreationTool.exe")

            # 下载工具（带进度显示）
            with requests.get(self.mct_url, stream=True) as r:
                total_size = int(r.headers.get("content-length", 0))
                downloaded = 0
                with open(mct_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        downloaded += len(chunk)
                        f.write(chunk)
                        progress = (downloaded / total_size) * 100
                        self.status_var.set(f"下载进度: {progress:.1f}%")
                        self.log(f"已下载: {progress:.1f}%")

            # 运行工具
            self.log("启动官方重装工具...")
            subprocess.run([mct_path, "/Selfhost"], check=True)
            self.status_var.set("工具已启动，请按提示操作")
        except Exception as e:
            self.log(f"错误: {str(e)}")
            self.status_var.set("操作失败")
        finally:
            self.btn_online.config(state="normal")
            self.btn_local.config(state="normal")

    def mount_iso(self, iso_path):
        """挂载ISO并启动安装程序"""
        try:
            self.status_var.set("正在挂载ISO镜像...")
            # 使用PowerShell挂载ISO
            subprocess.run(
                f'powershell Mount-DiskImage -ImagePath "{iso_path}"',
                shell=True, check=True
            )

            # 获取挂载盘符
            drive = subprocess.check_output(
                f'powershell "(Get-DiskImage \'{iso_path}\').DevicePath | Get-Volume | Select-Object -ExpandProperty DriveLetter"',
                shell=True
            ).decode().strip()

            if drive:
                self.log(f"ISO已挂载至 {drive}:\\，启动安装程序...")
                subprocess.run(f"{drive}:\\setup.exe", check=True)
                self.status_var.set("安装程序已启动")
            else:
                self.log("挂载成功，但未找到盘符，请手动运行setup.exe")
        except Exception as e:
            self.log(f"错误: {str(e)}")
            self.status_var.set("操作失败")
        finally:
            self.btn_online.config(state="normal")
            self.btn_local.config(state="normal")


if __name__ == "__main__":
    root = Tk()
    app = WindowsReinstaller(root)
    root.mainloop()
