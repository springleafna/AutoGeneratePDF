# -*- coding: utf-8 -*-
"""
打印工具 - 图形界面版（用户输入网址）
作者：springleaf
用途：唤唤专用
"""

import os
import sys
import time
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from pywinauto import Application
import re
from datetime import datetime

# ========== 配置区 ==========
LANGUAGE_BUTTONS = [
    ("打印中英文", "中英文"),
    ("打印英文", "英文"),
    ("打印中文", "中文")
]
# ============================

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 获取程序运行目录（支持 PyInstaller 打包）
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# ✅ 新增：在桌面上创建 AutoGeneratePDF/YYMMDD 文件夹
def create_date_folder_on_desktop():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    base_folder = os.path.join(desktop, "AutoGeneratePDF")
    os.makedirs(base_folder, exist_ok=True)  # 创建主文件夹

    # 生成日期子文件夹：如 250914 表示 2025年9月14日
    date_str = datetime.now().strftime("%y%m%d")  # 两位年+两位月+两位日
    date_folder = os.path.join(base_folder, date_str)
    os.makedirs(date_folder, exist_ok=True)  # 创建日期子文件夹

    logger.info(f"✅ 文件将保存至：{date_folder}")
    return date_folder

class PrintToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("打印工具 - 自动生成PDF（桌面保存）")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        ## self.root.iconbitmap(resource_path("icon.ico"))  # 设置图标

        # 主框架
        frame = ttk.Frame(root, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # 标题
        title_label = ttk.Label(frame, text="📚 打印工具", font=("微软雅黑", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # 网址输入
        ttk.Label(frame, text="请输入打印页面网址：", font=("微软雅黑", 11)).grid(row=1, column=0, sticky=tk.W)
        self.url_entry = ttk.Entry(frame, width=50)
        self.url_entry.grid(row=2, column=0, columnspan=2, pady=(0, 15), sticky=(tk.W, tk.E))
        self.url_entry.insert(0, "https://")

        # 按钮区域
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=15)

        self.start_btn = ttk.Button(button_frame, text="✅ 开始打印", command=self.start_printing)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.exit_btn = ttk.Button(button_frame, text="❌ 退出", command=root.quit)
        self.exit_btn.pack(side=tk.LEFT, padx=5)

        # 进度显示
        self.status_var = tk.StringVar(value="等待输入网址...")
        self.status_label = ttk.Label(frame, textvariable=self.status_var, foreground="gray", font=("微软雅黑", 9))
        self.status_label.grid(row=4, column=0, columnspan=2, pady=(10, 0))

        # 绑定回车键
        self.root.bind('<Return>', lambda e: self.start_printing())

    def start_printing(self):
        url = self.url_entry.get().strip()
        if not url or not url.startswith(("http://", "https://")):
            messagebox.showerror("错误", "请正确输入网址（以 http:// 或 https:// 开头）")
            return

        # 禁用按钮防止重复点击
        self.start_btn.config(state="disabled")
        self.status_var.set("正在准备保存路径...")

        # 在新线程中运行（避免 GUI 卡死）
        self.root.after(100, lambda: self.run_print_job(url))

    def run_print_job(self, url):
        try:
            # ✅ 创建桌面保存路径：AutoGeneratePDF/YYMMDD
            download_dir = create_date_folder_on_desktop()
            logger.info(f"📁 下载目录已设置为：{download_dir}")

            # Chrome 配置
            options = Options()
            options.add_experimental_option("prefs", {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "plugins.always_open_pdf_externally": False,
                "safebrowsing.enabled": True
            })
            options.add_argument("--start-maximized")
            options.add_argument("--disable-extensions")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")

            driver = None
            try:
                logger.info(f"正在打开网页：{url}")
                self.status_var.set("正在打开网页...")
                driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
                driver.get(url)

                # 等待页面加载
                time.sleep(5)

                # 提取学生姓名和日期（根据你的网页结构调整）
                student_name = "未知学生"
                date_str = datetime.now().strftime("%y%m%d")  # 用于文件名，与文件夹一致

                try:
                    # 👇 请根据你截图中的实际元素修改以下 XPath
                    name_elem = driver.find_element(By.XPATH, "//span[contains(@class, 'user-name')]")
                    student_name = name_elem.text.strip()
                except Exception:
                    logger.warning("未找到学生姓名元素，使用默认值")

                try:
                    date_elem = driver.find_element(By.XPATH, "//div[contains(@class, 'print-date')]")
                    date_str = date_elem.text.strip().replace("-", "")[-6:]  # 取后6位如 2025-09-14 → 250914
                except Exception:
                    logger.warning("未找到日期元素，使用当前日期")

                logger.info(f"学生姓名：{student_name}，日期：{date_str}")

                # 循环处理三种语言模式
                for btn_text, lang_tag in LANGUAGE_BUTTONS:
                    self.status_var.set(f"正在处理：{btn_text}...")
                    logger.info(f"正在处理：{btn_text}")

                    # 点击语言按钮
                    try:
                        button = driver.find_element(By.XPATH, f"//button[contains(text(), '{btn_text}')]")
                        button.click()
                        time.sleep(2)
                    except Exception as e:
                        logger.error(f"未找到按钮 {btn_text}：{e}")
                        continue

                    # 点击“打印”按钮
                    try:
                        print_btn = driver.find_element(By.XPATH, "//button[contains(text(), '打印')]")
                        print_btn.click()
                        time.sleep(2)
                    except Exception as e:
                        logger.error(f"第一次点击‘打印’失败：{e}")
                        continue

                    # 再次点击“打印”触发另存为对话框
                    try:
                        print_btn = driver.find_element(By.XPATH, "//button[contains(text(), '打印')]")
                        print_btn.click()
                        time.sleep(3)
                    except Exception as e:
                        logger.error(f"第二次点击‘打印’失败：{e}")
                        continue

                    # 等待“另存为”对话框弹出
                    app = None
                    for _ in range(15):
                        try:
                            app = Application(backend="uia").connect(title_re=".*另存为.*", timeout=1)
                            break
                        except:
                            time.sleep(0.5)

                    if not app:
                        self.status_var.set(f"❌ 未检测到保存对话框：{btn_text}")
                        continue

                    dlg = app.window(title_re=".*另存为.*")
                    new_filename = f"学习资料_{lang_tag}_{student_name}_{date_str}.pdf"
                    dlg.Edit.set_text(new_filename)
                    time.sleep(0.3)
                    dlg.Button.click()

                    # 等待文件生成（最大等待15秒）
                    file_path = os.path.join(download_dir, new_filename)
                    for _ in range(30):
                        if os.path.exists(file_path) and os.path.getsize(file_path) > 1000:
                            logger.info(f"✅ 成功保存：{new_filename}")
                            break
                        time.sleep(0.5)
                    else:
                        logger.warning(f"⚠️ 文件未生成：{new_filename}")

                    time.sleep(2)

                self.status_var.set("🎉 全部完成！请查看桌面 'AutoGeneratePDF' 文件夹")
                messagebox.showinfo("成功", f"所有文件已生成！\n\n保存位置：\n{download_dir}")

            except Exception as e:
                logger.error(f"执行过程中发生错误：{e}")
                self.status_var.set(f"❌ 发生错误：{str(e)}")
                messagebox.showerror("错误", f"程序运行出错：\n{str(e)}\n\n请检查网络或联系管理员。")

        finally:
            if 'driver' in locals():
                driver.quit()
            self.start_btn.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = PrintToolApp(root)
    root.mainloop()