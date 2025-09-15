# -*- coding: utf-8 -*-
"""
AutoGeneratePDF - 图形界面版（用户输入网址）
作者：springleaf
用途：唤唤专用
"""

import os
import re
import sys
import time
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pywinauto import Application, timings
from datetime import datetime


# ========== 配置区 (Configuration) ==========
# 将所有可能变动的字符串放在这里，方便统一管理
class Config:
    # 语言按钮配置: (界面显示的按钮文本, 文件名中使用的语言标签)
    LANGUAGE_BUTTONS = [
        ("打印中英文", "中英文"),
        ("打印英文", "英文"),
        ("打印中文", "中文")
    ]
    # 页面元素文本
    PRINT_BUTTON_TEXT = "在线打印"

    # "另存为" 对话框标题 (使用正则表达式匹配，以防万一)
    SAVE_AS_DIALOG_TITLE_RE = ".*另存为.*"

    # 等待超时时间 (秒)
    SELENIUM_TIMEOUT = 15  # Selenium 等待元素加载的超时时间
    DIALOG_TIMEOUT = 15  # Pywinauto 等待对话框出现的超时时间
    FILE_SAVE_TIMEOUT = 30  # 等待文件保存完成的超时时间


# ============================================

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


# 获取程序运行目录（支持 PyInstaller 打包）
def resource_path(relative_path):
    """ 获取资源的绝对路径，支持 PyInstaller 打包环境 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# 在桌面上创建 AutoGeneratePDF/YYMMDD 文件夹
def create_date_folder_on_desktop():
    """ 在桌面创建当日日期的文件夹，并返回其路径 """
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    base_folder = os.path.join(desktop, "AutoGeneratePDF")
    os.makedirs(base_folder, exist_ok=True)

    date_str = datetime.now().strftime("%y%m%d")
    date_folder = os.path.join(base_folder, date_str)
    os.makedirs(date_folder, exist_ok=True)

    logger.info(f"✅ 文件将保存至：{date_folder}")
    return date_folder

# 清理文件名的辅助函数
def _clean_filename(name):
    """ 移除文件名中的非法字符 """
    # Windows 文件名非法字符: \ / : * ? " < > |
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

class PrintToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoGeneratePDF")
        self.root.geometry("500x320")  # 稍微增加高度以容纳状态文本
        self.root.resizable(False, False)
        # self.root.iconbitmap(resource_path("icon.ico")) # 如有图标，取消此行注释

        self._setup_ui()

    def _setup_ui(self):
        """ 初始化用户界面 """
        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        title_label = ttk.Label(frame, text="📚 AutoGeneratePDF", font=("微软雅黑", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        ttk.Label(frame, text="请输入打印页面网址：", font=("微软雅黑", 11)).grid(row=1, column=0, sticky=tk.W)
        self.url_entry = ttk.Entry(frame, width=50)
        self.url_entry.grid(row=2, column=0, columnspan=2, pady=(0, 15), sticky=(tk.W, tk.E))
        self.url_entry.insert(0, "https://")

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=15)

        self.start_btn = ttk.Button(button_frame, text="✅ 开始打印", command=self.start_printing)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.exit_btn = ttk.Button(button_frame, text="❌ 退出", command=self.root.quit)
        self.exit_btn.pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar(value="等待输入网址...")
        self.status_label = ttk.Label(frame, textvariable=self.status_var, wraplength=450, foreground="gray",
                                      font=("微软雅黑", 9))
        self.status_label.grid(row=4, column=0, columnspan=2, pady=(10, 0))

        self.root.bind('<Return>', lambda e: self.start_printing())

    def update_status(self, message):
        """ 更新状态栏文本并刷新UI """
        self.status_var.set(message)
        logger.info(message)
        self.root.update_idletasks()  # 强制UI刷新

    def start_printing(self):
        """ 验证用户输入并启动打印流程 """
        url = self.url_entry.get().strip()
        if not url.startswith(("http://", "https://")):
            messagebox.showerror("错误", "请正确输入网址（以 http:// 或 https:// 开头）")
            return

        self.start_btn.config(state="disabled")
        self.update_status("🚀 任务开始，正在准备环境...")

        # 使用 after 避免阻塞GUI
        self.root.after(100, lambda: self.run_print_job(url))

    def _setup_driver(self, download_dir):
        """ 配置并返回一个 Chrome WebDriver 实例 """
        self.update_status("配置浏览器...")
        options = Options()
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,  # 确保PDF被视为下载
        }
        options.add_experimental_option("prefs", prefs)
        # 这个参数会让 Chrome 跳过打印预览，直接调用系统保存对话框
        options.add_argument("--kiosk-printing")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)

    def _handle_save_dialog(self, download_dir, lang_tag, base_filename):
        """
        使用 pywinauto 处理“另存为”对话框。
        - base_filename: 从 Selenium 获取的、已经清理过的网页标题。
        """
        self.update_status(f"等待 '{lang_tag}' 语言的保存对话框...")
        try:
            # 等待对话框出现，而不是使用固定时间的 sleep
            app = Application(backend="uia").connect(
                title_re=Config.SAVE_AS_DIALOG_TITLE_RE,
                timeout=Config.DIALOG_TIMEOUT
            )
            dlg = app.window(title_re=Config.SAVE_AS_DIALOG_TITLE_RE)
            dlg.wait('exists', timeout=Config.DIALOG_TIMEOUT)

            # 构造新文件名，不再从对话框读取
            if not base_filename:
                # 如果由于某种原因没获取到标题，给一个默认名
                base_filename = f"未命名文档_{datetime.now().strftime('%H%M%S')}"
                logger.warning("网页标题为空，使用默认文件名。")

            new_filename = f"{base_filename}_{lang_tag}.pdf"

            # 设置新文件名并保存
            dlg.Edit.set_text(new_filename)
            time.sleep(0.5)
            dlg.Button("保存(S)").click()  # 根据截图，按钮文本是 "保存(S)"
            self.update_status(f"已触发展存为：{new_filename}")

            # 等待文件下载完成的逻辑不变
            file_path = os.path.join(download_dir, new_filename)
            for _ in range(Config.FILE_SAVE_TIMEOUT * 2):
                if os.path.exists(file_path) and os.path.getsize(file_path) > 1000:
                    self.update_status(f"✅ 成功保存：{new_filename}")
                    return True
                time.sleep(0.5)

            logger.warning(f"⚠️ 文件保存超时：{new_filename}")
            return False

        except timings.TimeoutError:
            self.update_status(f"❌ 未检测到 '{lang_tag}' 的保存对话框，跳过。")
            return False
        except Exception as e:
            self.update_status(f"❌ 处理保存对话框时出错：{e}")
            return False

    def _process_single_language(self, driver, btn_text, lang_tag, download_dir):
        """ 核心逻辑：为单一语言获取标题、点击按钮、打印并保存 """
        self.update_status(f"正在处理：{btn_text}")
        wait = WebDriverWait(driver, Config.SELENIUM_TIMEOUT)

        try:
            # 1. 点击语言切换按钮
            lang_button_xpath = f"//button[.//span[contains(text(), '{btn_text}')]]"
            lang_button = wait.until(EC.element_to_be_clickable((By.XPATH, lang_button_xpath)))
            lang_button.click()
            time.sleep(1.5)  # 点击后等待标题和内容刷新

            # 2. 获取并清理网页标题作为文件名
            page_title = driver.title
            base_filename = _clean_filename(page_title)
            self.update_status(f"获取到原始文件名: {base_filename}")

            # 3. 点击“在线打印”按钮
            print_button_xpath = f"//button[.//span[contains(text(), '{Config.PRINT_BUTTON_TEXT}')]]"
            print_button = wait.until(EC.element_to_be_clickable((By.XPATH, print_button_xpath)))
            print_button.click()

            # 4. 处理弹出的“另存为”对话框，并把获取到的文件名传进去
            return self._handle_save_dialog(download_dir, lang_tag, base_filename)

        except Exception as e:
            self.update_status(f"❌ 处理 '{btn_text}' 时失败: {e}")
            logger.error(f"处理 '{btn_text}' 时失败: {e}", exc_info=True)
            return False

    def run_print_job(self, url):
        """ 完整执行从打开网页到全部打印完成的流程 """
        driver = None
        try:
            download_dir = create_date_folder_on_desktop()
            driver = self._setup_driver(download_dir)

            self.update_status(f"正在打开网页：{url}")
            driver.get(url)

            # 主循环，处理每一种语言
            for btn_text, lang_tag in Config.LANGUAGE_BUTTONS:
                self._process_single_language(driver, btn_text, lang_tag, download_dir)
                time.sleep(2)  # 在处理下一种语言前稍作停顿

            self.update_status("🎉 全部任务完成！请在桌面 'AutoGeneratePDF' 文件夹中查看结果。")
            messagebox.showinfo("成功", f"所有打印任务已处理完毕！\n\n保存位置：\n{download_dir}")

        except Exception as e:
            error_message = f"❌ 执行过程中发生严重错误：{str(e)}"
            self.update_status(error_message)
            logger.error(error_message, exc_info=True)  # 记录完整的堆栈信息
            messagebox.showerror("严重错误", f"程序运行出错：\n{str(e)}\n\n请检查网络或联系管理员。")

        finally:
            if driver:
                driver.quit()
            self.start_btn.config(state="normal")  # 无论成功失败，最后都恢复按钮
            self.update_status("等待新的任务...")


if __name__ == "__main__":
    root = tk.Tk()
    app = PrintToolApp(root)
    root.mainloop()