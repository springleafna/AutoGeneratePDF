# -*- coding: utf-8 -*-
"""
AutoGeneratePDF - 图形界面版
作者：springleaf
用途：唤唤专用
"""

import os
import re
import sys
import time
import logging
import base64
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime


# ========== 配置区 (Configuration) ==========
class Config:
    LANGUAGE_BUTTONS = [
        ("打印中英文", "中英文"),
        ("打印英文", "英文"),
        ("打印中文", "中文")
    ]
    SELENIUM_TIMEOUT = 20

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def resource_path(relative_path):
    """ 获取资源的绝对路径，支持 PyInstaller 打包环境 """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 会创建一个临时文件夹，并把路径存储在 _MEIPASS 中
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def create_date_folder_on_desktop():
    try:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        base_folder = os.path.join(desktop, "AutoGeneratePDF")
        os.makedirs(base_folder, exist_ok=True)
        date_str = datetime.now().strftime("%y%m%d")
        date_folder = os.path.join(base_folder, date_str)
        os.makedirs(date_folder, exist_ok=True)
        logger.info(f"✅ 文件将保存至：{date_folder}")
        return date_folder
    except Exception as e:
        logger.error(f"创建文件夹失败: {e}")
        return None


def _clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()


class PrintToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoGeneratePDF")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        # 隐藏窗口直到居中完成
        self.root.withdraw()

        self.url_entries = []
        self.url_queue = []
        self.total_urls = 0

        self._setup_ui()

        # 在界面初始化完成后居中并显示窗口
        self.root.after(100, self._show_centered_window)

    def _show_centered_window(self):
        """居中并显示窗口"""
        self.center_window()
        self.root.deiconify()  # 显示窗口
        self.root.focus_force()  # 获取焦点

    def center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()  # 确保获取到正确的窗口尺寸
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"+{x}+{y}")

    def _add_url_entry(self, is_first=False):
        row_frame = ttk.Frame(self.url_list_frame)
        row_frame.pack(fill=tk.X, pady=2)
        entry = ttk.Entry(row_frame, width=60, font=("微软雅黑", 10))
        entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.url_entries.append(entry)
        if not is_first:
            remove_btn = ttk.Button(row_frame, text="-", width=3,
                                    command=lambda rf=row_frame, en=entry: self._remove_url_entry(rf, en))
            remove_btn.pack(side=tk.LEFT, padx=(5, 0))

    def _remove_url_entry(self, frame_to_remove, entry_to_remove):
        frame_to_remove.destroy()
        self.url_entries.remove(entry_to_remove)

    def _setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        title_label = ttk.Label(main_frame, text="📚 AutoGeneratePDF", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=(0, 10))
        url_area_frame = ttk.LabelFrame(main_frame, text=" 网址列表 ", padding=10)
        url_area_frame.pack(fill=tk.X, pady=10)
        add_btn = ttk.Button(url_area_frame, text="✚ 添加网址", command=self._add_url_entry)
        add_btn.pack(anchor=tk.W, pady=(0, 10))
        self.add_url_button = add_btn
        canvas = tk.Canvas(url_area_frame, borderwidth=0, background="#ffffff")
        self.url_list_frame = ttk.Frame(canvas)
        scrollbar = ttk.Scrollbar(url_area_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((4, 4), window=self.url_list_frame, anchor="nw")
        self.url_list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        self.start_btn = ttk.Button(button_frame, text="✅ 开始打印", command=self.start_printing_all)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.exit_btn = ttk.Button(button_frame, text="❌ 退出", command=self.root.quit)
        self.exit_btn.pack(side=tk.LEFT, padx=5)
        self.status_var = tk.StringVar(value="请添加网址后开始任务...")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, wraplength=550, foreground="gray",
                                      font=("微软雅黑", 9))
        self.status_label.pack(pady=(10, 0))
        self._add_url_entry(is_first=True)

    def update_status(self, message):
        self.status_var.set(message)
        logger.info(message)
        self.root.update_idletasks()

    def start_printing_all(self):
        urls = [entry.get().strip() for entry in self.url_entries if entry.get().strip()]
        if not urls:
            messagebox.showerror("错误", "请输入至少一个网址！")
            return
        for url in urls:
            if not url.startswith(("http://", "https://")):
                messagebox.showerror("错误", f"网址格式不正确：\n{url}")
                return
        self.url_queue = urls
        self.total_urls = len(urls)
        self.start_btn.config(state="disabled")
        self.add_url_button.config(state="disabled")
        self.update_status(f"🚀 任务队列已创建，共 {self.total_urls} 个任务。")
        self.root.after(100, self._process_next_url)

    def _process_next_url(self):
        if not self.url_queue:
            self.update_status("🎉 全部任务完成！请在桌面 'AutoGeneratePDF' 文件夹中查看结果。")
            messagebox.showinfo("成功", f"所有 {self.total_urls} 个打印任务已处理完毕！")
            self.start_btn.config(state="normal")
            self.add_url_button.config(state="normal")
            return
        current_url = self.url_queue.pop(0)
        task_num = self.total_urls - len(self.url_queue)
        self.update_status(f"📄 开始处理第 {task_num} / {self.total_urls} 个任务: {current_url}")
        success = self.run_print_job(current_url)
        if not success:
            self.update_status(f"⚠️ 第 {task_num} 个任务处理失败，跳过...")
            logger.warning(f"任务 {current_url} 处理失败，已跳过。")
        self.root.after(100, self._process_next_url)

    def _setup_driver(self):
        """ 配置并返回一个 Microsoft Edge WebDriver 实例 (为原生PDF打印优化) """
        self.update_status("配置 Edge 浏览器...")
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        # 使用 resource_path 函数来获取驱动的绝对路径，以兼容打包后的环境
        driver_path = resource_path("msedgedriver.exe")

        self.update_status(f"使用本地 Edge WebDriver: {driver_path}")
        service = Service(executable_path=driver_path)

        self.update_status("启动 Edge 浏览器...")
        return webdriver.Edge(service=service, options=options)

    def _process_single_language(self, driver, btn_text, lang_tag, download_dir):
        self.update_status(f"正在处理：{btn_text}")
        wait = WebDriverWait(driver, Config.SELENIUM_TIMEOUT)
        try:
            lang_button_xpath = f"//button[.//span[contains(text(), '{btn_text}')]]"
            lang_button = wait.until(EC.element_to_be_clickable((By.XPATH, lang_button_xpath)))
            lang_button.click()
            time.sleep(2.5)
            page_title = driver.title
            base_filename = _clean_filename(page_title)
            self.update_status(f"获取到原始文件名: {base_filename}")
            if not base_filename:
                base_filename = f"未命名文档_{datetime.now().strftime('%H%M%S')}"
            print_options = {
                'landscape': False, 'displayHeaderFooter': False,
                'printBackground': True, 'preferCSSPageSize': True,
            }
            self.update_status("正在执行原生PDF生成命令...")
            result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
            pdf_data = base64.b64decode(result['data'])
            new_filename = f"{base_filename}_{lang_tag}.pdf"
            full_file_path = os.path.join(download_dir, new_filename)
            with open(full_file_path, 'wb') as f:
                f.write(pdf_data)
            self.update_status(f"✅ 成功保存：{new_filename}")
            return True
        except Exception as e:
            self.update_status(f"❌ 处理 '{btn_text}' 时失败: {e}")
            logger.error(f"处理 '{btn_text}' 时失败: {e}", exc_info=True)
            return False

    def run_print_job(self, url):
        driver = None
        try:
            download_dir = create_date_folder_on_desktop()
            if not download_dir: return False
            driver = self._setup_driver()
            if not driver: return False
            self.update_status(f"正在打开网页：{url}")
            driver.get(url)
            try:
                wait = WebDriverWait(driver, Config.SELENIUM_TIMEOUT)
                first_button_xpath = f"//button[.//span[contains(text(), '{Config.LANGUAGE_BUTTONS[0][0]}')]]"
                wait.until(EC.visibility_of_element_located((By.XPATH, first_button_xpath)))
                self.update_status("页面加载完成。")
            except Exception as e:
                self.update_status(f"❌ 等待页面初始元素超时。请检查网址。")
                logger.error(f"等待页面初始元素超时: {e}", exc_info=True)
                return False

            try:
                self.update_status("正在执行“预打印”以触发并稳定打印样式...")
                print_options = {
                    'landscape': False, 'displayHeaderFooter': False,
                    'printBackground': True, 'preferCSSPageSize': True,
                }
                driver.execute_cdp_cmd("Page.printToPDF", print_options)
                time.sleep(3)
                self.update_status("预打印完成，打印引擎已就绪。")
            except Exception as e:
                self.update_status(f"⚠️ 预打印失败: {e}。将继续尝试...")
                logger.warning(f"预打印失败: {e}", exc_info=True)

            self.update_status("开始处理语言版本...")
            all_langs_successful = True
            for btn_text, lang_tag in Config.LANGUAGE_BUTTONS:
                success = self._process_single_language(driver, btn_text, lang_tag, download_dir)
                if not success:
                    all_langs_successful = False
                time.sleep(2)
            return all_langs_successful
        except Exception as e:
            error_message = f"❌ 执行过程中发生严重错误：{str(e)}"
            self.update_status(error_message)
            logger.error(error_message, exc_info=True)
            return False
        finally:
            if driver:
                driver.quit()
                self.update_status("浏览器已关闭。")


if __name__ == "__main__":
    root = tk.Tk()
    app = PrintToolApp(root)
    root.mainloop()