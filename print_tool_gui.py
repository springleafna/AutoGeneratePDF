# -*- coding: utf-8 -*-
"""
AutoGeneratePDF - 图形界面版（用户输入网址自动打印PDF）
作者：springleaf
用途：唤唤专用
版本：2.0 (支持多网址)
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
from pywinauto import Application, timings, ElementAmbiguousError
from datetime import datetime


# ========== 配置区 (Configuration) ==========
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
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()


class PrintToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoGeneratePDF v2.0")
        self.root.geometry("900x750")
        self.root.resizable(True, True)  # 允许调整大小
        self.root.iconbitmap(resource_path("icon.ico"))

        # 新增：用于存储网址输入框的列表
        self.url_entries = []
        # 新增：用于处理任务队列
        self.url_queue = []
        self.total_urls = 0

        self._setup_ui()

    def _add_url_entry(self, is_first=False):
        """ 动态添加入网址输入框和删除按钮 """
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
        """ 移除一个网址输入框 """
        frame_to_remove.destroy()
        self.url_entries.remove(entry_to_remove)

    def _setup_ui(self):
        """ 初始化用户界面 (已修改以支持多网址) """
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="📚 AutoGeneratePDF", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=(0, 10))

        # --- 网址输入区 ---
        url_area_frame = ttk.LabelFrame(main_frame, text=" 网址列表 ", padding=10)
        url_area_frame.pack(fill=tk.X, pady=10)

        add_btn = ttk.Button(url_area_frame, text="✚ 添加网址", command=self._add_url_entry)
        add_btn.pack(anchor=tk.W, pady=(0, 10))
        self.add_url_button = add_btn  # 保存引用以便后续禁用

        # 创建一个可滚动的Frame来放置URL输入框
        canvas = tk.Canvas(url_area_frame, borderwidth=0, background="#ffffff")
        self.url_list_frame = ttk.Frame(canvas)  # 核心：所有输入框都放在这个Frame里
        scrollbar = ttk.Scrollbar(url_area_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((4, 4), window=self.url_list_frame, anchor="nw")

        self.url_list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # --- 按钮和状态区 ---
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

        # 初始化时添加第一个输入框
        self._add_url_entry(is_first=True)

    def update_status(self, message):
        """ 更新状态栏文本并刷新UI """
        self.status_var.set(message)
        logger.info(message)
        self.root.update_idletasks()  # 强制UI刷新

    def start_printing_all(self):
        """ 新增：验证所有用户输入并启动打印队列 """
        # 1. 收集所有非空的有效网址
        urls = [entry.get().strip() for entry in self.url_entries if entry.get().strip()]

        if not urls:
            messagebox.showerror("错误", "请输入至少一个网址！")
            return

        # 2. 验证所有网址格式
        for url in urls:
            if not url.startswith(("http://", "https://")):
                messagebox.showerror("错误", f"网址格式不正确：\n{url}\n\n请确保所有网址都以 http:// 或 https:// 开头。")
                return

        self.url_queue = urls
        self.total_urls = len(urls)

        # 3. 禁用按钮，防止重复点击
        self.start_btn.config(state="disabled")
        self.add_url_button.config(state="disabled")
        self.update_status(f"🚀 任务队列已创建，共 {self.total_urls} 个任务。正在准备环境...")

        # 4. 使用 after 启动第一个任务，避免阻塞GUI
        self.root.after(100, self._process_next_url)

    def _process_next_url(self):
        """ 新增：处理队列中的下一个网址 """
        if not self.url_queue:
            # 队列为空，所有任务完成
            self.update_status("🎉 全部任务完成！请在桌面 'AutoGeneratePDF' 文件夹中查看结果。")
            messagebox.showinfo("成功", f"所有 {self.total_urls} 个打印任务已处理完毕！")
            # 恢复按钮
            self.start_btn.config(state="normal")
            self.add_url_button.config(state="normal")
            return

        # 取出队列中的第一个网址进行处理
        current_url = self.url_queue.pop(0)
        task_num = self.total_urls - len(self.url_queue)
        self.update_status(f"📄 开始处理第 {task_num} / {self.total_urls} 个任务: {current_url}")

        # 执行单个打印任务
        success = self.run_print_job(current_url)

        if not success:
            # 如果中途失败，可以选择停止或继续。这里我们选择记录日志并继续
            self.update_status(f"⚠️ 第 {task_num} 个任务处理失败，跳过并继续下一个...")
            logger.warning(f"任务 {current_url} 处理失败，已跳过。")

        # 调度下一个任务
        self.root.after(100, self._process_next_url)

    def _setup_driver(self, download_dir):
        """ 配置并返回一个 Chrome WebDriver 实例 (无修改) """
        self.update_status("配置浏览器...")
        # ... (此函数内容与原来完全相同)
        options = Options()
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--kiosk-printing")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)

    def _handle_save_dialog(self, download_dir, lang_tag, base_filename):
        """ 处理“另存为”对话框 (无修改) """
        self.update_status(f"等待 '{lang_tag}' 语言的保存对话框...")
        # ... (此函数内容与原来完全相同)
        try:
            self.update_status("使用 'win32' 后端连接对话框...")
            app = Application(backend="win32").connect(
                title_re=Config.SAVE_AS_DIALOG_TITLE_RE,
                timeout=Config.DIALOG_TIMEOUT
            )
            dlg = app.window(title_re=Config.SAVE_AS_DIALOG_TITLE_RE)
            self.update_status("✅ 对话框已连接，正在精确定位控件...")
            dlg.wait("ready", timeout=Config.DIALOG_TIMEOUT)

            if not base_filename:
                base_filename = f"未命名文档_{datetime.now().strftime('%H%M%S')}"
                logger.warning("网页标题为空，使用默认文件名。")

            new_filename = f"{base_filename}_{lang_tag}.pdf"
            full_file_path = os.path.join(download_dir, new_filename)
            self.update_status(f"准备保存文件至: {full_file_path}")

            file_name_combo = dlg.child_window(class_name="ComboBox", found_index=0)
            filename_edit = file_name_combo.child_window(class_name="Edit")

            filename_edit.wait('visible', timeout=5)
            filename_edit.set_focus()
            filename_edit.set_edit_text("")
            filename_edit.type_keys(full_file_path, with_spaces=True)
            time.sleep(1)

            save_button = dlg.child_window(title="保存(&S)", class_name="Button")
            save_button.wait('enabled', timeout=5)
            save_button.click_input()
            self.update_status(f"已点击保存: {new_filename}")

            for _ in range(Config.FILE_SAVE_TIMEOUT * 2):
                if os.path.exists(full_file_path) and os.path.getsize(full_file_path) > 1024:
                    self.update_status(f"✅ 成功保存：{new_filename}")
                    return True
                time.sleep(0.5)

            logger.warning(f"⚠️ 文件保存超时或文件过小：{new_filename}")
            self.update_status(f"❌ 文件保存超时: {new_filename}")
            return False

        except ElementAmbiguousError as e:
            self.update_status(f"❌ 控件识别模糊，发现多个匹配项: {e}")
            logger.error(f"控件识别模糊，发现多个匹配项: {e}", exc_info=True)
            return False
        except timings.TimeoutError as e:
            self.update_status(f"❌ 在对话框内部查找控件时超时。")
            logger.error(f"查找控件时超时: {e}", exc_info=True)
            return False
        except Exception as e:
            self.update_status(f"❌ 处理保存对话框时发生未知错误：{e}")
            logger.error(f"处理保存对话框时发生异常: {e}", exc_info=True)
            return False

    def _process_single_language(self, driver, btn_text, lang_tag, download_dir):
        """ 核心逻辑：为单一语言获取标题、点击按钮、打印并保存 (无修改) """
        self.update_status(f"正在处理：{btn_text}")
        # ... (此函数内容与原来完全相同)
        wait = WebDriverWait(driver, Config.SELENIUM_TIMEOUT)
        try:
            lang_button_xpath = f"//button[.//span[contains(text(), '{btn_text}')]]"
            lang_button = wait.until(EC.element_to_be_clickable((By.XPATH, lang_button_xpath)))
            lang_button.click()
            time.sleep(1.5)

            page_title = driver.title
            base_filename = _clean_filename(page_title)
            self.update_status(f"获取到原始文件名: {base_filename}")

            print_button_xpath = f"//button[.//span[contains(text(), '{Config.PRINT_BUTTON_TEXT}')]]"
            print_button = wait.until(EC.element_to_be_clickable((By.XPATH, print_button_xpath)))
            print_button.click()

            return self._handle_save_dialog(download_dir, lang_tag, base_filename)

        except Exception as e:
            self.update_status(f"❌ 处理 '{btn_text}' 时失败: {e}")
            logger.error(f"处理 '{btn_text}' 时失败: {e}", exc_info=True)
            return False

    def run_print_job(self, url):
        """
        完整执行单个网址的打印流程。
        (已修改：移除原有的最终状态更新和按钮恢复逻辑，并返回执行结果)
        """
        driver = None
        try:
            download_dir = create_date_folder_on_desktop()
            driver = self._setup_driver(download_dir)

            self.update_status(f"正在打开网页：{url}")
            driver.get(url)

            # 主循环，处理每一种语言
            all_langs_successful = True
            for btn_text, lang_tag in Config.LANGUAGE_BUTTONS:
                success = self._process_single_language(driver, btn_text, lang_tag, download_dir)
                if not success:
                    all_langs_successful = False
                time.sleep(2)  # 在处理下一种语言前稍作停顿

            return all_langs_successful

        except Exception as e:
            error_message = f"❌ 执行过程中发生严重错误：{str(e)}"
            self.update_status(error_message)
            logger.error(error_message, exc_info=True)
            # 不再弹窗，只记录日志并返回失败，由上层循环决定如何处理
            return False

        finally:
            if driver:
                driver.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = PrintToolApp(root)
    root.mainloop()