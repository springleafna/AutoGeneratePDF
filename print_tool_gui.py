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


def get_unique_filepath(base_path):
    """
    确保文件路径唯一：如果文件已存在，则在文件名后添加 (1), (2), ...
    例如：report_中文.pdf → report_中文(1).pdf → report_中文(2).pdf
    """
    if not os.path.exists(base_path):
        return base_path

    root, ext = os.path.splitext(base_path)
    counter = 1
    while True:
        new_path = f"{root}({counter}){ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1


class PrintToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoGeneratePDF")
        self.root.geometry("650x600")
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
        row_frame.pack(fill="x", pady=2)
        entry = ttk.Entry(row_frame, width=60, font=("微软雅黑", 10))
        entry.pack(side="left", expand=True, fill="x")
        self.url_entries.append(entry)
        if not is_first:
            remove_btn = ttk.Button(row_frame, text="-", width=3,
                                    command=lambda rf=row_frame, en=entry: self._remove_url_entry(rf, en))
            remove_btn.pack(side="left", padx=(5, 0))

    def _remove_url_entry(self, frame_to_remove, entry_to_remove):
        frame_to_remove.destroy()
        self.url_entries.remove(entry_to_remove)

    def _setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill="both", expand=True)
        title_label = ttk.Label(main_frame, text="📚 AutoGeneratePDF", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=(0, 10))
        url_area_frame = ttk.LabelFrame(main_frame, text=" 网址列表 ", padding=10)
        url_area_frame.pack(fill="x", pady=10)
        add_btn = ttk.Button(url_area_frame, text="✚ 添加网址", command=self._add_url_entry)
        add_btn.pack(anchor="w", pady=(0, 10))
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
        self.start_btn.pack(side="left", padx=5)
        self.exit_btn = ttk.Button(button_frame, text="❌ 退出", command=self.root.quit)
        self.exit_btn.pack(side="left", padx=5)
        self.help_btn = ttk.Button(button_frame, text="📘 操作指南", command=self._show_help)
        self.help_btn.pack(side="left", padx=5)
        # 创建一个只读的 Text 控件用于显示日志
        self.status_text = tk.Text(main_frame, height=6, width=80, font=("微软雅黑", 9),
                                   bg="white", fg="gray", state="disabled", relief="sunken")
        self.status_text.pack(pady=(10, 0), fill="x")

        # 初始化内容
        self._append_status("请添加网址后开始任务...")
        self._add_url_entry(is_first=True)

    def _append_status(self, message):
        """向状态文本框追加一行信息，并自动滚动到底部"""
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)  # 滚动到最新行
        self.status_text.config(state="disabled")

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
        self._append_status(f"🚀 任务队列已创建，共 {self.total_urls} 个任务。")

        # 启动一次浏览器，传入后续任务
        self.root.after(100, lambda: self._process_all_urls())

    def _process_all_urls(self):
        driver = None
        try:
            driver = self._setup_driver()
            while self.url_queue:
                current_url = self.url_queue.pop(0)
                task_num = self.total_urls - len(self.url_queue)
                self._append_status(f"📄 开始处理第 {task_num} / {self.total_urls} 个任务: {current_url}")
                success = self._run_print_job_for_url(driver, current_url)
                if not success:
                    self._append_status(f"⚠️ 第 {task_num} 个任务处理失败，跳过...")
                    logger.warning(f"任务 {current_url} 处理失败，已跳过。")
                # 不 sleep，直接下一个（或加极短延迟防卡）
                self.root.update()  # 保持 GUI 响应
            self._append_status("🎉 全部任务完成！请在桌面 'AutoGeneratePDF' 文件夹中查看结果。")
            messagebox.showinfo("成功", f"所有 {self.total_urls} 个打印任务已处理完毕！")
        except Exception as e:
            logger.error(f"全局任务异常: {e}", exc_info=True)
            self._append_status(f"❌ 全局错误：{e}")
        finally:
            if driver:
                driver.quit()
                self._append_status("浏览器已关闭。")
            self.start_btn.config(state="normal")
            self.add_url_button.config(state="normal")

    def _setup_driver(self):
        """ 配置并返回一个 Microsoft Edge WebDriver 实例 (为原生PDF打印优化) """
        self._append_status("配置 Edge 浏览器...")
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

        self._append_status(f"使用本地 Edge WebDriver: {driver_path}")
        service = Service(executable_path=str(driver_path))

        self._append_status("启动 Edge 浏览器...")
        return webdriver.Edge(service=service, options=options)

    def _process_single_language(self, driver, btn_text, lang_tag, download_dir):
        self._append_status(f"正在处理：{btn_text}")
        wait = WebDriverWait(driver, Config.SELENIUM_TIMEOUT)
        try:
            lang_button_xpath = f"//button[.//span[contains(text(), '{btn_text}')]]"
            lang_button = wait.until(EC.element_to_be_clickable((By.XPATH, lang_button_xpath)))
            lang_button.click()
            time.sleep(2.5)
            page_title = driver.title
            base_filename = _clean_filename(page_title)
            self._append_status(f"获取到原始文件名: {base_filename}")
            if not base_filename:
                base_filename = f"未命名文档_{datetime.now().strftime('%H%M%S')}"
            print_options = {
                'landscape': False, 'displayHeaderFooter': False,
                'printBackground': True, 'preferCSSPageSize': True,
            }
            self._append_status("正在执行原生PDF生成命令...")
            result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
            pdf_data = base64.b64decode(result['data'])
            new_filename = f"{base_filename}_{lang_tag}.pdf"
            full_file_path = os.path.join(download_dir, new_filename)
            unique_file_path = get_unique_filepath(full_file_path)
            with open(unique_file_path, 'wb') as f:
                f.write(pdf_data)
            saved_name = os.path.basename(unique_file_path)
            self._append_status(f"✅ 成功保存：{saved_name}")
            return True
        except Exception as e:
            self._append_status(f"❌ 处理 '{btn_text}' 时失败: {e}")
            logger.error(f"处理 '{btn_text}' 时失败: {e}", exc_info=True)
            return False

    def _run_print_job_for_url(self, driver, url):
        """在已有 driver 上处理单个 URL"""
        try:
            download_dir = create_date_folder_on_desktop()
            if not download_dir:
                return False

            self._append_status(f"正在打开网页：{url}")
            driver.get(url)

            # 等待关键按钮出现（用 WebDriverWait，不用 sleep）
            wait = WebDriverWait(driver, Config.SELENIUM_TIMEOUT)
            first_button_xpath = f"//button[.//span[contains(text(), '{Config.LANGUAGE_BUTTONS[0][0]}')]]"
            wait.until(EC.visibility_of_element_located((By.XPATH, first_button_xpath)))
            self._append_status("页面加载完成。")

            # 预打印（可选，若非必要可删除）
            try:
                print_options = {
                    'landscape': False, 'displayHeaderFooter': False,
                    'printBackground': True, 'preferCSSPageSize': True,
                }
                driver.execute_cdp_cmd("Page.printToPDF", print_options)
                self._append_status("预打印完成。")
            except Exception as e:
                logger.warning(f"预打印失败: {e}")

            # 处理三种打印情况
            all_success = True
            for btn_text, lang_tag in Config.LANGUAGE_BUTTONS:
                success = self._process_single_language(driver, btn_text, lang_tag, download_dir)
                if not success:
                    all_success = False
                # 可考虑移除或缩短 sleep
                time.sleep(0.5)  # 极短延迟，防点击过快
            return all_success
        except Exception as e:
            logger.error(f"处理 {url} 时出错: {e}", exc_info=True)
            return False

    def _show_help(self):
        """弹出富文本操作指南窗口"""
        help_window = tk.Toplevel(self.root)
        help_window.title("📘 操作指南")
        help_window.geometry("550x450")
        help_window.resizable(True, True)
        help_window.transient(self.root)
        help_window.grab_set()

        # 先隐藏窗口，避免闪现左上角
        help_window.withdraw()

        # 主框架
        main_frame = ttk.Frame(help_window, padding=10)
        main_frame.pack(fill="both", expand=True)

        # 创建 Text 和 Scrollbar
        text_widget = tk.Text(
            main_frame,
            wrap="word",
            font=("微软雅黑", 10),
            bg="white",
            relief="sunken",
            padx=10,
            pady=10
        )
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 配置富文本样式（tags）
        text_widget.tag_configure("title", font=("微软雅黑", 14, "bold"), foreground="#1E88E5", spacing3=10)
        text_widget.tag_configure("section", font=("微软雅黑", 11, "bold"), foreground="#333333", spacing2=6,
                                  spacing3=8)
        text_widget.tag_configure("body", font=("微软雅黑", 10), foreground="#000000", lmargin1=20, lmargin2=30)
        text_widget.tag_configure("bullet", font=("微软雅黑", 10), foreground="#000000", lmargin1=20, lmargin2=30)
        text_widget.tag_configure("note", font=("微软雅黑", 10, "bold"), foreground="#D32F2F")
        text_widget.tag_configure("path", font=("Consolas", 9), background="#F5F5F5", relief="groove", borderwidth=1)

        # 插入内容的辅助函数
        def insert_title(text):
            text_widget.insert("end", text + "\n", "title")

        def insert_section(text):
            text_widget.insert("end", text + "\n", "section")

        def insert_body(text):
            text_widget.insert("end", text + "\n", "body")

        def insert_bullet(text):
            text_widget.insert("end", "• " + text + "\n", "bullet")

        def insert_note(text):
            text_widget.insert("end", text + "\n", "note")

        def insert_path(text):
            text_widget.insert("end", text, "path")

        # 插入内容
        insert_title("AutoGeneratePDF 使用指南")

        insert_section("📌 基本流程")
        insert_bullet("点击「✚ 添加网址」可添加多个待处理网页。")
        insert_bullet("每个网址必须以 http:// 或 https:// 开头。")
        insert_bullet("点击「✅ 开始打印」后，程序将自动：")
        insert_body("  - 打开每个网页（使用 Edge 浏览器）")
        insert_body("  - 依次点击“打印中英文”、“打印英文”、“打印中文”三个按钮")
        insert_body("  - 将每个语言版本生成 PDF 并保存到桌面的 AutoGeneratePDF/日期 文件夹中")

        insert_section("📌 注意事项")
        insert_bullet("网页必须包含以下按钮之一（按钮内含 span 文本）：")
        insert_body("  • \"打印中英文\"")
        insert_body("  • \"打印英文\"")
        insert_body("  • \"打印中文\"")
        insert_bullet("首次运行可能较慢（需加载浏览器），请耐心等待。")
        insert_bullet("若某任务失败，程序会跳过并继续处理下一个。")
        insert_bullet("运行时不要关闭窗口，直到出现“全部任务完成”弹窗。")

        insert_section("📌 输出位置")
        insert_body("所有 PDF 文件将保存在：")
        insert_path("桌面 → AutoGeneratePDF → YYMMDD（当天日期文件夹）")
        text_widget.insert("end", "\n\n", "body")

        insert_section("📌 常见问题")
        insert_note("Q: 点击开始后没反应？")
        insert_body("A: 请检查是否安装了 Microsoft Edge 浏览器，并确认 Edge 浏览器版本为：142.0.3595.94 (正式版本) (64 位)。")

        insert_note("Q: 如何查看 Edge 浏览器版本？")
        insert_body("A: 在 Edge 浏览器网址栏输入 edge://settings/help 查看具体版本号。")

        insert_note("Q: 保存的文件名乱码或为空？")
        insert_body("A: 网页标题可能为空，程序会自动生成时间戳文件名。\n")

        insert_body("如有问题，请联系作者：")
        insert_note("springleaf")

        text_widget.config(state="disabled")

        # 关闭按钮
        close_btn = ttk.Button(help_window, text="关闭", command=help_window.destroy)
        close_btn.pack(pady=10)

        # ===== 居中逻辑 =====
        def center_and_show():
            help_window.update_idletasks()  # 强制更新布局
            width = help_window.winfo_width()
            height = help_window.winfo_height()

            # 确保最小尺寸
            width = max(width, 550)
            height = max(height, 450)

            screen_width = help_window.winfo_screenwidth()
            screen_height = help_window.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2

            help_window.geometry(f"{width}x{height}+{x}+{y}")
            help_window.deiconify()  # 此时才显示窗口

        # 立即执行（无需延迟，因为窗口已构建完成）
        center_and_show()

if __name__ == "__main__":
    root = tk.Tk()
    app = PrintToolApp(root)
    root.mainloop()