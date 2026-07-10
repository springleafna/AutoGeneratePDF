# -*- coding: utf-8 -*-
"""
AutoGeneratePDF - 图形界面版
作者：springleaf
用途：唤唤专用
"""

import base64
import logging
import os
import re
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass, field
from datetime import datetime
from tkinter import ttk, messagebox
from typing import Optional, List, Tuple, Any

from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ========== 异常定义 ==========

class PrintTaskError(Exception):
    """打印任务基础异常"""
    pass


class PageLoadError(PrintTaskError):
    """页面加载失败异常"""
    pass


class ButtonNotFoundError(PrintTaskError):
    """按钮未找到异常"""
    pass


class PDFGenerationError(PrintTaskError):
    """PDF生成失败异常"""
    pass


class FolderCreationError(PrintTaskError):
    """文件夹创建失败异常"""
    pass


class DriverSetupError(PrintTaskError):
    """浏览器驱动初始化失败异常"""
    pass


# ========== 配置区 ==========

@dataclass(frozen=True)
class UIConfig:
    """UI界面配置"""
    WINDOW_TITLE: str = "AutoGeneratePDF"
    WINDOW_WIDTH: int = 650
    WINDOW_HEIGHT: int = 680  # 增加高度以容纳进度条
    WINDOW_RESIZABLE: bool = True
    MAIN_FRAME_PADDING: int = 20
    FONT_FAMILY: str = "微软雅黑"
    FONT_SIZE_TITLE: int = 16
    FONT_SIZE_NORMAL: int = 10
    FONT_SIZE_SMALL: int = 9
    URL_ENTRY_WIDTH: int = 60
    STATUS_TEXT_HEIGHT: int = 5  # 减少高度，为进度条腾出空间


@dataclass(frozen=True)
class BrowserConfig:
    """浏览器配置"""
    DRIVER_FILENAME: str = "msedgedriver.exe"
    SELENIUM_TIMEOUT: int = 20
    PAGE_LOAD_WAIT: float = 2.5  # 点击按钮后等待页面响应的时间
    BETWEEN_CLICKS_DELAY: float = 0.5  # 两次点击之间的延迟
    WINDOW_SIZE: str = "1920,1080"
    HEADLESS: bool = True
    DISABLE_GPU: bool = True
    DISABLE_EXTENSIONS: bool = True
    NO_SANDBOX: bool = True
    DISABLE_DEV_SHM_USAGE: bool = True


@dataclass(frozen=True)
class PathConfig:
    """路径配置"""
    BASE_FOLDER_NAME: str = "AutoGeneratePDF"
    DATE_FORMAT: str = "%y%m%d"
    TIME_FORMAT: str = "%H%M%S"
    DEFAULT_FILENAME_PREFIX: str = "未命名文档"
    DESKTOP_RELATIVE_PATH: str = "Desktop"


@dataclass
class PrintConfig:
    """打印配置"""
    LANGUAGE_BUTTONS: Tuple[Tuple[str, str], ...] = field(
        default_factory=lambda: (
            ("打印中英文", "中英文"),
            ("打印英文", "英文"),
            ("打印中文", "中文"),
            ("打印音标", "音标")
        )
    )
    MAX_RETRIES: int = 2  # 单个任务最大重试次数

    @staticmethod
    def get_pdf_options() -> dict:
        """获取PDF打印选项"""
        return {
            'landscape': False,
            'displayHeaderFooter': False,
            'printBackground': True,
            'preferCSSPageSize': True,
        }


class Config:
    """统一配置管理类"""
    UI = UIConfig()
    BROWSER = BrowserConfig()
    PATH = PathConfig()
    PRINT = PrintConfig()


# ========== 日志配置 ==========

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ========== 工具函数 ==========

def resource_path(relative_path: str) -> str:
    """
    获取资源的绝对路径，支持 PyInstaller 打包环境

    Args:
        relative_path: 相对路径

    Returns:
        绝对路径
    """
    if hasattr(sys, '_MEIPASS'):
        # noqa: SLF001 - _MEIPASS 是 PyInstaller 的标准属性，需要访问
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def clean_filename(name: str) -> str:
    """
    清理文件名，移除非法字符

    Args:
        name: 原始文件名

    Returns:
        清理后的文件名
    """
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()


def get_unique_filepath(base_path: str) -> str:
    """
    确保文件路径唯一：如果文件已存在，则在文件名后添加 (1), (2), ...

    Args:
        base_path: 基础文件路径

    Returns:
        唯一的文件路径

    Examples:
        >>> get_unique_filepath("test.pdf")  # 文件不存在时返回原路径
        'test.pdf'
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


def get_save_folder_path() -> str:
    """
    获取保存文件夹的完整路径（不创建文件夹）

    Returns:
        保存文件夹的完整路径
    """
    desktop = os.path.join(os.path.expanduser("~"), Config.PATH.DESKTOP_RELATIVE_PATH)
    return os.path.join(desktop, Config.PATH.BASE_FOLDER_NAME)


def create_date_folder_on_desktop() -> Optional[str]:
    """
    在桌面创建日期文件夹

    Returns:
        日期文件夹路径，失败返回 None

    Raises:
        FolderCreationError: 文件夹创建失败
    """
    try:
        desktop = os.path.join(os.path.expanduser("~"), Config.PATH.DESKTOP_RELATIVE_PATH)
        base_folder = os.path.join(desktop, Config.PATH.BASE_FOLDER_NAME)
        os.makedirs(base_folder, exist_ok=True)

        date_str = datetime.now().strftime(Config.PATH.DATE_FORMAT)
        date_folder = os.path.join(base_folder, date_str)
        os.makedirs(date_folder, exist_ok=True)

        logger.info(f"✅ 文件将保存至：{date_folder}")
        return date_folder
    except OSError as e:
        error_msg = f"创建文件夹失败: {e}"
        logger.error(error_msg)
        raise FolderCreationError(error_msg) from e


# ========== 主应用类 ==========

class PrintToolApp:
    """PDF打印工具主应用类"""

    def __init__(self, root: tk.Tk) -> None:
        """
        初始化应用

        Args:
            root: Tkinter根窗口
        """
        self.root = root
        self._setup_window()

        # 状态变量
        self.url_entries: List[ttk.Entry] = []
        self.url_queue: List[str] = []
        self.total_urls: int = 0
        self.is_running: bool = False
        self.driver: Optional[webdriver.Edge] = None

        # 进度相关
        self.completed_tasks: int = 0
        self.total_tasks: int = 0  # URL数量 × 语言数量

        self._setup_ui()
        self._show_centered_window()

    def _setup_window(self) -> None:
        """配置窗口基本属性"""
        self.root.title(Config.UI.WINDOW_TITLE)
        self.root.geometry(
            f"{Config.UI.WINDOW_WIDTH}x{Config.UI.WINDOW_HEIGHT}"
        )
        self.root.resizable(
            Config.UI.WINDOW_RESIZABLE,
            Config.UI.WINDOW_RESIZABLE
        )
        self.root.withdraw()  # 隐藏窗口直到居中完成

    def _show_centered_window(self) -> None:
        """居中并显示窗口"""
        self.center_window()
        self.root.after(100, lambda *_args: self._reveal_window())

    def _reveal_window(self) -> None:
        """显示窗口并获取焦点"""
        self.root.deiconify()
        self.root.focus_force()

    def center_window(self) -> None:
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"+{x}+{y}")

    # ==================== UI 组件 ====================

    def _setup_ui(self) -> None:
        """设置用户界面"""
        main_frame = ttk.Frame(self.root, padding=Config.UI.MAIN_FRAME_PADDING)
        main_frame.pack(fill="both", expand=True)

        # 标题
        self._create_title_label(main_frame)

        # 网址列表区域
        self._create_url_area(main_frame)

        # 进度条区域
        self._create_progress_area(main_frame)

        # 按钮区域
        self._create_button_area(main_frame)

        # 状态文本区域
        self._create_status_area(main_frame)

        # 初始化
        self._append_status("请添加网址后开始任务...")
        self._add_url_entry(is_first=True)

    def _create_title_label(self, parent: ttk.Frame) -> None:
        """创建标题标签"""
        title_label = ttk.Label(
            parent,
            text="📚 AutoGeneratePDF",
            font=(
                Config.UI.FONT_FAMILY,
                Config.UI.FONT_SIZE_TITLE,
                "bold"
            )
        )
        title_label.pack(pady=(0, 10))

    def _create_url_area(self, parent: ttk.Frame) -> None:
        """创建网址输入区域"""
        url_area_frame = ttk.LabelFrame(parent, text=" 网址列表 ", padding=10)
        url_area_frame.pack(fill="x", pady=10)

        add_btn = ttk.Button(
            url_area_frame,
            text="✚ 添加网址",
            command=self._add_url_entry
        )
        add_btn.pack(anchor="w", pady=(0, 10))
        self.add_url_button = add_btn

        canvas = tk.Canvas(
            url_area_frame,
            borderwidth=0,
            background="#ffffff"
        )
        self.url_list_frame = ttk.Frame(canvas)

        scrollbar = ttk.Scrollbar(
            url_area_frame,
            orient="vertical",
            command=canvas.yview
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((4, 4), window=self.url_list_frame, anchor="nw")

        self.url_list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

    def _create_progress_area(self, parent: ttk.Frame) -> None:
        """创建进度条区域"""
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill="x", pady=(10, 5))

        # 进度标签
        self.progress_label = ttk.Label(progress_frame, text="准备就绪")
        self.progress_label.pack(anchor="w")

        # 进度条
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode="determinate",
            maximum=100
        )
        self.progress_bar.pack(fill="x", pady=(5, 0))

    def _create_button_area(self, parent: ttk.Frame) -> None:
        """创建按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=10)

        self.start_btn = ttk.Button(
            button_frame,
            text="✅ 开始打印",
            command=self.start_printing_all
        )
        self.start_btn.pack(side="left", padx=5)

        self.exit_btn = ttk.Button(
            button_frame,
            text="❌ 退出",
            command=self._quit_application
        )
        self.exit_btn.pack(side="left", padx=5)

        self.open_folder_btn = ttk.Button(
            button_frame,
            text="📂 打开文件夹",
            command=self._open_save_folder
        )
        self.open_folder_btn.pack(side="left", padx=5)

        self.help_btn = ttk.Button(
            button_frame,
            text="📘 操作指南",
            command=self._show_help
        )
        self.help_btn.pack(side="left", padx=5)

    def _create_status_area(self, parent: ttk.Frame) -> None:
        """创建状态文本区域"""
        self.status_text = tk.Text(
            parent,
            height=Config.UI.STATUS_TEXT_HEIGHT,
            width=80,
            font=(
                Config.UI.FONT_FAMILY,
                Config.UI.FONT_SIZE_SMALL
            ),
            bg="white",
            fg="gray",
            state="disabled",
            relief="sunken"
        )
        self.status_text.pack(pady=(10, 0), fill="x")

    # ==================== 右键粘贴 ====================

    def _paste_on_right_click(self, event: tk.Event) -> str:
        """
        右键点击输入框时，直接将剪贴板内容粘贴到光标位置

        Args:
            event: 鼠标事件
        """
        widget = event.widget
        try:
            self.root.clipboard_get()
        except tk.TclError:
            return "break"

        try:
            widget.focus_set()
            widget.event_generate("<<Paste>>")
        except Exception as e:
            logger.warning(f"右键粘贴失败: {e}")
        return "break"

    # ==================== 网址管理 ====================

    def _add_url_entry(self, is_first: bool = False) -> None:
        """
        添加一个网址输入框

        Args:
            is_first: 是否是第一个输入框（第一个不显示删除按钮）
        """
        row_frame = ttk.Frame(self.url_list_frame)
        row_frame.pack(fill="x", pady=2)

        entry = ttk.Entry(
            row_frame,
            width=Config.UI.URL_ENTRY_WIDTH,
            font=(Config.UI.FONT_FAMILY, Config.UI.FONT_SIZE_NORMAL)
        )
        entry.pack(side="left", expand=True, fill="x")

        # 右键直接粘贴剪贴板内容（Windows/Linux 为 <Button-3>，macOS 为 <Button-2>）
        entry.bind("<Button-3>", self._paste_on_right_click)
        entry.bind("<Button-2>", self._paste_on_right_click)

        self.url_entries.append(entry)

        if not is_first:
            remove_btn = ttk.Button(
                row_frame,
                text="-",
                width=3,
                command=lambda: self._remove_url_entry(row_frame, entry)
            )
            remove_btn.pack(side="left", padx=(5, 0))

    def _remove_url_entry(
        self,
        frame_to_remove: ttk.Frame,
        entry_to_remove: ttk.Entry
    ) -> None:
        """
        移除一个网址输入框

        Args:
            frame_to_remove: 要移除的框架
            entry_to_remove: 要移除的输入框
        """
        frame_to_remove.destroy()
        self.url_entries.remove(entry_to_remove)

    # ==================== 状态更新 ====================

    def _append_status(self, message: str) -> None:
        """
        线程安全的日志更新方法

        Args:
            message: 要显示的消息
        """
        def _update(*_args: Any) -> None:
            self.status_text.config(state="normal")
            self.status_text.insert(tk.END, f"{message}\n")
            self.status_text.see(tk.END)
            self.status_text.config(state="disabled")

        self.root.after(0, _update)

    def _update_progress(
        self,
        current: int,
        total: int,
        message: str = ""
    ) -> None:
        """
        更新进度条

        Args:
            current: 当前进度
            total: 总任务数
            message: 进度描述信息
        """
        def _update_ui(*_args: Any) -> None:
            if total > 0:
                percentage = (current / total) * 100
                self.progress_bar["value"] = percentage

                if message:
                    self.progress_label.config(
                        text=f"{message} ({current}/{total})"
                    )
                else:
                    self.progress_label.config(
                        text=f"进度: {current}/{total} ({percentage:.1f}%)"
                    )
            else:
                self.progress_bar["value"] = 0
                self.progress_label.config(text=message or "准备就绪")

        self.root.after(0, _update_ui)

    def _reset_progress(self) -> None:
        """重置进度条"""
        self.progress_bar["value"] = 0
        self.progress_label.config(text="准备就绪")

    # ==================== 任务控制 ====================

    def start_printing_all(self) -> None:
        """开始打印所有URL"""
        urls = [entry.get().strip() for entry in self.url_entries if entry.get().strip()]

        if not urls:
            messagebox.showerror("错误", "请输入至少一个网址！")
            return

        # 验证URL格式
        for url in urls:
            if not url.startswith(("http://", "https://")):
                messagebox.showerror("错误", f"网址格式不正确：\n{url}")
                return

        # 初始化任务状态
        self.url_queue = urls
        self.total_urls = len(urls)
        self.total_tasks = self.total_urls * len(Config.PRINT.LANGUAGE_BUTTONS)
        self.completed_tasks = 0

        # 禁用按钮
        self.start_btn.config(state="disabled")
        self.add_url_button.config(state="disabled")

        self._append_status(f"🚀 任务队列已创建，共 {self.total_urls} 个URL，{self.total_tasks} 个打印任务。")
        self._update_progress(0, self.total_tasks, "准备中")

        # 启动后台线程
        self.is_running = True
        thread = threading.Thread(target=self._process_all_urls, daemon=True)
        thread.start()

    def _process_all_urls(self) -> None:
        """处理所有URL（在后台线程运行）"""
        try:
            self.driver = self._setup_driver()
            queue_copy = list(self.url_queue)

            for index, url in enumerate(queue_copy):
                if not self.is_running:
                    self._append_status("⏸️ 任务已手动停止")
                    break

                task_num = index + 1
                self._append_status(f"📄 [{task_num}/{self.total_urls}] 处理: {url}")

                try:
                    success = self._run_print_job_for_url(self.driver, url)
                    if success:
                        self._append_status(f"✅ [{task_num}/{self.total_urls}] 完成: {url}")
                    else:
                        self._append_status(f"⚠️ [{task_num}/{self.total_urls}] 部分失败: {url}")
                except PageLoadError as e:
                    self._append_status(f"❌ [{task_num}/{self.total_urls}] 页面加载失败: {e}")
                except Exception as e:
                    self._append_status(f"❌ [{task_num}/{self.total_urls}] 未知错误: {e}")
                    logger.error(f"处理URL出错: {e}", exc_info=True)

            if self.is_running:
                self._append_status("🎉 全部任务完成！请在桌面 'AutoGeneratePDF' 文件夹中查看结果。")
                self._update_progress(self.total_tasks, self.total_tasks, "全部完成")
                self.root.after(0, lambda *_args: messagebox.showinfo("成功", "所有任务已处理完毕！"))

        except DriverSetupError as e:
            self._append_status(f"❌ 浏览器初始化失败: {e}")
            self.root.after(0, lambda *_args: messagebox.showerror("错误", f"浏览器初始化失败:\n{e}"))
        except Exception as e:
            logger.error(f"全局任务异常: {e}", exc_info=True)
            self._append_status(f"❌ 全局错误：{e}")
        finally:
            self._cleanup_resources()
            self.root.after(0, self._reset_ui_state)

    def _run_print_job_for_url(self, driver: webdriver.Edge, url: str) -> bool:
        """
        在已有driver上处理单个URL

        Args:
            driver: WebDriver实例
            url: 要处理的URL

        Returns:
            是否全部成功
        """
        retry_count = 0
        max_retries = Config.PRINT.MAX_RETRIES

        while retry_count <= max_retries:
            try:
                download_dir = create_date_folder_on_desktop()

                self._append_status(f"正在打开网页：{url}")
                driver.get(url)

                # 等待页面加载
                wait = WebDriverWait(driver, Config.BROWSER.SELENIUM_TIMEOUT)
                first_button_xpath = (
                    f"//button[contains(normalize-space(.), "
                    f"'{Config.PRINT.LANGUAGE_BUTTONS[0][0]}')]"
                )

                try:
                    wait.until(EC.visibility_of_element_located((By.XPATH, first_button_xpath)))
                except TimeoutException:
                    raise PageLoadError(f"页面加载超时或未找到打印按钮: {url}")

                # 确认页面完全加载
                driver.execute_script("return document.readyState == 'complete'")
                self._append_status("页面加载完成")

                # 处理各语言打印
                all_success = True
                for btn_text, lang_tag in Config.PRINT.LANGUAGE_BUTTONS:
                    try:
                        success = self._process_single_language(
                            driver, btn_text, lang_tag, download_dir
                        )
                        if success:
                            self.completed_tasks += 1
                            self._update_progress(
                                self.completed_tasks,
                                self.total_tasks,
                                f"处理中"
                            )
                        else:
                            all_success = False
                    except ButtonNotFoundError:
                        self._append_status(f"❌ 未找到按钮: {btn_text}")
                        all_success = False
                    except PDFGenerationError as e:
                        self._append_status(f"❌ PDF生成失败 ({lang_tag}): {e}")
                        all_success = False

                    time.sleep(Config.BROWSER.BETWEEN_CLICKS_DELAY)

                return all_success

            except PageLoadError:
                retry_count += 1
                if retry_count > max_retries:
                    self._append_status(f"❌ 页面加载失败，已重试 {max_retries} 次，跳过此任务")
                    return False
                self._append_status(f"⚠️ 页面加载失败，正在重试 ({retry_count}/{max_retries})...")
                time.sleep(2)

            except FolderCreationError:
                return False

        # 理论上不会到达这里，但为了类型检查器添加默认返回值
        return False

    def _process_single_language(
        self,
        driver: webdriver.Edge,
        btn_text: str,
        lang_tag: str,
        download_dir: str
    ) -> bool:
        """
        处理单个语言版本的打印

        Args:
            driver: WebDriver实例
            btn_text: 按钮文本
            lang_tag: 语言标签
            download_dir: 下载目录

        Returns:
            是否成功

        Raises:
            ButtonNotFoundError: 按钮未找到
            PDFGenerationError: PDF生成失败
        """
        self._append_status(f"  → 处理：{btn_text}")
        wait = WebDriverWait(driver, Config.BROWSER.SELENIUM_TIMEOUT)

        try:
            # 查找并点击按钮
            lang_button_xpath = (
                f"//button[contains(normalize-space(.), '{btn_text}')]"
            )
            lang_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, lang_button_xpath))
            )
            lang_button.click()

            # 等待页面响应
            time.sleep(Config.BROWSER.PAGE_LOAD_WAIT)

            # 获取页面标题作为文件名
            page_title = driver.title
            base_filename = clean_filename(page_title)

            if not base_filename:
                base_filename = f"{Config.PATH.DEFAULT_FILENAME_PREFIX}_{datetime.now().strftime(Config.PATH.TIME_FORMAT)}"

            self._append_status(f"  → 文件名：{base_filename}_{lang_tag}.pdf")

            # 生成PDF
            result = driver.execute_cdp_cmd("Page.printToPDF", PrintConfig.get_pdf_options())
            pdf_data = base64.b64decode(result['data'])

            # 保存文件
            filename = f"{base_filename}_{lang_tag}.pdf"
            full_path = os.path.join(download_dir, filename)
            unique_path = get_unique_filepath(full_path)

            with open(unique_path, 'wb') as f:
                f.write(pdf_data)

            saved_name = os.path.basename(unique_path)
            self._append_status(f"  ✅ 已保存：{saved_name}")
            return True

        except TimeoutException:
            raise ButtonNotFoundError(f"未找到 '{btn_text}' 按钮")
        except WebDriverException as e:
            raise PDFGenerationError(f"PDF生成失败: {e}")
        except IOError as e:
            raise PDFGenerationError(f"文件保存失败: {e}")

    # ==================== 浏览器管理 ====================

    def _setup_driver(self) -> webdriver.Edge:
        """
        配置并返回Edge WebDriver实例

        Returns:
            配置好的WebDriver实例

        Raises:
            DriverSetupError: 驱动初始化失败
        """
        try:
            self._append_status("配置 Edge 浏览器...")
            options = Options()

            if Config.BROWSER.HEADLESS:
                options.add_argument("--headless")
            if Config.BROWSER.DISABLE_GPU:
                options.add_argument("--disable-gpu")
            if Config.BROWSER.DISABLE_EXTENSIONS:
                options.add_argument("--disable-extensions")
            if Config.BROWSER.NO_SANDBOX:
                options.add_argument("--no-sandbox")
            if Config.BROWSER.DISABLE_DEV_SHM_USAGE:
                options.add_argument("--disable-dev-shm-usage")

            options.add_argument(f"--window-size={Config.BROWSER.WINDOW_SIZE}")
            options.add_experimental_option('excludeSwitches', ['enable-logging'])

            driver_path = resource_path(Config.BROWSER.DRIVER_FILENAME)
            self._append_status(f"使用驱动: {driver_path}")

            service = Service(executable_path=str(driver_path))
            self._append_status("启动 Edge 浏览器...")

            return webdriver.Edge(service=service, options=options)

        except WebDriverException as e:
            error_msg = f"Edge WebDriver初始化失败: {e}"
            logger.error(error_msg)
            raise DriverSetupError(error_msg) from e
        except Exception as e:
            error_msg = f"未知错误: {e}"
            logger.error(error_msg)
            raise DriverSetupError(error_msg) from e

    def _cleanup_resources(self) -> None:
        """清理资源（关闭浏览器等）"""
        if self.driver:
            try:
                self.driver.quit()
                self._append_status("浏览器已关闭")
            except Exception as e:
                logger.error(f"关闭浏览器时出错: {e}")
            finally:
                self.driver = None

    def _reset_ui_state(self) -> None:
        """恢复界面按钮状态"""
        self.start_btn.config(state="normal")
        self.add_url_button.config(state="normal")
        self.is_running = False

    def _quit_application(self) -> None:
        """退出应用"""
        self.is_running = False
        self._cleanup_resources()
        self.root.quit()

    # ==================== 其他功能 ====================

    def _open_save_folder(self) -> None:
        """打开保存PDF的文件夹"""
        try:
            base_folder = get_save_folder_path()
            if os.path.exists(base_folder):
                os.startfile(base_folder)
            else:
                messagebox.showinfo(
                    "提示",
                    f"文件夹尚未创建\n首次打印时会自动创建：\n{base_folder}"
                )
        except Exception as e:
            messagebox.showerror("错误", f"打开文件夹失败：{e}")

    def _show_help(self) -> None:
        """弹出操作指南窗口"""
        help_window = tk.Toplevel(self.root)
        help_window.title("📘 操作指南")
        help_window.geometry("550x450")
        help_window.resizable(True, True)
        help_window.transient(self.root)
        help_window.grab_set()
        help_window.withdraw()

        main_frame = ttk.Frame(help_window, padding=10)
        main_frame.pack(fill="both", expand=True)

        text_widget = tk.Text(
            main_frame,
            wrap="word",
            font=(Config.UI.FONT_FAMILY, Config.UI.FONT_SIZE_NORMAL),
            bg="white",
            relief="sunken",
            padx=10,
            pady=10
        )
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 配置文本样式
        text_widget.tag_configure(
            "title",
            font=(Config.UI.FONT_FAMILY, 14, "bold"),
            foreground="#1E88E5",
            spacing3=10
        )
        text_widget.tag_configure(
            "section",
            font=(Config.UI.FONT_FAMILY, 11, "bold"),
            foreground="#333333",
            spacing2=6,
            spacing3=8
        )
        text_widget.tag_configure(
            "body",
            font=(Config.UI.FONT_FAMILY, Config.UI.FONT_SIZE_NORMAL),
            foreground="#000000",
            lmargin1=20,
            lmargin2=30
        )
        text_widget.tag_configure(
            "bullet",
            font=(Config.UI.FONT_FAMILY, Config.UI.FONT_SIZE_NORMAL),
            foreground="#000000",
            lmargin1=20,
            lmargin2=30
        )
        text_widget.tag_configure(
            "note",
            font=(Config.UI.FONT_FAMILY, Config.UI.FONT_SIZE_NORMAL, "bold"),
            foreground="#D32F2F"
        )
        text_widget.tag_configure(
            "path",
            font=("Consolas", 9),
            background="#F5F5F5",
            relief="groove",
            borderwidth=1
        )

        # 插入帮助内容
        text_widget.insert("end", "AutoGeneratePDF 使用指南\n", "title")

        text_widget.insert("end", "📌 基本流程\n", "section")
        text_widget.insert("end", "• 点击「✚ 添加网址」可添加多个待处理网页。\n", "bullet")
        text_widget.insert("end", "• 每个网址必须以 http:// 或 https:// 开头。\n", "bullet")
        text_widget.insert("end", "• 点击「✅ 开始打印」后，程序将自动：\n", "bullet")
        text_widget.insert("end", "  - 打开每个网页（使用 Edge 浏览器）\n", "body")
        text_widget.insert("end", "  - 依次点击‘打印中英文’、‘打印英文’、‘打印中文’、‘打印音标’四个按钮\n", "body")
        text_widget.insert("end", "  - 直接导出对应 PDF，无需手动点击‘在线打印’\n", "body")
        text_widget.insert("end", "  - 将每个语言版本生成 PDF 并保存到桌面的 AutoGeneratePDF/日期 文件夹中\n", "body")

        text_widget.insert("end", "\n📌 注意事项\n", "section")
        text_widget.insert("end", "• 网页必须包含打印按钮（打印中英文/打印英文/打印中文/打印音标）\n", "bullet")
        text_widget.insert("end", "• 首次运行可能较慢（需加载浏览器），请耐心等待\n", "bullet")
        text_widget.insert("end", "• 若某任务失败，程序会自动重试2次后跳过\n", "bullet")
        text_widget.insert("end", "• 运行时不要关闭窗口，直到出现‘全部任务完成’弹窗\n", "bullet")
        text_widget.insert("end", "• 如果浏览器启动失败，请检查 msedgedriver.exe 是否和 Edge 浏览器版本匹配\n", "bullet")

        text_widget.insert("end", "\n📌 输出位置\n", "section")
        text_widget.insert("end", "所有 PDF 文件将保存在：\n", "body")
        text_widget.insert("end", "桌面 → AutoGeneratePDF → YYMMDD（当天日期文件夹）", "path")
        text_widget.insert("end", "\n\n", "body")
        text_widget.insert("end", "每个网址通常会生成四个文件：中英文、英文、中文、音标。\n", "body")
        text_widget.insert("end", "如果同名文件已存在，程序会自动追加 (1)、(2) 等编号，避免覆盖旧文件。\n", "body")

        text_widget.insert("end", "📌 常见问题\n", "section")
        text_widget.insert("end", "Q: 点击开始后没反应？\n", "note")
        text_widget.insert("end", "A: 请检查是否安装了 Microsoft Edge 浏览器\n\n", "body")

        text_widget.insert("end", "Q: 如何查看 Edge 浏览器版本？\n", "note")
        text_widget.insert("end", "A: 在 Edge 浏览器网址栏输入 edge://settings/help\n\n", "body")

        text_widget.insert("end", "Q: 为什么只能生成部分 PDF？\n", "note")
        text_widget.insert("end", "A: 请确认网页四个打印按钮都能正常显示，并保持网络连接稳定\n\n", "body")

        text_widget.insert("end", "如有问题，请联系作者：springleaf\n", "note")

        text_widget.config(state="disabled")

        # 关闭按钮
        close_btn = ttk.Button(help_window, text="关闭", command=help_window.destroy)
        close_btn.pack(pady=10)

        # 居中并显示
        def center_and_show():
            help_window.update_idletasks()
            width = max(help_window.winfo_width(), 550)
            height = max(help_window.winfo_height(), 450)
            screen_width = help_window.winfo_screenwidth()
            screen_height = help_window.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            help_window.geometry(f"{width}x{height}+{x}+{y}")
            help_window.deiconify()

        center_and_show()


# ========== 主程序入口 ==========

if __name__ == "__main__":
    root = tk.Tk()
    app = PrintToolApp(root)
    root.mainloop()
