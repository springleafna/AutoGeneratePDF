# -*- coding: utf-8 -*-
"""
AutoGeneratePDF - 图形界面版
作者：springleaf
用途：唤唤专用
"""

import logging
import os
import re
import sys
import threading
import time
import tkinter as tk
import zipfile
from dataclasses import dataclass
from datetime import datetime
from tkinter import ttk, messagebox
from typing import Optional, List, Any

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


class DownloadTimeoutError(PrintTaskError):
    """压缩包下载超时异常"""
    pass


class ZIPExtractionError(PrintTaskError):
    """压缩包解压失败异常"""
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
    DESKTOP_RELATIVE_PATH: str = "Desktop"


@dataclass(frozen=True)
class PrintConfig:
    """打印配置"""
    EXPORT_BUTTON_TEXT: str = "一键导出PDF"  # 网页上一键导出按钮的文字
    EXPECTED_PDF_COUNT: int = 4  # 每个压缩包内预期的PDF数量（中文/英文/中英文/音标）
    MAX_RETRIES: int = 2  # 单个任务最大重试次数
    DOWNLOAD_TIMEOUT: float = 60.0  # 等待压缩包下载完成的超时时间（秒）
    DOWNLOAD_POLL_INTERVAL: float = 0.5  # 下载轮询间隔（秒）


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
    """获取资源的绝对路径（兼容 PyInstaller 打包环境）"""
    if hasattr(sys, '_MEIPASS'):
        # noqa: SLF001
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def clean_filename(name: str) -> str:
    """清理文件名，移除非法字符"""
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()


def get_unique_filepath(base_path: str) -> str:
    """确保文件路径唯一：已存在则在文件名后追加 (1)、(2) ..."""
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
    """获取保存文件夹路径（不创建）"""
    desktop = os.path.join(os.path.expanduser("~"), Config.PATH.DESKTOP_RELATIVE_PATH)
    return os.path.join(desktop, Config.PATH.BASE_FOLDER_NAME)


def create_date_folder_on_desktop() -> Optional[str]:
    """在桌面创建当天日期文件夹，失败抛出 FolderCreationError"""
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
        """初始化应用"""
        self.root = root
        self._setup_window()

        self.url_entries: List[ttk.Entry] = []
        self.url_queue: List[str] = []
        self.total_urls: int = 0
        self.is_running: bool = False
        self.driver: Optional[webdriver.Edge] = None

        self.completed_tasks: int = 0
        self.total_tasks: int = 0  # URL数 × 每个URL的PDF数

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

        self._create_title_label(main_frame)
        self._create_url_area(main_frame)
        self._create_progress_area(main_frame)
        self._create_button_area(main_frame)
        self._create_status_area(main_frame)

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

        self.progress_label = ttk.Label(progress_frame, text="准备就绪")
        self.progress_label.pack(anchor="w")

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
        """右键点击输入框时，直接将剪贴板内容粘贴到光标位置"""
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
        """添加一个网址输入框（首个不显示删除按钮）"""
        row_frame = ttk.Frame(self.url_list_frame)
        row_frame.pack(fill="x", pady=2)

        entry = ttk.Entry(
            row_frame,
            width=Config.UI.URL_ENTRY_WIDTH,
            font=(Config.UI.FONT_FAMILY, Config.UI.FONT_SIZE_NORMAL)
        )
        entry.pack(side="left", expand=True, fill="x")

        # Windows/Linux 为 <Button-3>，macOS 为 <Button-2>
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
        """移除一个网址输入框"""
        frame_to_remove.destroy()
        self.url_entries.remove(entry_to_remove)

    # ==================== 状态更新 ====================

    def _append_status(self, message: str) -> None:
        """线程安全地向状态栏追加一条日志"""
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
        """更新进度条（线程安全）"""
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

        for url in urls:
            if not url.startswith(("http://", "https://")):
                messagebox.showerror("错误", f"网址格式不正确：\n{url}")
                return

        self.url_queue = urls
        self.total_urls = len(urls)
        self.total_tasks = self.total_urls * Config.PRINT.EXPECTED_PDF_COUNT
        self.completed_tasks = 0

        self.start_btn.config(state="disabled")
        self.add_url_button.config(state="disabled")

        self._append_status(f"🚀 任务队列已创建，共 {self.total_urls} 个URL，{self.total_tasks} 个打印任务。")
        self._update_progress(0, self.total_tasks, "准备中")

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
        """处理单个URL：点击一键导出按钮，下载压缩包并解压（失败自动重试）"""
        retry_count = 0
        max_retries = Config.PRINT.MAX_RETRIES

        while retry_count <= max_retries:
            try:
                download_dir = create_date_folder_on_desktop()

                # 本站为hash路由单页应用，直接切hash不会重新加载文档，
                # 先跳空白页强制重载，避免导出上一位学生的数据
                driver.get("about:blank")

                self._append_status(f"正在打开网页：{url}")
                driver.get(url)

                wait = WebDriverWait(driver, Config.BROWSER.SELENIUM_TIMEOUT)
                button_xpath = self._export_button_xpath()
                try:
                    wait.until(EC.visibility_of_element_located((By.XPATH, button_xpath)))
                except TimeoutException:
                    raise PageLoadError(f"页面加载超时或未找到导出按钮: {url}")

                driver.execute_script("return document.readyState == 'complete'")
                self._append_status("页面加载完成")

                try:
                    driver.execute_cdp_cmd(
                        "Browser.setDownloadBehavior",
                        {"behavior": "allow", "downloadPath": download_dir}
                    )
                except WebDriverException as e:
                    logger.warning(f"设置下载目录失败，将使用浏览器默认下载目录: {e}")

                before_files = set(os.listdir(download_dir))

                export_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, button_xpath))
                )
                export_button.click()
                self._append_status("已点击一键导出，等待压缩包下载...")

                zip_path = self._wait_for_zip_download(download_dir, before_files)
                saved_files = self._extract_zip(zip_path, download_dir)

                self.completed_tasks += len(saved_files)
                self._update_progress(
                    min(self.completed_tasks, self.total_tasks),
                    self.total_tasks,
                    "处理中"
                )
                return True

            except (PageLoadError, ButtonNotFoundError,
                    DownloadTimeoutError, ZIPExtractionError) as e:
                retry_count += 1
                if retry_count > max_retries:
                    self._append_status(f"❌ {e}，已重试 {max_retries} 次，跳过此任务")
                    return False
                self._append_status(f"⚠️ {e}，正在重试 ({retry_count}/{max_retries})...")
                time.sleep(2)

            except FolderCreationError:
                return False

        # 不可达，仅为类型检查器保留
        return False

    @staticmethod
    def _export_button_xpath() -> str:
        """构造一键导出按钮的XPath（兼容button、a及role=button元素）"""
        text = Config.PRINT.EXPORT_BUTTON_TEXT
        return (
            f"(//button[contains(normalize-space(.), '{text}')]"
            f"|//a[contains(normalize-space(.), '{text}')]"
            f"|//*[@role='button'][contains(normalize-space(.), '{text}')])"
        )

    def _wait_for_zip_download(self, download_dir: str, before_files: set) -> str:
        """轮询等待新下载的ZIP完成并返回路径，超时抛 DownloadTimeoutError"""
        deadline = time.time() + Config.PRINT.DOWNLOAD_TIMEOUT
        last_size = -1

        while time.time() < deadline:
            time.sleep(Config.PRINT.DOWNLOAD_POLL_INTERVAL)
            new_files = set(os.listdir(download_dir)) - before_files

            if any(f.endswith((".crdownload", ".tmp", ".part")) for f in new_files):
                last_size = -1
                continue

            zip_files = sorted(
                (f for f in new_files if f.lower().endswith(".zip")),
                key=lambda f: os.path.getmtime(os.path.join(download_dir, f)),
                reverse=True
            )
            if zip_files:
                zip_path = os.path.join(download_dir, zip_files[0])
                try:
                    size = os.path.getsize(zip_path)
                except OSError:
                    continue
                # 大小连续两次轮询稳定视为下载完成
                if size > 0 and size == last_size:
                    return zip_path
                last_size = size

        raise DownloadTimeoutError(
            f"{int(Config.PRINT.DOWNLOAD_TIMEOUT)}秒内未检测到下载完成的压缩包"
        )

    def _extract_zip(self, zip_path: str, download_dir: str) -> List[str]:
        """解压ZIP中的PDF到下载目录（保留原文件名），完成后删除压缩包"""
        saved_files: List[str] = []
        try:
            with zipfile.ZipFile(zip_path) as zf:
                pdf_entries = [
                    name for name in zf.namelist()
                    if not name.endswith('/') and name.lower().endswith('.pdf')
                ]
                if not pdf_entries:
                    raise ZIPExtractionError(
                        f"压缩包 {os.path.basename(zip_path)} 中未找到PDF文件"
                    )

                for entry in pdf_entries:
                    # 只取文件名，防止压缩包内路径逃逸
                    filename = clean_filename(
                        os.path.basename(entry.replace('\\', '/'))
                    )
                    if not filename:
                        continue
                    target_path = get_unique_filepath(
                        os.path.join(download_dir, filename)
                    )
                    with zf.open(entry) as src, open(target_path, 'wb') as dst:
                        dst.write(src.read())
                    saved_files.append(target_path)
                    self._append_status(f"  ✅ 已保存：{os.path.basename(target_path)}")
        except zipfile.BadZipFile as e:
            raise ZIPExtractionError(f"压缩包损坏或不是有效的ZIP文件: {e}") from e
        except (OSError, IOError) as e:
            raise ZIPExtractionError(f"解压保存PDF失败: {e}") from e

        expected = Config.PRINT.EXPECTED_PDF_COUNT
        if len(saved_files) != expected:
            self._append_status(
                f"  ⚠️ 本次解压出 {len(saved_files)} 个PDF，预期 {expected} 个"
            )

        if saved_files:
            try:
                os.remove(zip_path)
                self._append_status(f"  🗑️ 已删除压缩包：{os.path.basename(zip_path)}")
            except OSError as e:
                logger.warning(f"删除压缩包失败: {e}")

        return saved_files

    # ==================== 浏览器管理 ====================

    def _setup_driver(self) -> webdriver.Edge:
        """配置并返回 Edge WebDriver 实例，失败抛 DriverSetupError"""
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

            # 直接下载到日期文件夹，不弹保存对话框
            try:
                download_dir = create_date_folder_on_desktop()
            except FolderCreationError as e:
                raise DriverSetupError(str(e)) from e
            options.add_experimental_option(
                'prefs',
                {
                    'download.default_directory': download_dir,
                    'download.prompt_for_download': False,
                    'plugins.always_open_pdf_externally': True,
                }
            )
            self._append_status(f"下载目录：{download_dir}")

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

        text_widget.insert("end", "AutoGeneratePDF 使用指南\n", "title")

        text_widget.insert("end", "📌 基本流程\n", "section")
        text_widget.insert("end", "• 点击「✚ 添加网址」可添加多个待处理网页。\n", "bullet")
        text_widget.insert("end", "• 每个网址必须以 http:// 或 https:// 开头。\n", "bullet")
        text_widget.insert("end", "• 点击「✅ 开始打印」后，程序将自动：\n", "bullet")
        text_widget.insert("end", "  - 打开每个网页（使用 Edge 浏览器）\n", "body")
        text_widget.insert("end", "  - 点击网页上的「一键导出PDF」按钮\n", "body")
        text_widget.insert("end", "  - 网站会生成包含中文、英文、中英文、音标四个版本的压缩包并自动下载\n", "body")
        text_widget.insert("end", "  - 程序自动解压 PDF 到桌面的 AutoGeneratePDF/日期 文件夹，并删除压缩包\n", "body")

        text_widget.insert("end", "\n📌 注意事项\n", "section")
        text_widget.insert("end", "• 网页必须包含「一键导出PDF」按钮\n", "bullet")
        text_widget.insert("end", "• 首次运行可能较慢（需加载浏览器），请耐心等待\n", "bullet")
        text_widget.insert("end", "• 若某任务失败，程序会自动重试2次后跳过\n", "bullet")
        text_widget.insert("end", "• 运行时不要关闭窗口，直到出现'全部任务完成'弹窗\n", "bullet")
        text_widget.insert("end", "• 如果浏览器启动失败，请检查 msedgedriver.exe 是否和 Edge 浏览器版本匹配\n", "bullet")

        text_widget.insert("end", "\n📌 输出位置\n", "section")
        text_widget.insert("end", "所有 PDF 文件将保存在：\n", "body")
        text_widget.insert("end", "桌面 → AutoGeneratePDF → YYMMDD（当天日期文件夹）", "path")
        text_widget.insert("end", "\n\n", "body")
        text_widget.insert("end", "每个网址通常会解压出四个文件：中文、英文、中英文、音标，文件名与压缩包内一致。\n", "body")
        text_widget.insert("end", "如果同名文件已存在，程序会自动追加 (1)、(2) 等编号，避免覆盖旧文件。\n", "body")

        text_widget.insert("end", "📌 常见问题\n", "section")
        text_widget.insert("end", "Q: 点击开始后没反应？\n", "note")
        text_widget.insert("end", "A: 请检查是否安装了 Microsoft Edge 浏览器\n\n", "body")

        text_widget.insert("end", "Q: 如何查看 Edge 浏览器版本？\n", "note")
        text_widget.insert("end", "A: 在 Edge 浏览器网址栏输入 edge://settings/help\n\n", "body")

        text_widget.insert("end", "Q: 为什么只能生成部分 PDF？\n", "note")
        text_widget.insert("end", "A: 请确认网页「一键导出PDF」按钮能正常显示、压缩包下载完整，并保持网络连接稳定\n\n", "body")

        text_widget.insert("end", "如有问题，请联系作者：springleaf\n", "note")

        text_widget.config(state="disabled")

        close_btn = ttk.Button(help_window, text="关闭", command=help_window.destroy)
        close_btn.pack(pady=10)

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
