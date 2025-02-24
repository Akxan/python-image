#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
优化后的批量图片转换工具

主要改进点：
1. 功能模块化：将配置管理、国际化、图像处理、文件转换和 UI 部分进行分离；
2. 使用 pathlib 进行路径处理，提高跨平台兼容性；
3. 引入多线程处理转换任务，避免阻塞 GUI 主线程；
4. 统一常量与日志配置，减少重复代码；
5. 详细注释每一步骤，方便阅读和后续维护；
6. **新增“一键清除所有文件”功能，及修改预览统计显示格式。**

作者: Julio (优化示例)
版本: 1.0.2
"""

# ================= 导入必要的标准库和第三方库 =================
import os             # 操作系统接口，用于部分系统调用（例如获取扩展名等）
import io             # 内存中的字节流操作
import base64         # 进行Base64编码（用于SVG嵌入图片数据）
import socket         # 网络通信模块，用于单实例检测
import sys            # 系统相关功能，如退出程序
import locale         # 系统区域设置，用于自动获取默认语言
import json           # JSON序列化与反序列化，用于配置管理
import threading      # 线程模块，用于异步执行转换任务
import shutil         # 文件复制、移动等操作
import logging        # 日志记录模块，用于记录调试、信息和错误日志
from pathlib import Path  # 使用Path对象处理文件路径，提高跨平台性

# Tkinter 相关库，用于创建图形用户界面
import tkinter as tk                                 # Tkinter主模块
from tkinter import filedialog, messagebox           # 文件对话框和消息框
import tkinter.ttk as ttk                            # Tkinter 的ttk模块，提供现代化控件

# 拖拽支持模块
from tkinterdnd2 import DND_FILES, TkinterDnD         # 支持拖拽文件进窗口

# 图像处理相关库
from PIL import Image, ImageTk                       # Pillow库：Image用于图像操作，ImageTk用于在Tkinter中显示图片
import fitz                                          # PyMuPDF库，用于处理PDF，将PDF页面转换为图片


# ================= 日志配置模块 =================
logging.basicConfig(
    level=logging.DEBUG,  # 输出DEBUG及以上级别日志
    format="%(asctime)s - %(levelname)s - %(message)s",  # 日志格式：时间-日志级别-消息
    handlers=[logging.StreamHandler(), logging.FileHandler('app.log', encoding='utf-8')]
)
logger = logging.getLogger(__name__)  # 创建logger对象

# ================= 配置管理模块 =================
CONFIG_FILE = Path("config.json")

def load_config():
    """
    从 JSON 配置文件加载配置，若文件不存在或加载失败，返回空字典
    """
    if CONFIG_FILE.exists():
        try:
            with CONFIG_FILE.open("r", encoding="utf-8") as f:
                config = json.load(f)
            return config
        except Exception as e:
            logger.error("Error loading config: " + str(e))
    return {}

def save_config(config):
    """
    将配置保存到 JSON 文件中
    """
    try:
        with CONFIG_FILE.open("w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
    except Exception as e:
        logger.error("Error saving config: " + str(e))

# ================= 常量定义模块 =================
SUPPORTED_EXTS = [
    ".png", ".jpg", ".jpeg", ".jfif", ".bmp", ".gif", ".tiff",
    ".webp", ".ico", ".ppm", ".tga", ".jp2", ".pdf", ".svg", ".heic",
    ".xlsx", ".xls", ".doc", ".docx", ".csv"
]

FORMAT_MAPPING = {
    "JPEG": ("JPEG", ".jpg"),
    "JPG": ("JPEG", ".jpg"),
    "JPGE": ("JPEG", ".jpg"),
    "JFIF": ("JPEG", ".jfif"),
    "PNG": ("PNG", ".png"),
    "BMP": ("BMP", ".bmp"),
    "GIF": ("GIF", ".gif"),
    "TIFF": ("TIFF", ".tiff"),
    "WEBP": ("WEBP", ".webp"),
    "ICO": ("ICO", ".ico"),
    "PPM": ("PPM", ".ppm"),
    "TGA": ("TGA", ".tga"),
    "JPEG2000": ("JPEG2000", ".jp2"),
    "PDF": ("PDF", ".pdf"),
    "SVG": ("SVG", ".svg"),
    "HEIC": ("HEIC", ".heic"),
    "EXCEL": ("EXCEL", ".xlsx"),
    "WORD": ("WORD", ".docx"),
    "CSV": ("CSV", ".csv")
}

OUTPUT_FORMAT_LIST = list(FORMAT_MAPPING.keys())

# ================= 国际化（多语言）模块 =================
translations = {
    'en': {
        'menu_help': "Help",
        'menu_supported_formats': "Supported Formats",
        'menu_about': "About",
        'menu_language': "Language",
        'lang_english': "English",
        'lang_spanish': "Spanish",
        'lang_russian': "Russian",
        'lang_chinese': "Chinese",
        'msg_file_added_success': "Added {n} files.",
        'msg_invalid_file': "No valid image, HEIC, Excel, Word, CSV or PDF files were dropped.",
        'msg_select_files': "Selected {n} files.",
        'msg_select_output_folder': "Output folder: {folder}",
        'msg_warning_no_file': "Please add files first!",
        'msg_warning_no_output': "Please choose an output folder!",
        'msg_convert_complete': "Successfully converted {n} files (or pages).",
        'msg_conflict': "Another instance is already running. Please do not start multiple instances.",
        'label_instruction': "Please drag and drop images below, or use the button to select files",
        'label_drop_area': "Drop files here",
        'btn_select_files': "Select Files",
        'btn_select_output': "Select Output Folder",
        'label_output_format': "Output Format:",
        'btn_convert': "Start Conversion",
        'msg_supported_formats': "Supported input formats:\n{inputs}\n\nSupported output formats:\n{outputs}",
        'msg_about': "Batch Image Converter\n\nAuthor: Julio\nVersion: 1.0.2",
        'label_all_files': "All files: {n}",  # 新增文本：显示所有文件数量
        'btn_clear_all': "Clear All",         # 新增按钮文本：清除所有文件
        'window_title': "Converter BOX"
    },
    'es': {
        'menu_help': "Ayuda",
        'menu_supported_formats': "Formatos Soportados",
        'menu_about': "Acerca de",
        'menu_language': "Idioma",
        'lang_english': "Inglés",
        'lang_spanish': "Español",
        'lang_russian': "Ruso",
        'lang_chinese': "Chino",
        'msg_file_added_success': "Se han añadido {n} archivos.",
        'msg_invalid_file': "Ningún archivo válido fue arrastrado.",
        'msg_select_files': "Se han seleccionado {n} archivos.",
        'msg_select_output_folder': "Carpeta de salida: {folder}",
        'msg_warning_no_file': "¡Por favor, añade archivos primero!",
        'msg_warning_no_output': "¡Por favor, elige una carpeta de salida!",
        'msg_convert_complete': "Se han convertido exitosamente {n} archivos (o páginas).",
        'msg_conflict': "Otra instancia ya está en ejecución. Por favor, no abras múltiples instancias.",
        'label_instruction': "Arrastra y suelta imágenes, abajo, o usa el botón para seleccionar archivos",
        'label_drop_area': "Suelta los archivos aquí",
        'btn_select_files': "Seleccionar Archivos",
        'btn_select_output': "Seleccionar Carpeta de Salida",
        'label_output_format': "Formato de salida:",
        'btn_convert': "Iniciar Conversión",
        'msg_supported_formats': "Formatos de entrada soportados:\n{inputs}\n\nFormatos de salida soportados:\n{outputs}",
        'msg_about': "Conversor de Imágenes por Lotes\n\nAutor: Julio\nVersión: 1.0.2",
        'label_all_files': "Todos los archivos: {n}",
        'btn_clear_all': "Borrar Todo",
        'window_title': "Converter BOX"
    },
    'ru': {
        'menu_help': "Справка",
        'menu_supported_formats': "Поддерживаемые форматы",
        'menu_about': "О программе",
        'menu_language': "Язык",
        'lang_english': "Английский",
        'lang_spanish': "Испанский",
        'lang_russian': "Русский",
        'lang_chinese': "Китайский",
        'msg_file_added_success': "Добавлено {n} файлов.",
        'msg_invalid_file': "Перетащенные файлы не являются допустимыми.",
        'msg_select_files': "Выбрано {n} файлов.",
        'msg_select_output_folder': "Папка вывода: {folder}",
        'msg_warning_no_file': "Пожалуйста, сначала добавьте файлы!",
        'msg_warning_no_output': "Пожалуйста, выберите папку вывода!",
        'msg_convert_complete': "Успешно конвертировано {n} файлов (или страниц).",
        'msg_conflict': "Программа уже запущена. Пожалуйста, не запускайте несколько экземпляров.",
        'label_instruction': "Перетащите изображения или используйте кнопку для выбора файлов",
        'label_drop_area': "Перетащите файлы сюда",
        'btn_select_files': "Выбрать файлы",
        'btn_select_output': "Выбрать папку вывода",
        'label_output_format': "Формат вывода:",
        'btn_convert': "Начать конвертацию",
        'msg_supported_formats': "Поддерживаемые форматы ввода:\n{inputs}\n\nПоддерживаемые форматы вывода:\n{outputs}",
        'msg_about': "Пакетный конвертер изображений\n\nАвтор: Julio\nВерсия: 1.0.2",
        'label_all_files': "Все файлы: {n}",
        'btn_clear_all': "Очистить всё",
        'window_title': "Converter BOX"
    },
    'zh': {
        'menu_help': "帮助",
        'menu_supported_formats': "支持格式",
        'menu_about': "关于",
        'menu_language': "语言",
        'lang_english': "英文",
        'lang_spanish': "西班牙文",
        'lang_russian': "俄语",
        'lang_chinese': "中文",
        'msg_file_added_success': "已添加 {n} 个文件。",
        'msg_invalid_file': "拖入的文件中没有有效文件。",
        'msg_select_files': "已选择 {n} 个文件。",
        'msg_select_output_folder': "输出文件夹：{folder}",
        'msg_warning_no_file': "请先添加文件！",
        'msg_warning_no_output': "请先选择输出文件夹！",
        'msg_convert_complete': "成功转换 {n} 个文件（或页面）。",
        'msg_conflict': "程序已在运行，请勿多次启动。",
        'label_instruction': "请将图片拖拽到下方区域，或使用按钮选择文件",
        'label_drop_area': "将文件拖拽到此处",
        'btn_select_files': "选择文件",
        'btn_select_output': "选择输出文件夹",
        'label_output_format': "输出格式：",
        'btn_convert': "开始转换",
        'msg_supported_formats': "支持的输入格式：\n{inputs}\n\n支持的输出格式：\n{outputs}",
        'msg_about': "批量图片转换工具\n\n作者: Julio\n版本: 1.0.2",
        'label_all_files': "所有文件为: {n}",
        'btn_clear_all': "清除所有",
        'window_title': "转换盒子"
    }
}

supported_langs = ['en', 'es', 'ru', 'zh']
loc = locale.getlocale()
sys_lang = loc[0] if loc[0] else 'en'
default_lang = sys_lang[:2] if sys_lang[:2] in supported_langs else 'en'
config = load_config()
current_lang = config.get("language", default_lang)

def _(key, **kwargs):
    """
    国际化函数，根据当前语言返回对应文本，并支持字符串格式化
    参数:
        key: 文本键
        kwargs: 格式化参数，例如 {n} 或 {folder}
    返回:
        根据当前语言格式化后的文本
    """
    text = translations.get(current_lang, translations['en']).get(key, key)
    return text.format(**kwargs)

# ================= 图像处理模块 =================
def remove_background(img, bg_color=(255, 255, 255), tolerance=30):
    """
    移除图像背景：将与指定背景颜色接近的像素设为透明
    参数:
        img: PIL Image 对象
        bg_color: 背景色元组，默认为白色(255,255,255)
        tolerance: 颜色容差，像素差在此范围内则视为背景色
    返回:
        转换为RGBA模式并处理后的图像
    """
    img = img.convert("RGBA")
    datas = img.getdata()
    newData = []
    for item in datas:
        if (abs(item[0] - bg_color[0]) < tolerance and
            abs(item[1] - bg_color[1]) < tolerance and
            abs(item[2] - bg_color[2]) < tolerance):
            newData.append((item[0], item[1], item[2], 0))
        else:
            newData.append(item)
    img.putdata(newData)
    return img

def save_as_svg(img, output_path):
    """
    将图像保存为SVG文件
    实现方法：将图像转换为PNG并进行Base64编码，然后嵌入到SVG文件中
    参数:
        img: PIL Image 对象
        output_path: 输出SVG文件的路径
    """
    img = remove_background(img)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    data = base64.b64encode(buf.getvalue()).decode('utf-8')
    width, height = img.size
    svg_str = f'''<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">
  <image href="data:image/png;base64,{data}" width="{width}" height="{height}" />
</svg>
'''
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(svg_str)
        logger.debug(f"SVG saved to {output_path}")
    except Exception as e:
        logger.error(f"Error saving SVG: {e}")

# ================= 文件转换模块 =================
def convert_image_file(file_path, output_folder, output_format, ext_out):
    """
    转换图片文件，并根据选择的输出格式保存
    参数:
        file_path: 输入文件路径（Path对象）
        output_folder: 输出文件夹路径（Path对象）
        output_format: 输出格式字符串，例如 "JPEG"、"SVG"等
        ext_out: 输出文件的扩展名
    """
    base_name = file_path.stem
    out_path = output_folder / (base_name + ext_out)
    with Image.open(file_path) as img:
        if output_format in ["JPEG", "JPEG2000", "HEIC"] and img.mode != "RGB":
            img = img.convert("RGB")
        if output_format == "SVG":
            save_as_svg(img, out_path)
        elif output_format == "JPEG":
            img.save(out_path, output_format, quality=100)
        elif output_format == "PDF":
            dpi = img.info.get("dpi", (300, 300))[0]
            img.save(out_path, output_format, resolution=dpi)
        else:
            img.save(out_path, output_format)
    logger.info(f"Converted image {file_path} to {out_path}")

def convert_pdf_file(file_path, output_folder, output_format, ext_out):
    """
    转换PDF文件：将PDF的每一页转换为图片后保存
    参数:
        file_path: 输入PDF文件的路径（Path对象）
        output_folder: 输出文件夹路径（Path对象）
        output_format: 输出格式
        ext_out: 输出文件扩展名
    返回:
        转换成功的页面数
    """
    base_name = file_path.stem
    try:
        doc = fitz.open(str(file_path))
    except Exception as e:
        logger.error(f"Error opening PDF {file_path}: {e}")
        return 0
    count = 0
    zoom = 3.0
    mat = fitz.Matrix(zoom, zoom)
    for i, page in enumerate(doc):
        try:
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            page_img = Image.open(io.BytesIO(img_data))
            if output_format in ["JPEG", "JPEG2000"] and page_img.mode != "RGB":
                page_img = page_img.convert("RGB")
            out_name = f"{base_name}_page{i + 1}{ext_out}"
            out_path = output_folder / out_name
            if output_format == "SVG":
                save_as_svg(page_img, out_path)
            elif output_format == "JPEG":
                page_img.save(out_path, output_format, quality=100)
            elif output_format == "PDF":
                dpi = page_img.info.get("dpi", (300, 300))[0]
                page_img.save(out_path, output_format, resolution=dpi)
            else:
                page_img.save(out_path, output_format)
            logger.info(f"Converted PDF page {i+1} to {out_path}")
            count += 1
        except Exception as e:
            logger.error(f"Error converting PDF page {i+1} of {file_path}: {e}")
    return count

def convert_document_file(file_path, output_folder, output_format, ext_out):
    """
    处理文档类文件（Word、Excel、CSV）：目前只进行文件复制
    参数:
        file_path: 输入文件路径（Path对象）
        output_folder: 输出文件夹路径（Path对象）
        output_format: 目标格式（应为 "EXCEL"、"WORD" 或 "CSV"）
        ext_out: 输出文件扩展名
    """
    base_name = file_path.stem
    out_path = output_folder / (base_name + ext_out)
    try:
        shutil.copy(str(file_path), str(out_path))
        logger.info(f"Copied document {file_path} to {out_path}")
    except Exception as e:
        logger.error(f"Error copying document {file_path}: {e}")

def convert_file(file_path, output_folder, selected_format):
    """
    根据文件类型和选择的输出格式转换文件
    参数:
        file_path: 输入文件路径（Path对象）
        output_folder: 输出文件夹路径（Path对象）
        selected_format: 选择的输出格式（例如 "JPEG"、"PDF"、"EXCEL"等）
    返回:
        转换成功的文件数（对于PDF可能返回多个页面）
    """
    output_format, ext_out = FORMAT_MAPPING.get(selected_format, ("JPEG", ".jpg"))
    ext_in = file_path.suffix.lower()
    if ext_in == ".pdf":
        return convert_pdf_file(file_path, output_folder, output_format, ext_out)
    elif ext_in in [".doc", ".docx", ".xlsx", ".xls", ".csv"]:
        if output_format not in ["EXCEL", "WORD", "CSV"]:
            raise ValueError(f"Cannot convert document file {file_path} to {output_format} format.")
        convert_document_file(file_path, output_folder, output_format, ext_out)
        return 1
    else:
        if output_format in ["EXCEL", "WORD", "CSV"]:
            raise ValueError(f"Cannot convert image file {file_path} to {output_format} format.")
        convert_image_file(file_path, output_folder, output_format, ext_out)
        return 1

# ================= 单实例检测模块 =================
def check_single_instance(port=9999):
    """
    使用 socket 绑定方式检测是否已有程序实例运行
    参数:
        port: 检测所使用的端口号（默认为9999）
    返回:
        如果成功绑定返回socket对象，否则返回None
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
    except socket.error:
        return None
    s.listen(1)
    return s

# ================= GUI 主程序模块 =================
class ImageConverterApp:
    """
    主程序类：负责创建和管理GUI，包含文件预览、拖拽、选择输出文件夹以及转换操作
    """
    def __init__(self, master):
        """
        初始化主窗口和所有控件
        参数:
            master: Tkinter主窗口对象
        """
        self.master = master
        master.title(_( "window_title"))
        master.geometry("600x800")
        self.files = []            # 存储待转换文件列表（Path对象）
        self.output_folder = None  # 输出文件夹（Path对象）
        self.preview_images = []   # 存储预览图片的PhotoImage引用
        self.resize_after_id = None  # 防抖定时器ID
        self._suspend_refresh = False  # 控制预览更新的标志

        # ---------- 创建菜单栏 ----------
        self.create_menu()

        # ---------- 创建提示标签 ----------
        self.label = tk.Label(master, text=_("label_instruction"), font=("Arial", 12))
        self.label.pack(pady=10)

        # ---------- 创建拖拽区域 ----------
        self.drop_area = tk.Label(master, text=_("label_drop_area"),
                                  width=100, height=10, bg="lightgrey", relief="ridge")

        self.drop_area.pack(pady=10)
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.drop)

        # 创建一个Frame容器来装按钮
        button_frame = tk.Frame(master)
        button_frame.pack(pady=20)  # 只在Frame外部设置垂直间距

        # 使用 grid 布局在同一行放置按钮
        self.select_files_button = tk.Button(button_frame, text=_("btn_select_files"),
                                             command=self.select_files)
        self.select_files_button.grid(row=0, column=0, padx=5)

        self.select_output_button = tk.Button(button_frame, text=_("btn_select_output"),
                                              command=self.select_output_folder)
        self.select_output_button.grid(row=0, column=1, padx=5)

        self.convert_button = tk.Button(button_frame, text=_("btn_convert"),
                                        command=self.start_conversion_thread)
        self.convert_button.grid(row=0, column=2, padx=5)

        self.clear_all_button = tk.Button(button_frame, text=_("btn_clear_all"), command=self.clear_all_files)
        self.clear_all_button.grid(row=0, column=3, padx=5)

        # ---------- 创建输出格式选择区域 ----------
        self.format_frame = tk.Frame(master)
        self.format_frame.pack(pady=10)
        self.format_label = tk.Label(self.format_frame, text=_("label_output_format"), font=("Arial", 12))
        self.format_label.pack(side=tk.LEFT, padx=5)
        self.format_var = tk.StringVar()
        self.format_var.set(OUTPUT_FORMAT_LIST[0])
        self.format_combobox = ttk.Combobox(self.format_frame, textvariable=self.format_var,
                                            values=OUTPUT_FORMAT_LIST, state="readonly", width=12)
        self.format_combobox.pack(side=tk.LEFT)

        # ---------- 创建预览区域上方的文件统计标签 ----------
        # 修改统计标签文本格式，显示“所有文件为：X”
        self.preview_count_label = tk.Label(master, text=_("label_all_files", n=0), font=("Arial", 12))
        self.preview_count_label.pack(pady=5)
        self.preview_detail_label = tk.Label(master, text="---", font=("Arial", 12))
        self.preview_detail_label.pack(pady=5)

        # ---------- 创建带滚动条的预览区域 ----------
        preview_frame = tk.Frame(master)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.preview_canvas = tk.Canvas(preview_frame, bg="white")
        self.preview_canvas.pack(side="left", fill="both", expand=True)
        self.v_scrollbar = tk.Scrollbar(preview_frame, orient="vertical", command=self.preview_canvas.yview)
        self.v_scrollbar.pack(side="right", fill="y")
        self.h_scrollbar = tk.Scrollbar(master, orient="horizontal", command=self.preview_canvas.xview)
        self.h_scrollbar.pack(fill="x")
        self.preview_canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        self.preview_container = tk.Frame(self.preview_canvas, bg="white")
        self.preview_canvas.create_window((0, 0), window=self.preview_container, anchor="nw")
        self.preview_container.bind("<Configure>", self.on_frame_configure)

        # ---------- 绑定窗口大小变化事件（防抖机制） ----------
        master.bind("<Configure>", self.on_master_configure)

        # ---------- 创建转换进度条 ----------
        self.progress = ttk.Progressbar(master, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=5)
        self.progress.pack_forget()

    # ---------- UI 状态控制方法 ----------
    def disable_ui(self):
        """禁用文件选择、输出选择、转换等操作按钮"""
        self.select_files_button.config(state="disabled")
        self.select_output_button.config(state="disabled")
        self.convert_button.config(state="disabled")
        self.clear_all_button.config(state="disabled")

    def enable_ui(self):
        """恢复操作按钮的可用状态"""
        self.select_files_button.config(state="normal")
        self.select_output_button.config(state="normal")
        self.convert_button.config(state="normal")
        self.clear_all_button.config(state="normal")

    # ---------- 菜单栏创建与语言切换方法 ----------
    def create_menu(self):
        """创建菜单栏，包含“帮助”和“语言”菜单"""
        self.menubar = tk.Menu(self.master)
        self.master.config(menu=self.menubar)
        help_menu = tk.Menu(self.menubar, tearoff=0)
        help_menu.add_command(label=_("menu_supported_formats"), command=self.show_supported_formats)
        help_menu.add_separator()
        help_menu.add_command(label=_("menu_about"), command=self.show_about)
        self.menubar.add_cascade(label=_("menu_help"), menu=help_menu)
        lang_menu = tk.Menu(self.menubar, tearoff=0)
        lang_menu.add_command(label=translations['en']['lang_english'], command=lambda: self.set_language('en'))
        lang_menu.add_command(label=translations['es']['lang_spanish'], command=lambda: self.set_language('es'))
        lang_menu.add_command(label=translations['ru']['lang_russian'], command=lambda: self.set_language('ru'))
        lang_menu.add_command(label=translations['zh']['lang_chinese'], command=lambda: self.set_language('zh'))
        self.menubar.add_cascade(label=_("menu_language"), menu=lang_menu)

    def set_language(self, lang):
        """切换界面语言并更新所有控件文本，同时保存配置"""
        global current_lang
        current_lang = lang
        self.master.title(_("window_title"))
        self.label.config(text=_("label_instruction"))
        self.drop_area.config(text=_("label_drop_area"))
        self.select_files_button.config(text=_("btn_select_files"))
        self.select_output_button.config(text=_("btn_select_output"))
        self.format_label.config(text=_("label_output_format"))
        self.convert_button.config(text=_("btn_convert"))
        self.clear_all_button.config(text=_("btn_clear_all"))
        self.create_menu()
        self.update_preview()
        config = load_config()
        config["language"] = lang
        save_config(config)
        logger.info(f"Language set to {lang}")

    def show_supported_formats(self):
        """显示支持的输入和输出格式信息"""
        input_formats = " ".join(SUPPORTED_EXTS)
        output_formats = " ".join(OUTPUT_FORMAT_LIST)
        fmt_info = _("msg_supported_formats", inputs=input_formats, outputs=output_formats)
        messagebox.showinfo(_("menu_supported_formats"), fmt_info)

    def show_about(self):
        """显示“关于”信息"""
        messagebox.showinfo(_("menu_about"), _("msg_about"))

    # ---------- 预览区域更新与防抖方法 ----------
    def on_frame_configure(self, event):
        """更新Canvas的滚动区域"""
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

    def on_master_configure(self, event):
        """主窗口尺寸变化时采用防抖机制更新预览区域"""
        if not self._suspend_refresh:
            if self.resize_after_id:
                self.master.after_cancel(self.resize_after_id)
            self.resize_after_id = self.master.after(500, self.update_preview)

    def update_preview(self):
        """
        更新预览区域：清空原有内容，重新生成缩略图或文件类型标记，并更新文件统计信息
        """
        for widget in self.preview_container.winfo_children():
            widget.destroy()
        self.preview_images = []
        available_width = self.preview_canvas.winfo_width()
        if available_width < 100:
            available_width = 550
        thumb_width = 100
        padding = 10
        col_count = max(1, available_width // (thumb_width + padding))
        # 更新文件统计标签，显示“所有文件为: X”
        self.preview_count_label.config(text=_("label_all_files", n=len(self.files)))
        type_counts = {}
        for file in self.files:
            ext = file.suffix.lower()
            type_counts[ext] = type_counts.get(ext, 0) + 1
        detail_str = " / ".join([f"{ext.upper()} ({count})" for ext, count in type_counts.items()]) or "---"
        self.preview_detail_label.config(text=detail_str)
        max_name_length = 15
        for idx, file in enumerate(self.files):
            row = idx // col_count
            col = idx % col_count
            frame = tk.Frame(self.preview_container, bg="white", bd=1, relief="solid")
            frame.grid(row=row, column=col, padx=5, pady=5, sticky="n")
            ext = file.suffix.lower()
            if ext not in [".pdf", ".doc", ".docx", ".xlsx", ".xls", ".csv"]:
                try:
                    with Image.open(file) as img:
                        img.thumbnail((thumb_width, 100))
                        photo = ImageTk.PhotoImage(img)
                        self.preview_images.append(photo)
                        img_label = tk.Label(frame, image=photo, bg="white")
                        img_label.pack(padx=5, pady=5)
                except Exception as e:
                    logger.error(f"Error previewing {file}: {e}")
                    tk.Label(frame, text="Error", bg="white").pack(padx=5, pady=5)
            else:
                if ext == ".pdf":
                    file_type = "PDF"
                elif ext in [".doc", ".docx"]:
                    file_type = "DOC"
                elif ext in [".xlsx", ".xls"]:
                    file_type = "EXCEL"
                elif ext == ".csv":
                    file_type = "CSV"
                else:
                    file_type = ""
                tk.Label(frame, text=file_type, bg="white", font=("Arial", 16)).pack(padx=5, pady=20)
            name = file.name
            base, ext_str = os.path.splitext(name)
            if len(name) > max_name_length:
                truncated_length = max(0, max_name_length - len(ext_str) - 3)
                display_name = base[:truncated_length] + "..." + ext_str
            else:
                display_name = name
            cancel_btn = tk.Button(frame, text="✕", command=lambda f=file: self.remove_file(f),
                                   font=("Arial", 10), fg="white", bg="Green", bd=0, padx=4, pady=2)
            cancel_btn.pack(side="top", anchor="ne", padx=2, pady=2)
            name_label = tk.Label(frame, text=display_name, bg="white", font=("Arial", 10))
            name_label.pack(padx=2, pady=2)

    def remove_file(self, file):
        """
        从文件列表中删除指定文件，并更新预览区域
        参数:
            file: 要删除的文件（Path对象）
        """
        if file in self.files:
            if self.resize_after_id:
                try:
                    self.master.after_cancel(self.resize_after_id)
                except Exception:
                    pass
                self.resize_after_id = None
            self._suspend_refresh = True
            self.files.remove(file)
            self.update_preview()
            self._suspend_refresh = False

    def clear_all_files(self):
        """
        清除预览区域中所有已添加的文件，并更新统计信息
        """
        self.files.clear()
        self.update_preview()
        messagebox.showinfo(_("btn_clear_all"), _("msg_file_added_success", n=0))

    # ---------- 文件选择与拖拽方法 ----------
    def drop(self, event):
        """
        处理拖拽事件，将有效文件添加到文件列表中
        参数:
            event: 拖拽事件对象，包含拖入文件的路径列表
        """
        dropped_files = self.master.tk.splitlist(event.data)
        valid_files = []
        for file in dropped_files:
            path = Path(file)
            if path.is_file() and path.suffix.lower() in SUPPORTED_EXTS:
                valid_files.append(path)
        if valid_files:
            self.files.extend(valid_files)
            messagebox.showinfo(_("btn_select_files"), _("msg_file_added_success", n=len(valid_files)))
            self.update_preview()
        else:
            messagebox.showwarning(_("btn_select_files"), _("msg_invalid_file"))

    def select_files(self):
        """
        弹出文件选择对话框，允许用户选择文件并添加到文件列表中
        """
        file_types = [
            ("All Supported Files", tuple("*" + ext for ext in SUPPORTED_EXTS)),
            ("PNG", "*.png"),
            ("JPG", "*.jpg"),
            ("JPEG", "*.jpeg"),
            ("JFIF", "*.jfif"),
            ("BMP", "*.bmp"),
            ("GIF", "*.gif"),
            ("TIFF", "*.tiff"),
            ("WEBP", "*.webp"),
            ("ICO", "*.ico"),
            ("PPM", "*.ppm"),
            ("TGA", "*.tga"),
            ("JPEG2000", "*.jp2"),
            ("PDF", "*.pdf"),
            ("SVG", "*.svg"),
            ("HEIC", "*.heic"),
            ("Excel", ("*.xlsx", "*.xls")),
            ("Word", ("*.doc", "*.docx")),
            ("CSV", "*.csv")
        ]
        selected_files = filedialog.askopenfilenames(title=_("btn_select_files"), filetypes=file_types)
        if selected_files:
            paths = [Path(f) for f in selected_files]
            self.files.extend(paths)
            messagebox.showinfo(_("btn_select_files"), _("msg_select_files", n=len(paths)))
            self.update_preview()

    def select_output_folder(self):
        """
        弹出文件夹选择对话框，设置转换后的输出文件夹
        """
        folder = filedialog.askdirectory(title=_("btn_select_output"))
        if folder:
            self.output_folder = Path(folder)
            messagebox.showinfo(_("btn_select_output"), _("msg_select_output_folder", folder=str(self.output_folder)))
            logger.info(f"Output folder set to {self.output_folder}")

    # ---------- 转换任务（异步处理）方法 ----------
    def start_conversion_thread(self):
        """
        创建并启动后台线程执行转换任务，防止阻塞GUI主线程
        """
        if not self.files:
            messagebox.showwarning(_("btn_convert"), _("msg_warning_no_file"))
            return
        if not self.output_folder:
            messagebox.showwarning(_("btn_convert"), _("msg_warning_no_output"))
            return
        self.disable_ui()
        self.progress.pack(fill="x", padx=10, pady=5)
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        threading.Thread(target=self.convert_files, daemon=True).start()

    def convert_files(self):
        """
        遍历所有待转换文件，根据文件类型调用对应转换方法，
        在后台线程中运行，同时更新进度条，转换完成后恢复UI状态
        """
        count = 0
        errors = []
        selected_format = self.format_var.get().upper()
        for idx, file in enumerate(self.files):
            try:
                result = convert_file(file, self.output_folder, selected_format)
                count += result
            except Exception as e:
                err_msg = f"Error converting {file}: {e}"
                logger.error(err_msg)
                errors.append(err_msg)
            self.master.after(0, self.progress.step, 1)
        self.master.after(0, lambda: self.conversion_complete(count, errors))

    def conversion_complete(self, count, errors):
        """
        转换任务完成后调用，显示结果提示信息，并重置UI状态
        参数:
            count: 转换成功的文件（或页面）数
            errors: 转换过程中产生的错误信息列表
        """
        messagebox.showinfo(_("btn_convert"), _("msg_convert_complete", n=count))
        if errors:
            messagebox.showerror("Conversion Errors", "\n".join(errors))
        self.files.clear()
        self.update_preview()
        self.progress.pack_forget()
        self.enable_ui()

# ================= 程序入口 =================
if __name__ == "__main__":
    lock_socket = check_single_instance()
    if lock_socket is None:
        temp = tk.Tk()
        temp.withdraw()
        messagebox.showwarning("Warning", _("msg_conflict"))
        sys.exit(0)

    root = TkinterDnD.Tk()

    # 可选：设置窗口图标
    try:
        ico_img = ImageTk.PhotoImage(file="2.ico")  # 请确保图标路径正确
        root.iconphoto(True, ico_img)  # 设置窗口图标
    except Exception as e:
        logger.error("Ico not set: " + str(e))  # 错误日志

    app = ImageConverterApp(root)
    root.mainloop()
