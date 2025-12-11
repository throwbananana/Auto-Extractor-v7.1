#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
书籍翻译工具 GUI
支持PDF、TXT、EPUB格式的书籍翻译
可接入Gemini API、OpenAI API等多种翻译API
云端配额耗尽时可自动切换到本地LM Studio模型

依赖库:
pip install PyPDF2 ebooklib beautifulsoup4 google-generativeai openai requests
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import time
from pathlib import Path
import re
import json
import shutil
from copy import deepcopy
import hashlib
import json

# 文件读取相关
try:
    import PyPDF2
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    import ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    EPUB_SUPPORT = True
except ImportError:
    EPUB_SUPPORT = False

# API相关
try:
    import google.generativeai as genai
    GEMINI_SUPPORT = True
except ImportError:
    GEMINI_SUPPORT = False

try:
    import openai
    OPENAI_SUPPORT = True
except ImportError:
    OPENAI_SUPPORT = False

try:
    import requests
    REQUESTS_SUPPORT = True
except ImportError:
    REQUESTS_SUPPORT = False

from file_processor import FileProcessor

DEFAULT_TARGET_LANGUAGE = "中文"
DEFAULT_LM_STUDIO_CONFIG = {
    'api_key': 'lm-studio',
    'model': 'qwen2.5-7b-instruct-1m',
    'base_url': 'http://127.0.0.1:1234/v1'
}

DEFAULT_API_CONFIGS = {
    'gemini': {'api_key': '', 'model': 'gemini-2.5-flash'},
    'openai': {'api_key': '', 'model': 'gpt-3.5-turbo', 'base_url': ''},
    'custom': {'api_key': '', 'model': '', 'base_url': ''},
    'lm_studio': deepcopy(DEFAULT_LM_STUDIO_CONFIG)
}

class BookTranslatorGUI:
    """书籍翻译工具主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("书籍翻译工具 v1.4 - 支持LM Studio与自定义目标语言")
        self.root.geometry("900x700")

        # 初始化辅助模块
        self.file_processor = FileProcessor()
        self.progress_cache_path = Path(__file__).parent / 'translation_cache.json'

        # 程序退出时自动保存配置
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 翻译状态
        self.is_translating = False
        self.current_text = ""
        self.translated_text = ""
        self.translation_thread = None
        self.source_segments = []
        self.translated_segments = []
        self.failed_segments = []
        self.selected_failed_index = None
        # 是否已启用本地LM Studio备用方案
        self.lm_studio_fallback_active = False
        # 进度缓存/恢复控制
        self.text_signature = None
        self.resume_from_index = 0
        self.max_consecutive_failures = 3
        self.consecutive_failures = 0
        self.paused_due_to_failures = False

        # 大文件处理
        self.show_full_text = False
        self.preview_limit = 10000  # 预览显示前10000字符

        # API配置
        self.api_configs = deepcopy(DEFAULT_API_CONFIGS)
        self.target_language_var = tk.StringVar(value=DEFAULT_TARGET_LANGUAGE)

        self.setup_ui()
        self.load_config()
        self.try_resume_cached_progress()

    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # 1. 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="选择文件:").grid(row=0, column=0, sticky=tk.W)
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, state='readonly').grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=5
        )
        ttk.Button(file_frame, text="浏览...", command=self.browse_file).grid(
            row=0, column=2, padx=5
        )

        # 支持的格式提示
        formats = []
        if PDF_SUPPORT:
            formats.append("PDF")
        if EPUB_SUPPORT:
            formats.append("EPUB")
        formats.append("TXT")

        support_label = f"支持格式: {', '.join(formats)}"
        ttk.Label(file_frame, text=support_label, foreground="gray").grid(
            row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0)
        )

        # 文件大小和预览控制
        preview_frame = ttk.Frame(file_frame)
        preview_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))

        self.file_info_var = tk.StringVar(value="")
        ttk.Label(preview_frame, textvariable=self.file_info_var, foreground="blue").pack(
            side=tk.LEFT
        )

        self.toggle_preview_btn = ttk.Button(
            preview_frame,
            text="显示完整原文",
            command=self.toggle_full_text_display,
            state='disabled'
        )
        self.toggle_preview_btn.pack(side=tk.RIGHT, padx=5)

        # 2. API配置区域
        api_frame = ttk.LabelFrame(main_frame, text="API配置", padding="10")
        api_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        api_frame.columnconfigure(1, weight=1)

        # API类型选择
        ttk.Label(api_frame, text="API类型:").grid(row=0, column=0, sticky=tk.W)
        self.api_type_var = tk.StringVar(value="gemini")
        api_types = []
        if GEMINI_SUPPORT:
            api_types.append(("Gemini API", "gemini"))
        if OPENAI_SUPPORT:
            api_types.append(("OpenAI API", "openai"))
            api_types.append(("本地 LM Studio", "lm_studio"))
        if REQUESTS_SUPPORT:
            api_types.append(("自定义API", "custom"))

        api_combo = ttk.Combobox(
            api_frame,
            textvariable=self.api_type_var,
            values=[name for name, _ in api_types],
            state='readonly',
            width=20
        )
        api_combo.grid(row=0, column=1, sticky=tk.W, padx=5)
        api_combo.bind('<<ComboboxSelected>>', self.on_api_type_change)

        ttk.Button(api_frame, text="配置API", command=self.open_api_config).grid(
            row=0, column=2, padx=5
        )

        # API状态
        self.api_status_var = tk.StringVar(value="未配置")
        self.api_status_label = ttk.Label(api_frame, textvariable=self.api_status_var, foreground="orange")
        self.api_status_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

        ttk.Label(api_frame, text="目标语言:").grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        lang_options = ["中文", "英文", "English", "日语", "韩语", "德语", "法语"]
        lang_combo = ttk.Combobox(
            api_frame,
            textvariable=self.target_language_var,
            values=lang_options,
            state='normal',
            width=20
        )
        lang_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=(5, 0))
        ttk.Label(
            api_frame,
            text="可输入任意目标语言，例如：中文、英文、日语、法语",
            foreground="gray",
            font=('', 8)
        ).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(2, 0))

        ttk.Label(
            api_frame,
            text="API配额用尽时将自动切换到本地LM Studio (默认 http://127.0.0.1:1234/v1)",
            foreground="gray",
            font=('', 8)
        ).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

        # 3. 翻译内容显示区域
        content_frame = ttk.LabelFrame(main_frame, text="翻译内容", padding="10")
        content_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        # 创建Notebook用于显示原文和译文
        self.notebook = ttk.Notebook(content_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 原文标签页
        original_frame = ttk.Frame(self.notebook)
        self.notebook.add(original_frame, text="原文")
        original_frame.columnconfigure(0, weight=1)
        original_frame.rowconfigure(0, weight=1)

        self.original_text = scrolledtext.ScrolledText(
            original_frame, wrap=tk.WORD, height=15
        )
        self.original_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 译文标签页
        translated_frame = ttk.Frame(self.notebook)
        self.notebook.add(translated_frame, text="译文")
        translated_frame.columnconfigure(0, weight=1)
        translated_frame.rowconfigure(0, weight=1)

        self.translated_text_widget = scrolledtext.ScrolledText(
            translated_frame, wrap=tk.WORD, height=15
        )
        self.translated_text_widget.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 失败段落标签页
        failed_frame = ttk.Frame(self.notebook)
        self.notebook.add(failed_frame, text="失败段落")
        failed_frame.columnconfigure(1, weight=1)
        failed_frame.rowconfigure(1, weight=1)
        failed_frame.rowconfigure(3, weight=1)

        ttk.Label(
            failed_frame,
            text="检测到的失败/未完成段落，可重试或手动翻译后替换。",
            foreground="gray"
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))

        self.failed_listbox = tk.Listbox(failed_frame, height=8)
        self.failed_listbox.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.W), padx=(0, 10))
        self.failed_listbox.bind('<<ListboxSelect>>', self.on_failed_select)

        detail_frame = ttk.Frame(failed_frame)
        detail_frame.grid(row=1, column=1, sticky=(tk.N, tk.S, tk.E, tk.W))
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(1, weight=1)
        detail_frame.rowconfigure(3, weight=1)

        ttk.Label(detail_frame, text="原文（只读）").grid(row=0, column=0, sticky=tk.W)
        self.failed_source_text = scrolledtext.ScrolledText(
            detail_frame, wrap=tk.WORD, height=6, state='disabled'
        )
        self.failed_source_text.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(0, 5))

        ttk.Label(detail_frame, text="手动翻译").grid(row=2, column=0, sticky=tk.W)
        self.manual_translation_text = scrolledtext.ScrolledText(
            detail_frame, wrap=tk.WORD, height=6
        )
        self.manual_translation_text.grid(row=3, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        button_frame = ttk.Frame(detail_frame)
        button_frame.grid(row=4, column=0, sticky=tk.E, pady=(5, 0))
        ttk.Button(button_frame, text="重试翻译", command=self.retry_failed_segment).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="保存手动翻译", command=self.save_manual_translation).grid(row=0, column=1, padx=5)

        self.failed_status_var = tk.StringVar(value="暂无失败段落")
        ttk.Label(failed_frame, textvariable=self.failed_status_var).grid(
            row=2, column=0, columnspan=3, sticky=tk.W, pady=(5, 0)
        )

        # 4. 进度和控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(0, weight=1)

        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            control_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        # 进度文本
        self.progress_text_var = tk.StringVar(value="就绪")
        ttk.Label(control_frame, textvariable=self.progress_text_var).grid(
            row=1, column=0, sticky=tk.W
        )

        # 5. 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))

        self.translate_btn = ttk.Button(
            button_frame, text="开始翻译", command=self.start_translation
        )
        self.translate_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = ttk.Button(
            button_frame, text="停止翻译", command=self.stop_translation, state='disabled'
        )
        self.stop_btn.grid(row=0, column=1, padx=5)

        ttk.Button(
            button_frame, text="导出译文", command=self.export_translation
        ).grid(row=0, column=2, padx=5)

        ttk.Button(
            button_frame, text="清空", command=self.clear_all
        ).grid(row=0, column=3, padx=5)

    def browse_file(self):
        """浏览并选择文件"""
        filetypes = [("所有支持的文件", "*.txt *.pdf *.epub")]
        if PDF_SUPPORT:
            filetypes.append(("PDF文件", "*.pdf"))
        if EPUB_SUPPORT:
            filetypes.append(("EPUB文件", "*.epub"))
        filetypes.append(("文本文件", "*.txt"))

        filename = filedialog.askopenfilename(
            title="选择要翻译的书籍",
            filetypes=filetypes
        )

        if filename:
            self.file_path_var.set(filename)
            self.load_file_content(filename)
            # 加载新文件时清理旧缓存
            self.clear_progress_cache()

    def load_file_content(self, filepath):
        """加载文件内容"""
        try:
            content = ""
            
            # 使用 FileProcessor 读取文件
            def update_progress(msg):
                self.progress_text_var.set(msg)
                self.root.update()

            content = self.file_processor.read_file(filepath, progress_callback=update_progress)

            if not content:
                raise ValueError("文件内容为空")

            self.current_text = content
            self.text_signature = self.compute_text_signature(self.current_text)
            self.source_segments = []
            self.translated_segments = []
            self.failed_segments = []
            self.resume_from_index = 0
            self.original_text.delete('1.0', tk.END)
            
            # 统计信息
            char_count = len(content)
            word_count = len(content.split())

            # 判断是否为大文件
            is_large_file = char_count > self.preview_limit

            # 更新显示
            self.update_text_display()

            # 更新文件信息
            if is_large_file:
                self.file_info_var.set(
                    f"⚠️ 大文件 ({char_count:,} 字符) - 仅显示前 {self.preview_limit:,} 字符"
                )
                self.toggle_preview_btn.config(state='normal')
            else:
                self.file_info_var.set(f"✓ 已加载完整文件 ({char_count:,} 字符)")
                self.toggle_preview_btn.config(state='disabled')

            self.progress_text_var.set(f"已加载文件 | 字符数: {char_count:,} | 词数: {word_count:,}")

            # 提示信息
            msg = f"文件加载成功!\n\n字符数: {char_count:,}\n词数: {word_count:,}"
            if is_large_file:
                msg += f"\n\n⚠️ 这是一个大文件！\n为了性能，预览窗口仅显示前 {self.preview_limit:,} 字符。\n\n翻译时会使用完整文本。"
            messagebox.showinfo("成功", msg)

        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败:\n{str(e)}")

    def on_api_type_change(self, event=None):
        """API类型改变时更新状态"""
        self.update_api_status()

    def update_api_status(self):
        """更新API配置状态"""
        api_type = self.get_current_api_type()
        config = self.api_configs.get(api_type, {})

        if config.get('api_key'):
            self.api_status_var.set(f"已配置 API Key: {config['api_key'][:10]}...")
            self.api_status_label.config(foreground="green")
        else:
            self.api_status_var.set("未配置 API Key")
            self.api_status_label.config(foreground="orange")

    def get_current_api_type(self):
        """获取当前选择的API类型"""
        api_name = self.api_type_var.get()
        api_map = {
            "Gemini API": "gemini",
            "OpenAI API": "openai",
            "本地 LM Studio": "lm_studio",
            "自定义API": "custom"
        }
        return api_map.get(api_name, "gemini")

    def get_target_language(self):
        """获取用户设置的目标语言"""
        target = (self.target_language_var.get() or "").strip()
        return target if target else DEFAULT_TARGET_LANGUAGE

    def is_target_language_chinese(self, target_language=None):
        """判断目标语言是否为中文"""
        target = (target_language or self.get_target_language() or "").lower()
        return any(key in target for key in ["中文", "汉语", "chinese", "zh"])

    def is_target_language_english(self, target_language=None):
        """判断目标语言是否为英文"""
        target = (target_language or self.get_target_language() or "").lower()
        return any(key in target for key in ["英文", "英语", "english", "en"])

    def compute_text_signature(self, text):
        """计算文本签名用于断点恢复"""
        return hashlib.md5(text.encode('utf-8')).hexdigest() if text else None

    def save_progress_cache(self):
        """保存当前翻译进度到磁盘"""
        try:
            if not self.current_text or not self.source_segments:
                return

            data = {
                'file_path': self.file_path_var.get(),
                'signature': self.text_signature,
                'source_segments': self.source_segments,
                'translated_segments': self.translated_segments,
                'failed_segments': self.failed_segments,
                'lm_studio_fallback_active': self.lm_studio_fallback_active,
                'target_language': self.get_target_language(),
                'resume_from_index': len(self.translated_segments)
            }
            with open(self.progress_cache_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception as e:
            print(f"保存进度缓存失败: {e}")

    def clear_progress_cache(self):
        """清除翻译进度缓存"""
        try:
            if self.progress_cache_path.exists():
                self.progress_cache_path.unlink()
        except Exception as e:
            print(f"清除进度缓存失败: {e}")

    def try_resume_cached_progress(self):
        """启动时检查并询问是否恢复未完成进度"""
        if not self.progress_cache_path.exists():
            return

        try:
            with open(self.progress_cache_path, 'r', encoding='utf-8') as f:
                cache = json.load(f)
        except Exception as e:
            print(f"读取进度缓存失败: {e}")
            return

        file_path = cache.get('file_path')
        signature = cache.get('signature')
        if not file_path or not Path(file_path).exists():
            print("进度缓存对应的文件不存在，已忽略")
            self.clear_progress_cache()
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as rf:
                content = rf.read()
        except Exception as e:
            print(f"读取缓存文件失败: {e}")
            self.clear_progress_cache()
            return

        current_sig = self.compute_text_signature(content)
        if signature != current_sig:
            print("文件内容已变化，无法恢复进度，已清除缓存")
            self.clear_progress_cache()
            return

        # 恢复状态
        self.file_path_var.set(file_path)
        self.current_text = content
        self.text_signature = signature
        self.source_segments = cache.get('source_segments', [])
        self.translated_segments = cache.get('translated_segments', [])
        self.failed_segments = cache.get('failed_segments', [])
        self.lm_studio_fallback_active = cache.get('lm_studio_fallback_active', False)
        self.resume_from_index = cache.get('resume_from_index', len(self.translated_segments))
        cached_target = cache.get('target_language')
        if cached_target:
            self.target_language_var.set(cached_target)

        # 更新界面显示
        self.update_text_display()
        self.translated_text = "\n\n".join(self.translated_segments)
        self.update_translated_text(self.translated_text)

        total_segments = len(self.source_segments) or 1
        progress = (len(self.translated_segments) / total_segments) * 100
        self.progress_var.set(progress)
        self.progress_text_var.set(f"检测到未完成的翻译进度（{len(self.translated_segments)}/{total_segments} 段）")

        continue_resume = messagebox.askyesno(
            "继续未完成的翻译",
            f"检测到未完成的翻译任务：\n文件: {Path(file_path).name}\n进度: {len(self.translated_segments)}/{total_segments}\n\n是否继续？"
        )
        if not continue_resume:
            # 放弃恢复，清理缓存并重置状态
            self.clear_progress_cache()
            self.translated_segments = []
            self.source_segments = []
            self.failed_segments = []
            self.resume_from_index = 0
            self.translated_text = ""
            self.translated_text_widget.delete('1.0', tk.END)
            self.progress_var.set(0)
            self.progress_text_var.set("就绪")

    def open_api_config(self):
        """打开API配置对话框"""
        api_type = self.get_current_api_type()
        config = self.api_configs[api_type]

        config_window = tk.Toplevel(self.root)
        config_window.title(f"{self.api_type_var.get()} 配置")
        config_window.geometry("500x300")
        config_window.transient(self.root)
        config_window.grab_set()

        frame = ttk.Frame(config_window, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        frame.columnconfigure(1, weight=1)

        # API Key
        ttk.Label(frame, text="API Key:").grid(row=0, column=0, sticky=tk.W, pady=5)
        api_key_var = tk.StringVar(value=config.get('api_key', ''))
        ttk.Entry(frame, textvariable=api_key_var, width=50).grid(
            row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5
        )

        # Model
        ttk.Label(frame, text="模型:").grid(row=1, column=0, sticky=tk.W, pady=5)
        model_var = tk.StringVar(value=config.get('model', ''))
        ttk.Entry(frame, textvariable=model_var, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5
        )

        # Base URL (for OpenAI compatible APIs)
        if api_type in ['openai', 'custom', 'lm_studio']:
            ttk.Label(frame, text="Base URL:").grid(row=2, column=0, sticky=tk.W, pady=5)
            base_url_var = tk.StringVar(value=config.get('base_url', ''))
            ttk.Entry(frame, textvariable=base_url_var, width=50).grid(
                row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5
            )
            ttk.Label(
                frame,
                text="(可选，用于自定义OpenAI兼容API)",
                foreground="gray"
            ).grid(row=3, column=1, sticky=tk.W, pady=5)
        else:
            base_url_var = tk.StringVar(value='')

        # 说明文本
        help_text = {
            'gemini': "请在 Google AI Studio 获取 API Key\n模型示例: gemini-2.5-flash, gemini-2.5-pro",
            'openai': "请在 OpenAI 获取 API Key\n模型示例: gpt-3.5-turbo, gpt-4",
            'custom': (
                "输入兼容OpenAI API格式的自定义服务\n"
                "Base URL示例: https://api.example.com/v1\n"
                "本地LM Studio示例: http://127.0.0.1:1234/v1 (模型如 qwen2.5-7b-instruct-1m)"
            ),
            'lm_studio': (
                "连接本地 LM Studio 提供的 OpenAI 兼容接口\n"
                "默认地址: http://127.0.0.1:1234/v1\n"
                "请确保 LM Studio Server 已启动并加载目标模型"
            )
        }

        help_label = ttk.Label(
            frame,
            text=help_text.get(api_type, ''),
            foreground="gray",
            justify=tk.LEFT
        )
        help_label.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=10)

        # 测试连接按钮
        def test_connection():
            """测试API连接"""
            test_api_key = api_key_var.get().strip()
            test_model = model_var.get().strip()

            if not test_api_key:
                messagebox.showwarning("警告", "请先输入API Key")
                return

            if not test_model:
                messagebox.showwarning("警告", "请先输入模型名称")
                return

            # 显示测试中提示
            test_btn.config(state='disabled', text="测试中...")
            config_window.update()

            try:
                if api_type == 'gemini':
                    import google.generativeai as genai
                    genai.configure(api_key=test_api_key)
                    model_obj = genai.GenerativeModel(test_model)
                    response = model_obj.generate_content("测试连接：请回复'OK'")
                    result = response.text
                    messagebox.showinfo("成功", f"✓ API连接测试成功！\n\n模型: {test_model}\n响应: {result[:100]}")

                elif api_type == 'openai':
                    import openai
                    client_kwargs = {'api_key': test_api_key}
                    if base_url_var.get().strip():
                        client_kwargs['base_url'] = base_url_var.get().strip()

                    client = openai.OpenAI(**client_kwargs)
                    response = client.chat.completions.create(
                        model=test_model,
                        messages=[{"role": "user", "content": "测试连接：请回复'OK'"}]
                    )
                    result = response.choices[0].message.content
                    messagebox.showinfo("成功", f"✓ API连接测试成功！\n\n模型: {test_model}\n响应: {result[:100]}")

                elif api_type == 'custom':
                    import requests
                    test_url = base_url_var.get().strip()
                    if not test_url:
                        messagebox.showwarning("警告", "请先输入Base URL")
                        return

                    url = f"{test_url.rstrip('/')}/chat/completions"
                    headers = {
                        'Content-Type': 'application/json',
                        'Authorization': f"Bearer {test_api_key}"
                    }
                    data = {
                        'model': test_model,
                        'messages': [{"role": "user", "content": "测试连接：请回复'OK'"}]
                    }
                    response = requests.post(url, headers=headers, json=data, timeout=30)
                    response.raise_for_status()
                    result = response.json()['choices'][0]['message']['content']
                    messagebox.showinfo("成功", f"✓ API连接测试成功！\n\n模型: {test_model}\n响应: {result[:100]}")

                elif api_type == 'lm_studio':
                    import openai
                    client = openai.OpenAI(
                        api_key=test_api_key,
                        base_url=base_url_var.get().strip() or DEFAULT_LM_STUDIO_CONFIG['base_url']
                    )
                    response = client.chat.completions.create(
                        model=test_model,
                        messages=[{"role": "user", "content": "测试连接：请回复'OK'"}],
                        temperature=0
                    )
                    result = response.choices[0].message.content
                    messagebox.showinfo("成功", f"✓ API连接测试成功！\n\n模型: {test_model}\n响应: {result[:100]}")

            except Exception as e:
                messagebox.showerror("测试失败", f"✗ API连接测试失败\n\n错误: {str(e)}\n\n请检查:\n1. API Key是否正确\n2. 模型名称是否正确\n3. 网络连接是否正常\n4. API服务是否可用")

            finally:
                test_btn.config(state='normal', text="测试连接")

        # 保存按钮
        def save_config():
            new_api_key = api_key_var.get().strip()
            new_model = model_var.get().strip()

            # 验证输入
            if not new_api_key:
                messagebox.showwarning("警告", "API Key不能为空")
                return

            if not new_model:
                messagebox.showwarning("警告", "模型名称不能为空")
                return

            # 保存配置
            self.api_configs[api_type]['api_key'] = new_api_key
            self.api_configs[api_type]['model'] = new_model
            if api_type in ['openai', 'custom', 'lm_studio']:
                self.api_configs[api_type]['base_url'] = base_url_var.get().strip()

            # 自动保存到文件
            self.save_config(show_message=True)
            self.update_api_status()
            config_window.destroy()
            messagebox.showinfo("成功", "✓ API配置已保存\n✓ 已自动创建备份\n\n配置将在下次启动时自动加载")

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)

        test_btn = ttk.Button(button_frame, text="测试连接", command=test_connection)
        test_btn.grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="保存", command=save_config).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="取消", command=config_window.destroy).grid(row=0, column=2, padx=5)

    def merge_api_configs(self, incoming_configs):
        """将磁盘配置与默认值合并，确保新字段有默认值"""
        incoming_configs = incoming_configs or {}

        for name, defaults in DEFAULT_API_CONFIGS.items():
            merged = deepcopy(defaults)
            merged.update(incoming_configs.get(name, {}))
            self.api_configs[name] = merged

        # 保留未知的扩展配置，避免意外丢失
        for extra_key, extra_val in incoming_configs.items():
            if extra_key not in self.api_configs:
                self.api_configs[extra_key] = extra_val

    def save_config(self, show_message=False):
        """保存配置到文件（自动保存）"""
        config_file = Path(__file__).parent / 'translator_config.json'
        try:
            # 保存主配置
            with open(config_file, 'w', encoding='utf-8') as f:
                data = {
                    'api_configs': self.api_configs,
                    'target_language': self.get_target_language()
                }
                json.dump(data, f, indent=2, ensure_ascii=False)

            # 自动创建备份（保留最近3个）
            self.backup_config()

            if show_message:
                print(f"✓ 配置已自动保存: {config_file}")
        except Exception as e:
            error_msg = f"保存配置失败: {e}"
            print(error_msg)
            if show_message:
                messagebox.showerror("错误", error_msg)

    def backup_config(self):
        """自动备份配置（保留最近3个备份）"""
        try:
            config_file = Path(__file__).parent / 'translator_config.json'
            if not config_file.exists():
                return

            backup_dir = Path(__file__).parent / 'config_backups'
            backup_dir.mkdir(exist_ok=True)

            # 生成备份文件名（时间戳）
            from datetime import datetime
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = backup_dir / f'config_backup_{timestamp}.json'

            # 复制当前配置到备份
            import shutil
            shutil.copy2(config_file, backup_file)

            # 只保留最近3个备份
            backups = sorted(backup_dir.glob('config_backup_*.json'), reverse=True)
            for old_backup in backups[3:]:
                old_backup.unlink()

        except Exception as e:
            print(f"备份配置失败: {e}")

    def load_config(self):
        """从文件加载配置（自动加载）"""
        config_file = Path(__file__).parent / 'translator_config.json'
        try:
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    if isinstance(loaded_config, dict) and 'api_configs' in loaded_config:
                        api_config_section = loaded_config.get('api_configs', {})
                        target_language = loaded_config.get('target_language', DEFAULT_TARGET_LANGUAGE)
                    else:
                        api_config_section = loaded_config if isinstance(loaded_config, dict) else {}
                        target_language = (
                            loaded_config.get('target_language', DEFAULT_TARGET_LANGUAGE)
                            if isinstance(loaded_config, dict) else DEFAULT_TARGET_LANGUAGE
                        )
                    self.merge_api_configs(api_config_section)
                    self.target_language_var.set(target_language or DEFAULT_TARGET_LANGUAGE)
                self.update_api_status()
                print(f"✓ 配置已加载: {len([k for k, v in self.api_configs.items() if v.get('api_key')])} 个API已配置")
            else:
                self.merge_api_configs({})
                self.target_language_var.set(DEFAULT_TARGET_LANGUAGE)
                print("ℹ️ 未找到配置文件，使用默认配置")
        except Exception as e:
            error_msg = f"加载配置失败: {e}"
            print(error_msg)
            # 尝试从备份恢复
            if self.restore_from_backup():
                print("✓ 已从备份恢复配置")
            else:
                messagebox.showwarning("警告", f"{error_msg}\n将使用默认配置")

    def restore_from_backup(self):
        """从最新备份恢复配置"""
        try:
            backup_dir = Path(__file__).parent / 'config_backups'
            if not backup_dir.exists():
                return False

            backups = sorted(backup_dir.glob('config_backup_*.json'), reverse=True)
            if not backups:
                return False

            # 使用最新的备份
            latest_backup = backups[0]
            with open(latest_backup, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                if isinstance(loaded_config, dict) and 'api_configs' in loaded_config:
                    api_config_section = loaded_config.get('api_configs', {})
                    target_language = loaded_config.get('target_language', DEFAULT_TARGET_LANGUAGE)
                else:
                    api_config_section = loaded_config if isinstance(loaded_config, dict) else {}
                    target_language = (
                        loaded_config.get('target_language', DEFAULT_TARGET_LANGUAGE)
                        if isinstance(loaded_config, dict) else DEFAULT_TARGET_LANGUAGE
                    )
                self.merge_api_configs(api_config_section)
                self.target_language_var.set(target_language or DEFAULT_TARGET_LANGUAGE)

            # 恢复成功后保存为主配置
            config_file = Path(__file__).parent / 'translator_config.json'
            with open(config_file, 'w', encoding='utf-8') as f:
                data = {
                    'api_configs': self.api_configs,
                    'target_language': self.get_target_language()
                }
                json.dump(data, f, indent=2, ensure_ascii=False)

            return True
        except Exception as e:
            print(f"从备份恢复失败: {e}")
            return False

    def on_closing(self):
        """程序退出时的处理（自动保存配置）"""
        # 如果正在翻译，询问用户
        if self.is_translating:
            if not messagebox.askyesno("确认退出", "翻译正在进行中，确定要退出吗？\n\n配置将自动保存"):
                return

        # 自动保存配置
        try:
            self.save_config(show_message=False)
            print("✓ 配置已自动保存")
        except Exception as e:
            print(f"保存配置时出错: {e}")

        # 关闭窗口
        self.root.destroy()

    def start_translation(self):
        """开始翻译"""
        if not self.current_text:
            messagebox.showwarning("警告", "请先加载要翻译的文件")
            return

        api_type = self.get_current_api_type()
        config = self.api_configs[api_type]

        if not config.get('api_key'):
            messagebox.showwarning("警告", "请先配置API Key")
            self.open_api_config()
            return

        # 计算签名用于断点恢复判断
        current_signature = self.compute_text_signature(self.current_text)
        resume_possible = (
            self.text_signature == current_signature
            and self.source_segments
            and 0 < len(self.translated_segments) < len(self.source_segments)
        )

        # 是否从断点继续
        self.resume_from_index = 0
        if resume_possible:
            resume = messagebox.askyesno(
                "继续翻译",
                f"检测到上次未完成的翻译，是否从第 {len(self.translated_segments) + 1} 段继续？"
            )
            if resume:
                self.resume_from_index = len(self.translated_segments)
                # 确保译文长度与起始段对齐
                if len(self.translated_segments) > self.resume_from_index:
                    self.translated_segments = self.translated_segments[:self.resume_from_index]
            else:
                self.translated_segments = []
                self.source_segments = []
                self.failed_segments = []
        else:
            self.translated_segments = []
            self.source_segments = []
            self.failed_segments = []

        # 开始翻译
        self.lm_studio_fallback_active = False
        self.consecutive_failures = 0
        self.paused_due_to_failures = False
        self.is_translating = True
        self.translate_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.progress_var.set(
            (self.resume_from_index / max(len(self.source_segments), 1)) * 100
            if self.resume_from_index and self.source_segments else 0
        )
        if not self.resume_from_index:
            self.translated_text = ""
            self.translated_text_widget.delete('1.0', tk.END)
        self.failed_segments = []
        self.selected_failed_index = None
        self.refresh_failed_segments_view()

        # 在新线程中执行翻译
        self.translation_thread = threading.Thread(target=self.translate_text, daemon=True)
        self.translation_thread.start()

    def stop_translation(self):
        """停止翻译"""
        self.is_translating = False
        self.translate_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress_text_var.set("翻译已停止")

    def translate_text(self):
        """执行翻译（在后台线程中）"""
        try:
            # 获取当前API类型
            api_type = self.get_current_api_type()
            self.consecutive_failures = 0

            # 准备
            self.root.after(0, self.progress_text_var.set, "正在进行文本分段...")

            # 使用 FileProcessor 进行分段
            self.source_segments = self.file_processor.split_text_into_segments(self.current_text, max_length=800)
            total_segments = len(self.source_segments)
            self.text_signature = self.compute_text_signature(self.current_text)
            start_index = min(self.resume_from_index or 0, total_segments)

            self.root.after(0, self.progress_text_var.set, f"文本已分为 {total_segments} 段，准备开始翻译...")
            if start_index:
                self.root.after(
                    0,
                    self.progress_var.set,
                    (start_index / total_segments) * 100 if total_segments else 0
                )
                self.root.after(
                    0,
                    self.progress_text_var.set,
                    f"继续翻译：从第 {start_index + 1} 段开始..."
                )

            # 逐段翻译
            for idx in range(start_index, total_segments):
                segment = self.source_segments[idx]
                if not self.is_translating:
                    break

                # 更新进度
                progress = (idx / total_segments) * 100
                self.root.after(0, self.progress_var.set, progress)
                self.root.after(
                    0,
                    self.progress_text_var.set,
                    f"正在翻译... {idx + 1}/{total_segments} 段"
                )

                try:
                    translated = self.translate_segment(api_type, segment)
                    self.translated_segments.append(translated)

                    # 实时更新译文显示
                    self.translated_text = "\n\n".join(self.translated_segments)
                    self.root.after(
                        0,
                        self.update_translated_text,
                        self.translated_text
                    )
                    self.consecutive_failures = 0
                    self.save_progress_cache()

                except Exception as e:
                    self.consecutive_failures += 1
                    error_msg = f"[翻译错误: {str(e)}]\n{segment}"
                    # 达到阈值则暂停，保留进度
                    if self.consecutive_failures >= self.max_consecutive_failures:
                        self.paused_due_to_failures = True
                        self.resume_from_index = idx
                        self.save_progress_cache()
                        self.root.after(
                            0,
                            self.progress_text_var.set,
                            "API连续失败，已暂停。稍后点击“开始翻译”可从当前进度继续。"
                        )
                        print(f"连续失败{self.consecutive_failures}次，暂停在段 {idx + 1}: {e}")
                        break
                    else:
                        self.translated_segments.append(error_msg)
                        print(f"翻译段落 {idx + 1} 失败: {e}")

                # 稍微延迟避免API限流
                time.sleep(0.5)

            # 翻译完成
            if self.is_translating and not self.paused_due_to_failures:
                self.root.after(0, self.progress_text_var.set, "正在检查译文...")
                self.verify_and_retry_segments(api_type)

                self.translated_text = "\n\n".join(self.translated_segments)
                self.root.after(0, self.update_translated_text, self.translated_text)
                self.root.after(0, self.refresh_failed_segments_view)

                self.root.after(0, self.progress_var.set, 100)
                status_msg = (
                    f"翻译完成，但有 {len(self.failed_segments)} 段需要处理"
                    if self.failed_segments else "翻译完成!"
                )
                self.root.after(0, self.progress_text_var.set, status_msg)
                self.root.after(0, self.on_translation_complete)
                if not self.failed_segments:
                    self.clear_progress_cache()
            else:
                status_msg = "翻译已停止"
                if self.paused_due_to_failures:
                    status_msg = "已暂停，等待API恢复后可继续"
                self.root.after(0, self.progress_text_var.set, status_msg)

        except Exception as e:
            self.root.after(
                0,
                messagebox.showerror,
                "错误",
                f"翻译过程中出错:\n{str(e)}"
            )
        finally:
            self.root.after(0, self.translate_btn.config, {'state': 'normal'})
            self.root.after(0, self.stop_btn.config, {'state': 'disabled'})
            self.is_translating = False

    def detect_language(self, text):
        """简单的语言检测：检查是否主要是中文"""
        if not text or len(text.strip()) == 0:
            return 'unknown'

        # 统计中文字符
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        total_chars = len(re.findall(r'\S', text))

        if total_chars == 0:
            return 'unknown'

        chinese_ratio = chinese_chars / total_chars

        # 如果中文占比超过60%，认为是中文
        if chinese_ratio > 0.6:
            return 'zh'
        # 如果中文占比很低，可能是英文或其他语言
        elif chinese_ratio < 0.1:
            return 'en'
        else:
            return 'mixed'

    def translate_segment(self, api_type, text):
        """按当前API类型翻译单段文本（支持自动回退到本地模型）"""
        target_language = self.get_target_language()
        target_is_chinese = self.is_target_language_chinese(target_language)
        target_is_english = self.is_target_language_english(target_language)

        # 检测语言，如果已经是目标语言就跳过翻译
        lang = self.detect_language(text)
        if (target_is_chinese and lang == 'zh') or (target_is_english and lang == 'en'):
            print(f"✓ 检测到已是目标语言（{target_language}），跳过翻译（{len(text)} 字符）")
            return text  # 直接返回原文

        # 如果已经切换到本地备用方案，直接使用本地模型
        if self.lm_studio_fallback_active and api_type != 'lm_studio':
            return self.translate_with_lm_studio(text)

        # 首先尝试使用API翻译
        try:
            if api_type == 'gemini':
                return self.translate_with_gemini(text)
            elif api_type == 'openai':
                return self.translate_with_openai(text)
            elif api_type == 'custom':
                return self.translate_with_custom_api(text)
            elif api_type == 'lm_studio':
                return self.translate_with_lm_studio(text)
            else:
                raise ValueError("不支持的API类型")

        except Exception as api_error:
            # 检查是否为配额/限流错误
            error_msg = str(api_error).lower()
            is_quota_error = any(keyword in error_msg for keyword in [
                'quota', 'rate limit', 'resource_exhausted', '429',
                'insufficient_quota', 'quota exceeded', 'rate_limit_exceeded'
            ])

            # 如果是配额错误则切换到本地LM Studio备用方案
            if is_quota_error and api_type != 'lm_studio':
                try:
                    print("⚠️ API配额用尽，切换到本地LM Studio...")
                    self.lm_studio_fallback_active = True
                    self.root.after(
                        0,
                        self.progress_text_var.set,
                        "API配额用尽，已切换到本地LM Studio模型..."
                    )
                    return self.translate_with_lm_studio(text)
                except Exception as lm_error:
                    print(f"✗ 本地模型调用失败: {lm_error}")
                    raise Exception(f"本地备用翻译失败: {lm_error}") from api_error
            else:
                # 非配额错误直接抛出
                raise api_error

    def is_translation_incomplete(self, translated, source, target_language=None):
        """检测译文是否异常或未完成"""
        target_language = target_language or self.get_target_language()
        target_is_chinese = self.is_target_language_chinese(target_language)
        target_is_english = self.is_target_language_english(target_language)

        if not translated or not translated.strip():
            return True

        normalized = translated.strip()
        if normalized.startswith("[翻译错误") or normalized.startswith("[未翻译") or normalized.startswith("[待手动翻译"):
            return True

        # 明显过短或与原文相同视为未完成
        if len(normalized) < 5:
            return True
        if normalized == source.strip():
            return True

        min_length_ratio = 0.2 if target_is_chinese else 0.15
        if len(source) > 50 and len(normalized) < len(source) * min_length_ratio:
            return True

        # 语言/字符占比检查：译文缺少中文或仍以英文/日文为主则视为未完成
        def count_chars(text, pattern):
            return len(re.findall(pattern, text))

        chinese_chars = count_chars(normalized, r'[\u4e00-\u9fff]')
        latin_chars = count_chars(normalized, r'[A-Za-z]')
        japanese_chars = count_chars(normalized, r'[\u3040-\u30ff\u31f0-\u31ff]')
        total_chars = len(re.findall(r'\S', normalized)) or 1  # 避免除0

        chinese_ratio = chinese_chars / total_chars
        latin_ratio = latin_chars / total_chars
        japanese_ratio = japanese_chars / total_chars

        source_has_latin = bool(re.search(r'[A-Za-z]', source))
        source_has_japanese = bool(re.search(r'[\u3040-\u30ff\u31f0-\u31ff]', source))

        if target_is_chinese:
            # 原文是英文/日文，且译文中文比例低，则判定未完成
            if source_has_latin and chinese_ratio < 0.2:
                return True
            if source_has_japanese and chinese_ratio < 0.2:
                return True

            # 译文整体缺少中文且仍以英文/日文为主
            if chinese_ratio < 0.15 and (latin_ratio > 0.35 or japanese_ratio > 0.2):
                return True

            # 明显以英文或日文占主导也视为未完成
            if latin_ratio > 0.6 or japanese_ratio > 0.3:
                return True
        elif target_is_english:
            # 英文目标时，如果译文仍以中文为主或明显过短则视为未完成
            if chinese_ratio > 0.3 and chinese_ratio > latin_ratio:
                return True
            if latin_ratio < 0.15 and len(source) > 50:
                return True
        else:
            # 其他目标语言：只做基础完整性检查，避免误判
            if chinese_ratio > 0.6 and target_language:
                return True

        return False

    def verify_and_retry_segments(self, api_type):
        """翻译完成后检查并自动重试失败段落"""
        failed = []
        target_language = self.get_target_language()
        for idx, (source, translated) in enumerate(zip(self.source_segments, self.translated_segments)):
            if self.is_translation_incomplete(translated, source, target_language=target_language):
                try:
                    retry_text = self.translate_segment(api_type, source)
                except Exception as e:
                    retry_text = f"[翻译错误: {str(e)}]\n{source}"

                if not self.is_translation_incomplete(retry_text, source, target_language=target_language):
                    self.translated_segments[idx] = retry_text
                else:
                    placeholder = f"[待手动翻译 - 段 {idx + 1}]"
                    self.translated_segments[idx] = placeholder
                    failed.append({
                        'index': idx,
                        'source': source,
                        'last_error': translated
                    })

        self.failed_segments = failed

    def refresh_failed_segments_view(self):
        """刷新失败段落列表和状态"""
        if not hasattr(self, 'failed_listbox'):
            return

        self.failed_listbox.delete(0, tk.END)
        self.selected_failed_index = None

        self.failed_source_text.config(state='normal')
        self.failed_source_text.delete('1.0', tk.END)
        self.failed_source_text.config(state='disabled')

        self.manual_translation_text.delete('1.0', tk.END)

        if not self.failed_segments:
            self.failed_status_var.set("暂无失败段落")
            return

        for item in self.failed_segments:
            snippet = item['source'].replace("\n", " ")
            if len(snippet) > 60:
                snippet = snippet[:60] + "..."
            self.failed_listbox.insert(tk.END, f"段 {item['index'] + 1}: {snippet}")

        self.failed_status_var.set(f"待处理段落: {len(self.failed_segments)} 个")

    def on_failed_select(self, event=None):
        """选中失败段落时展示详情"""
        if not self.failed_segments:
            return

        selection = self.failed_listbox.curselection()
        if not selection:
            return

        idx = selection[0]
        self.selected_failed_index = idx
        info = self.failed_segments[idx]

        self.failed_source_text.config(state='normal')
        self.failed_source_text.delete('1.0', tk.END)
        self.failed_source_text.insert('1.0', info['source'])
        self.failed_source_text.config(state='disabled')

        self.manual_translation_text.delete('1.0', tk.END)

    def retry_failed_segment(self):
        """对选中失败段落重新翻译"""
        if self.selected_failed_index is None or not self.failed_segments:
            messagebox.showinfo("提示", "请先选择需要重试的段落")
            return

        info = self.failed_segments[self.selected_failed_index]
        api_type = self.get_current_api_type()

        try:
            retry_text = self.translate_segment(api_type, info['source'])
        except Exception as e:
            messagebox.showerror("错误", f"重试翻译失败: {e}")
            return

        if self.is_translation_incomplete(retry_text, info['source'], target_language=self.get_target_language()):
            messagebox.showwarning("提示", "重试后仍未完成，请手动翻译。")
            return

        self.translated_segments[info['index']] = retry_text
        self.failed_segments.pop(self.selected_failed_index)
        self.rebuild_translated_text()
        self.refresh_failed_segments_view()
        messagebox.showinfo("成功", f"段 {info['index'] + 1} 已重新翻译并替换")

    def save_manual_translation(self):
        """将手动译文写回对应段落"""
        if self.selected_failed_index is None or not self.failed_segments:
            messagebox.showinfo("提示", "请先选择需要替换的段落")
            return

        manual_text = self.manual_translation_text.get('1.0', tk.END).strip()
        if not manual_text:
            messagebox.showwarning("警告", "手动翻译内容不能为空")
            return

        info = self.failed_segments[self.selected_failed_index]
        self.translated_segments[info['index']] = manual_text
        self.failed_segments.pop(self.selected_failed_index)
        self.rebuild_translated_text()
        self.refresh_failed_segments_view()
        messagebox.showinfo("成功", f"段 {info['index'] + 1} 已使用手动译文替换")
        self.save_progress_cache()

    def rebuild_translated_text(self):
        """根据分段译文重建完整译文"""
        self.translated_text = "\n\n".join(self.translated_segments) if self.translated_segments else ""
        self.update_translated_text(self.translated_text)

    def translate_with_gemini(self, text):
        """使用Gemini API翻译"""
        config = self.api_configs['gemini']
        genai.configure(api_key=config['api_key'])

        model = genai.GenerativeModel(config['model'])
        target_language = self.get_target_language()
        prompt = f"请将以下文本翻译成{target_language}，保持原文的格式和段落结构:\n\n{text}"

        response = model.generate_content(prompt)
        return response.text

    def translate_with_openai(self, text):
        """使用OpenAI API翻译"""
        config = self.api_configs['openai']
        target_language = self.get_target_language()

        client_kwargs = {'api_key': config['api_key']}
        if config.get('base_url'):
            client_kwargs['base_url'] = config['base_url']

        client = openai.OpenAI(**client_kwargs)

        response = client.chat.completions.create(
            model=config['model'],
            messages=[
                {
                    "role": "system",
                    "content": f"你是一个专业的翻译助手，请将用户提供的文本翻译成{target_language}，保持原文的格式和段落结构。"
                },
                {"role": "user", "content": text}
            ]
        )

        return response.choices[0].message.content

    def translate_with_custom_api(self, text):
        """使用自定义API翻译（OpenAI兼容格式）"""
        config = self.api_configs['custom']
        target_language = self.get_target_language()

        if not config.get('base_url'):
            raise ValueError("自定义API需要配置Base URL")

        url = f"{config['base_url'].rstrip('/')}/chat/completions"

        headers = {
            'Content-Type': 'application/json',
            'Authorization': f"Bearer {config['api_key']}"
        }

        data = {
            'model': config['model'],
            'messages': [
                {
                    "role": "system",
                    "content": f"你是一个专业的翻译助手，请将用户提供的文本翻译成{target_language}，保持原文的格式和段落结构。"
                },
                {"role": "user", "content": text}
            ]
        }

        response = requests.post(url, headers=headers, json=data, timeout=60)
        response.raise_for_status()

        result = response.json()
        return result['choices'][0]['message']['content']

    def translate_with_lm_studio(self, text):
        """使用本地LM Studio模型翻译（OpenAI兼容接口）"""
        if not OPENAI_SUPPORT:
            raise ImportError("缺少 openai 库，无法调用本地LM Studio")

        config = self.api_configs.get('lm_studio', {})
        target_language = self.get_target_language()

        client = openai.OpenAI(
            api_key=config.get('api_key') or DEFAULT_LM_STUDIO_CONFIG['api_key'],
            base_url=config.get('base_url') or DEFAULT_LM_STUDIO_CONFIG['base_url']
        )

        response = client.chat.completions.create(
            model=config.get('model') or DEFAULT_LM_STUDIO_CONFIG['model'],
            messages=[
                {
                    "role": "system",
                    "content": f"你是一个专业的翻译助手，请将用户提供的文本翻译成{target_language}，保持原文的格式和段落结构。"
                },
                {"role": "user", "content": text}
            ],
            temperature=0.2
        )

        return response.choices[0].message.content

    def update_translated_text(self, text):
        """更新译文显示"""
        self.translated_text_widget.delete('1.0', tk.END)
        self.translated_text_widget.insert('1.0', text)
        # 自动滚动到底部
        self.translated_text_widget.see(tk.END)

    def on_translation_complete(self):
        """翻译完成后的处理"""
        if self.failed_segments:
            self.notebook.select(2)  # 切换到失败段落标签页
            message = f"翻译完成，但 {len(self.failed_segments)} 个段落需要手动翻译或重试。"
            messagebox.showwarning("完成", message)
        else:
            self.notebook.select(1)  # 切换到译文标签页
            messagebox.showinfo("完成", "翻译已完成!")

    def export_translation(self):
        """导出翻译结果"""
        if not self.translated_text or not self.translated_text.strip():
            messagebox.showwarning("警告", "没有可导出的译文\n\n请先完成翻译后再导出")
            return

        # 建议默认文件名
        original_file = self.file_path_var.get()
        target_language = self.get_target_language()
        safe_lang = re.sub(r'[\\/:*?"<>|]', "_", target_language).strip() or "译文"
        if original_file:
            base_name = Path(original_file).stem
            default_name = f"{base_name}_{safe_lang}译文.txt"
        else:
            default_name = f"{safe_lang}译文.txt"

        filename = filedialog.asksaveasfilename(
            title="保存译文",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )

        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.translated_text)

                # 统计信息
                char_count = len(self.translated_text)
                word_count = len(self.translated_text.split())

                messagebox.showinfo(
                    "成功",
                    f"译文已保存到:\n{filename}\n\n"
                    f"字符数: {char_count:,}\n"
                    f"词数: {word_count:,}"
                )
                # 完整导出后清除进度缓存
                if self.source_segments and len(self.translated_segments) == len(self.source_segments):
                    self.clear_progress_cache()
            except Exception as e:
                messagebox.showerror("错误", f"保存文件失败:\n{str(e)}")

    def update_text_display(self):
        """更新文本显示（预览或完整）"""
        if not self.current_text:
            return

        self.original_text.delete('1.0', tk.END)

        char_count = len(self.current_text)
        is_large_file = char_count > self.preview_limit

        if is_large_file and not self.show_full_text:
            # 显示预览
            preview_text = self.current_text[:self.preview_limit]
            preview_text += f"\n\n{'='*60}\n"
            preview_text += f"⚠️ 预览模式：仅显示前 {self.preview_limit:,} / {char_count:,} 字符\n"
            preview_text += f"点击上方'显示完整原文'按钮查看全文\n"
            preview_text += f"{'='*60}"
            self.original_text.insert('1.0', preview_text)
        else:
            # 显示完整文本
            self.original_text.insert('1.0', self.current_text)

    def toggle_full_text_display(self):
        """切换显示完整文本或预览"""
        if not self.current_text:
            return

        self.show_full_text = not self.show_full_text
        char_count = len(self.current_text)

        if self.show_full_text:
            # 切换到完整显示
            self.toggle_preview_btn.config(text="仅显示预览")
            self.file_info_var.set(f"✓ 显示完整文件 ({char_count:,} 字符)")
            self.progress_text_var.set("正在加载完整文本...")
            self.root.update()

            # 使用after延迟更新，避免界面冻结
            self.root.after(100, self._update_full_text)
        else:
            # 切换到预览
            self.toggle_preview_btn.config(text="显示完整原文")
            self.file_info_var.set(
                f"⚠️ 大文件 ({char_count:,} 字符) - 仅显示前 {self.preview_limit:,} 字符"
            )
            self.update_text_display()
            self.progress_text_var.set(f"已加载文件 | 字符数: {char_count:,}")

    def _update_full_text(self):
        """更新完整文本（在延迟后执行）"""
        self.update_text_display()
        char_count = len(self.current_text)
        word_count = len(self.current_text.split())
        self.progress_text_var.set(
            f"已加载完整文件 | 字符数: {char_count:,} | 词数: {word_count:,}"
        )

    def clear_all(self):
        """清空所有内容"""
        if messagebox.askyesno("确认", "确定要清空所有内容吗?"):
            self.file_path_var.set("")
            self.current_text = ""
            self.translated_text = ""
            self.source_segments = []
            self.translated_segments = []
            self.failed_segments = []
            self.selected_failed_index = None
            self.show_full_text = False
            self.original_text.delete('1.0', tk.END)
            self.translated_text_widget.delete('1.0', tk.END)
            self.progress_var.set(0)
            self.progress_text_var.set("就绪")
            self.file_info_var.set("")
            self.toggle_preview_btn.config(state='disabled', text="显示完整原文")
            self.refresh_failed_segments_view()


def main():
    """主程序入口"""
    # 检查依赖
    missing_libs = []
    if not PDF_SUPPORT:
        missing_libs.append("PyPDF2 (用于PDF支持)")
    if not EPUB_SUPPORT:
        missing_libs.append("ebooklib, beautifulsoup4 (用于EPUB支持)")
    if not GEMINI_SUPPORT:
        missing_libs.append("google-generativeai (用于Gemini API)")
    if not OPENAI_SUPPORT:
        missing_libs.append("openai (用于OpenAI API)")
    if not REQUESTS_SUPPORT:
        missing_libs.append("requests (用于自定义API)")

    if missing_libs:
        print("=" * 60)
        print("警告: 以下库未安装，部分功能将不可用:")
        for lib in missing_libs:
            print(f"  - {lib}")
        print("\n安装命令:")
        print("pip install PyPDF2 ebooklib beautifulsoup4 google-generativeai openai requests")
        print("=" * 60)
        print()

    root = tk.Tk()
    app = BookTranslatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
