# -*- coding: utf-8 -*-
"""
自动解压工具（含GUI）v7.1（修复版）
修复：
- 修正 7z 魔数字节写法，引发的 “TypeError: can't concat str to bytes”。
- 修正多个正则表达式中误用的 '\\s'（被当成字面量），改回 \s。
- 其余功能与 v7 相同：当前进度/阶段显示、心跳、测试与解压的明确日志等。
"""
import os
import re
import sys
import time
import threading
import queue
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Optional, Tuple, List, Iterable, Dict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# --------------------------- 工具函数 ---------------------------

KNOWN_EXTS = {'.zip', '.7z', '.rar', '.001', '.z01'}

# 目录名里常见的“解压密码：xxx”推断
PWD_PREFIX_RE = re.compile(r'(解压码|解压密码|密码)(统一为|为|是)?\s*[：:\s]\s*(.+)')
PWD_INLINE_RE = re.compile(r'(解压码|解压密码|压缩密码|提取码|密码|pw|pass|password|key)[：:\s=]*([^\s\]\\/:<>\"\'`]+)', re.I)
PWD_BRACKET_RE = re.compile(r'[\[(（【]\s*(?:pwd|password|pass|密码|解压码|提取码)[：:\s=]*([^\]\)）】\s]+)', re.I)
PWD_HINT_EXTS = {'.txt', '.md', '.nfo', '.url', '.ini'}

LANG_TEXT = {
    'zh': {
        'title': "自动解压工具 v7.1",
        'frame_basic': "基本设置",
        'tab_scan': "扫描并解压",
        'tab_list': "仅扫描（选择后解压）",
        'tab_help': "说明 / Help",
        'scan_desc': "此模式：扫描后立即按设置解压所有发现的压缩包。",
        'start_all': "开始解压（全量）",
        'stop': "停止",
        'scan': "扫描",
        'filter_kw': "过滤关键词：",
        'size_label': "大小(MB)：",
        'to': "至",
        'apply_filter': "应用过滤",
        'export': "导出列表",
        'extract_sel': "解压选中",
        'select_all': "全选",
        'select_none': "全不选",
        'listed': "已列出：",
        'help_body': (
            "功能概要：\n"
            "1. 扫描压缩包（zip/7z/rar/分卷），支持递归/过滤/排序。\n"
            "2. 预估密码：从文件名、目录名、同目录提示文件中提取。\n"
            "3. 解压：Bandizip/7-Zip，预测试、失败自动切换、二次解压、删源包可选。\n"
            "4. 并发：列表模式支持并发解压，带心跳和目录增长监控。\n"
            "5. 右键操作：打开目录、删除文件、移除列表、收藏、批量更正密码、复制到指定目录。\n"
            "6. 表格交互：可勾选批量解压，双击密码单元格直接编辑，Ctrl+C 复制单元格。\n"
            "7. 过滤：关键词 + 大小区间（MB）。\n\n"
            "使用提示：\n"
            "- 先设定 Bandizip/7-Zip 路径；找不到会自动尝试常见安装路径。\n"
            "- 推荐先“仅扫描”，在列表中勾选需要的文件，再“解压选中”。\n"
            "- 勾选优先于选中：若存在勾选，解压仅处理勾选项。\n"
            "- “完成后动作”默认无，可切换为退出或关机（停止时不会执行）。\n"
            "- 导出列表会带出勾选/收藏状态和推断密码。\n"
        )
    },
    'en': {
        'title': "Auto Extractor v7.1",
        'frame_basic': "Basic Settings",
        'tab_scan': "Scan & Extract",
        'tab_list': "Scan Only (Pick to Extract)",
        'tab_help': "Guide / Help",
        'scan_desc': "This mode scans then extracts every found archive immediately.",
        'start_all': "Start (all)",
        'stop': "Stop",
        'scan': "Scan",
        'filter_kw': "Filter keyword:",
        'size_label': "Size (MB):",
        'to': "to",
        'apply_filter': "Apply filter",
        'export': "Export list",
        'extract_sel': "Extract selected",
        'select_all': "Select all",
        'select_none': "Clear selection",
        'listed': "Listed: ",
        'help_body': (
            "Overview:\n"
            "1) Scan archives (zip/7z/rar/multi-part) with recurse/filter/sort.\n"
            "2) Password guess: from filename, parent folder, hint files nearby.\n"
            "3) Extract via Bandizip/7-Zip; pre-test, fallback, nested extract, optional delete source.\n"
            "4) Concurrency: list-mode extraction runs in parallel with heartbeat and output-dir growth monitor.\n"
            "5) Context menu: open folder, delete file, remove row, favorite, bulk password fix, copy to folder.\n"
            "6) Table: checkboxes for batch extract, double-click password cell to edit, Ctrl+C copies a cell.\n"
            "7) Filters: keyword + size range (MB).\n\n"
            "Tips:\n"
            "- Set Bandizip/7-Zip path first; common install paths are auto-detected.\n"
            "- Use 'Scan Only' first, check the needed items, then 'Extract selected'.\n"
            "- Checked rows take priority: if any checked, extraction uses those only.\n"
            "- 'After finish' action defaults to none; can exit or shutdown (skipped when stopped).\n"
            "- Exported list includes check/favorite states and guessed passwords.\n"
        )
    }
}

MAGIC_SIGS = {
    'zip': [b'PK\x03\x04', b'PK\x05\x06', b'PK\x07\x08'],
    '7z':  [b'7z\xBC\xAF\x27\x1C'],  # 正确的 7z 文件头
    'rar': [b'Rar!\x1A\x07\x00', b'Rar!\x1A\x07\x01\x00'],
    'html': [b'<!DOCTYP', b'<html', b'<HTML'],
    'xml': [b'<?xml'],
    'pdf': [b'%PDF'],
}

def human(n: int) -> str:
    units = ['B','KB','MB','GB','TB']
    s = 0
    f = float(n)
    while f >= 1024 and s < len(units)-1:
        f /= 1024.0; s += 1
    return f'{f:.1f}{units[s]}'

def file_size(path: Path) -> int:
    try:
        return path.stat().st_size
    except Exception:
        return 0

def find_on_path(names: Iterable[str]) -> Optional[str]:
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    candidates = []
    program_files = os.environ.get('ProgramFiles', r'C:\Program Files')
    program_files_x86 = os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)')
    candidates += [
        rf'F:\Bandizip\bz.exe',
        rf'{program_files}\Bandizip\bz.exe',
        rf'{program_files_x86}\Bandizip\bz.exe',
        rf'{program_files}\7-Zip\7z.exe',
        rf'{program_files_x86}\7-Zip\7z.exe',
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None

def normalize_extension(p: Path) -> Path:
    if p.suffix.lower() in KNOWN_EXTS:
        return p
    ext = p.suffix.lower()
    if not ext:
        return p
    cleaned = re.sub(r'[^0-9a-z]', '', ext)
    target = None
    if 'rar' in cleaned:
        target = '.rar'
    elif '7z' in cleaned:
        target = '.7z'
    elif 'zip' in cleaned:
        target = '.zip'
    if target:
        newp = p.with_suffix(target)
        try:
            p.rename(newp)
            return newp
        except Exception:
            return p
    return p

def is_multipart_first(archive: Path) -> Tuple[bool, bool]:
    name = archive.name.lower()
    if re.search(r'\.part0*1\.rar$', name) or re.search(r'\.part1\.rar$', name):
        return True, True
    if re.search(r'\.part\d+\.rar$', name):
        return True, False
    if name.endswith('.7z.001') or name.endswith('.zip.001'):
        return True, True
    if name.endswith('.001'):
        return True, True  # 兜底按首卷处理
    if name.endswith('.z01'):
        return True, True
    if re.search(r'\.z\d{2}$', name):
        return True, False
    return False, False

def derive_password_from_dir(dirname: str) -> str:
    m = PWD_PREFIX_RE.search(dirname.strip())
    if m:
        return m.group(3).strip()
    return dirname.strip()

def _clean_pwd(pwd: str) -> str:
    return pwd.strip().strip('，。,:：;；)]}】）')

def _extract_pwd_from_text(text: str) -> Optional[str]:
    for pat in (PWD_PREFIX_RE, PWD_BRACKET_RE, PWD_INLINE_RE):
        m = pat.search(text)
        if m:
            val = m.group(m.lastindex).strip() if m.lastindex else m.group(1).strip()
            if val:
                return _clean_pwd(val)
    return None

def infer_password(arc: Path, cache: Dict[str, Optional[str]]= {}) -> Optional[str]:
    """多策略推断密码：文件名 -> 父目录名 -> 目录内提示文件。"""
    # 1) 文件名/无后缀名
    for blob in (arc.name, arc.stem):
        pwd = _extract_pwd_from_text(blob)
        if pwd:
            return pwd
    # 2) 父目录名
    pwd = _extract_pwd_from_text(arc.parent.name)
    if pwd:
        return pwd
    # 3) 目录提示文件（缓存避免重复读）
    dir_key = str(arc.parent.resolve())
    if dir_key in cache:
        return cache[dir_key]
    for f in arc.parent.iterdir():
        if not f.is_file():
            continue
        if f.suffix.lower() not in PWD_HINT_EXTS:
            continue
        try:
            if f.stat().st_size > 64 * 1024:  # 避免大文件
                continue
            content = f.read_text('utf-8', errors='ignore')[:4000]
        except Exception:
            continue
        pwd = _extract_pwd_from_text(content)
        if pwd:
            cache[dir_key] = pwd
            return pwd
    cache[dir_key] = None
    return None

def gather_archives(root: Path, recursive: bool=True) -> List[Path]:
    found = []
    if recursive:
        walker = os.walk(root)
    else:
        walker = [(root, [], [f for f in os.listdir(root) if (root/f).is_file()])]
    for dirpath, _, files in walker:
        for f in files:
            p = Path(dirpath) / f
            p = normalize_extension(p)
            low = p.name.lower()
            if any([low.endswith('.zip'), low.endswith('.7z'), low.endswith('.rar'),
                    low.endswith('.001'), low.endswith('.z01'),
                    re.search(r'\.part\d+\.rar$', low) is not None]):
                is_multi, is_first = is_multipart_first(p)
                if is_multi and not is_first:
                    continue
                found.append(p)
    return found

def sniff_signature(path: Path, read_len: int = 8) -> str:
    try:
        with open(path, 'rb') as f:
            head = f.read(read_len)
    except Exception:
        return 'unknown'
    for kind, sigs in MAGIC_SIGS.items():
        for sig in sigs:
            if head.startswith(sig):
                return kind
    return 'unknown'

def overwrite_flag(policy: str) -> str:
    return {'skip': '-aos', 'rename': '-aou', 'overwrite': '-aoa'}[policy]

def bandizip_cmd(bz: str, archive: Path, outdir: Path, password: Optional[str], policy: str) -> list:
    cmd = [bz, 'x', f'-cp:65001', overwrite_flag(policy), f'-o:{str(outdir)}']
    if password:
        cmd.insert(2, f'-p:{password}')
    cmd.append(str(archive))
    return cmd

def bandizip_test_cmd(bz: str, archive: Path, password: Optional[str]) -> list:
    cmd = [bz, 't']
    if password:
        cmd.append(f'-p:{password}')
    cmd.append(str(archive))
    return cmd

def sevenzip_cmd(sz: str, archive: Path, outdir: Path, password: Optional[str], policy: str) -> list:
    # 传入空密码以禁止 7z 交互式等待
    pwd = '' if password is None else password
    cmd = [sz, 'x', f'-o{str(outdir)}', overwrite_flag(policy), f'-p{pwd}', '-y']
    cmd.append(str(archive))
    return cmd

def sevenzip_test_cmd(sz: str, archive: Path, password: Optional[str]) -> list:
    # 传入空密码以禁止 7z 交互式等待
    pwd = '' if password is None else password
    cmd = [sz, 't', f'-p{pwd}', '-y']
    cmd.append(str(archive))
    return cmd

def get_all_multipart_siblings(first_part: Path) -> list:
    name = first_part.name
    parent = first_part.parent
    siblings = []
    if name.lower().endswith('.7z.001') or name.lower().endswith('.zip.001'):
        stem = name[:-4]
        for p in parent.glob(stem + '.*'):
            if re.match(r'.*\.(\d{3})$', p.name.lower()):
                siblings.append(p)
    elif name.lower().endswith('.z01'):
        base = name[:-3]
        for p in parent.glob(base + 'z*'):
            siblings.append(p)
    else:
        m = re.match(r'(.+?)\.part0*1\.rar$', name, flags=re.I) or re.match(r'(.+?)\.part1\.rar$', name, flags=re.I)
        if m:
            prefix = m.group(1)
            for p in parent.glob(prefix + '.part*.rar'):
                siblings.append(p)
    if first_part not in siblings:
        siblings.append(first_part)
    return siblings

def dir_size_bytes(path: Path) -> int:
    total = 0
    if not path.exists():
        return 0
    for dp, _, files in os.walk(path):
        for f in files:
            try:
                total += (Path(dp)/f).stat().st_size
            except Exception:
                pass
    return total

def run_cmd(cmd: list, log, stop_flag: threading.Event,
            monitor_dir: Optional[Path] = None, quiet_limit: int = 30, phase_name: str = '') -> int:
    """
    统一执行子进程：
    - 实时读取 stdout 并写入日志；
    - 即使 monitor_dir=None（如测试阶段），也会每 quiet_limit 秒输出一次心跳；
    - 如设置 monitor_dir，则同时监控目录尺寸变化。
    """
    last_activity = time.time()
    try:
        p = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='ignore'
        )
        mon_stop = threading.Event()

        def monitor():
            last_sz = -1
            nonlocal last_activity
            while not mon_stop.is_set():
                if p.poll() is not None:
                    break
                now = time.time()
                # 目录尺寸变化监控
                if monitor_dir is not None:
                    try:
                        sz = dir_size_bytes(monitor_dir)
                        if sz != last_sz:
                            log(f"  · 目标目录大小 {human(sz)}")
                            last_sz = sz
                            last_activity = now
                    except Exception:
                        pass
                # 心跳（无论是否有 monitor_dir）
                if now - last_activity >= quiet_limit:
                    tag = f"（阶段：{phase_name}）" if phase_name else ""
                    log(f"  … {quiet_limit}s 未见输出{tag}，仍在等待子进程完成")
                    last_activity = now
                time.sleep(2)

        t_mon = threading.Thread(target=monitor, daemon=True)
        t_mon.start()

        while True:
            if stop_flag.is_set():
                p.terminate()
                mon_stop.set()
                return -1
            line = p.stdout.readline()
            if not line:
                break
            last_activity = time.time()
            log(line.rstrip('\n'))
        p.wait()
        mon_stop.set()
        t_mon.join(timeout=5)
        return p.returncode
    except FileNotFoundError:
        return 9001
    except Exception as e:
        log(f'!! 执行出错: {e}')
        return 9002

# --------------------------- GUI 应用 ---------------------------

class AutoExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("自动解压工具 v7.1")
        self.geometry("1150x820")
        self.minsize(1020, 720)

        self.queue = queue.Queue()
        self.stop_flag = threading.Event()
        self.worker: Optional[threading.Thread] = None

        # 公共设置
        self.lang = tk.StringVar(value='zh')
        self.var_root = tk.StringVar()
        self.var_out = tk.StringVar()
        self.var_bz = tk.StringVar(value=find_on_path(['bz.exe']) or '')
        self.var_7z = tk.StringVar(value=find_on_path(['7z.exe']) or '')
        self.var_recursive = tk.BooleanVar(value=True)
        self.var_delete = tk.BooleanVar(value=False)
        self.var_nested = tk.BooleanVar(value=True)
        self.var_pretest = tk.BooleanVar(value=True)
        self.var_cross_try = tk.BooleanVar(value=True)

        self.var_policy = tk.StringVar(value='skip')  # skip/rename/overwrite
        self.var_quiet = tk.IntVar(value=30)
        self.var_end_action = tk.StringVar(value='none')  # none/exit/shutdown

        # 仅扫描页：过滤 & 排序 & 并行
        self.var_filter = tk.StringVar()
        self.var_size_min = tk.StringVar()
        self.var_size_max = tk.StringVar()
        self.var_workers = tk.IntVar(value=3)
        self.scan_rows: List[Dict] = []
        self.bytes_map: Dict[str, int] = {}
        self.checked_map: Dict[str, bool] = {}
        self.favorite_map: Dict[str, bool] = {}
        self.sort_state = {'col': 'name', 'desc': False}

        self._build_ui()
        self.after(100, self._drain_queue)

    def _build_ui(self):
        # 顶部设置（两个模式公用）
        frm_top = ttk.LabelFrame(self, text="基本设置"); self.frm_top = frm_top
        frm_top.pack(fill='x', padx=12, pady=10)

        row_lang = ttk.Frame(frm_top); row_lang.pack(fill='x', pady=4)
        ttk.Label(row_lang, text="语言 / Language：").pack(side='left')
        ttk.Radiobutton(row_lang, text="中文", variable=self.lang, value='zh', command=self._apply_lang).pack(side='left', padx=4)
        ttk.Radiobutton(row_lang, text="English", variable=self.lang, value='en', command=self._apply_lang).pack(side='left')

        row1 = ttk.Frame(frm_top); row1.pack(fill='x', pady=6)
        ttk.Label(row1, text="扫描根目录：").grid(row=0, column=0, sticky='w')
        ttk.Entry(row1, textvariable=self.var_root, width=70).grid(row=0, column=1, sticky='we', padx=6)
        ttk.Button(row1, text="选择文件夹", command=self.choose_root).grid(row=0, column=2)
        row1.columnconfigure(1, weight=1)

        row2 = ttk.Frame(frm_top); row2.pack(fill='x', pady=6)
        ttk.Label(row2, text="输出根目录：").grid(row=0, column=0, sticky='w')
        ttk.Entry(row2, textvariable=self.var_out, width=70).grid(row=0, column=1, sticky='we', padx=6)
        ttk.Button(row2, text="选择文件夹", command=self.choose_out).grid(row=0, column=2)
        ttk.Label(row2, text="（留空=解压到压缩包所在目录）").grid(row=0, column=3, padx=6)

        row3 = ttk.Frame(frm_top); row3.pack(fill='x', pady=6)
        ttk.Label(row3, text="Bandizip (bz.exe)：").grid(row=0, column=0, sticky='w')
        ttk.Entry(row3, textvariable=self.var_bz, width=62).grid(row=0, column=1, sticky='we', padx=6)
        ttk.Button(row3, text="浏览", command=lambda: self.choose_exe(self.var_bz, 'bz.exe')).grid(row=0, column=2)

        ttk.Label(row3, text="7-Zip (7z.exe)：").grid(row=1, column=0, sticky='w', pady=(6,0))
        ttk.Entry(row3, textvariable=self.var_7z, width=62).grid(row=1, column=1, sticky='we', padx=6, pady=(6,0))
        ttk.Button(row3, text="浏览", command=lambda: self.choose_exe(self.var_7z, '7z.exe')).grid(row=1, column=2, pady=(6,0))
        row3.columnconfigure(1, weight=1)

        row_pol = ttk.Frame(frm_top); row_pol.pack(fill='x', pady=6)
        ttk.Label(row_pol, text="已存在文件：").grid(row=0, column=0, sticky='w')
        ttk.Radiobutton(row_pol, text="跳过（-aos）", variable=self.var_policy, value='skip').grid(row=0, column=1, padx=6)
        ttk.Radiobutton(row_pol, text="自动改名（-aou）", variable=self.var_policy, value='rename').grid(row=0, column=2, padx=6)
        ttk.Radiobutton(row_pol, text="覆盖（-aoa）", variable=self.var_policy, value='overwrite').grid(row=0, column=3, padx=6)
        ttk.Label(row_pol, text="静默阈值（秒）：").grid(row=0, column=4, padx=(16,4))
        ttk.Spinbox(row_pol, from_=10, to=600, increment=5, textvariable=self.var_quiet, width=6).grid(row=0, column=5)

        row_misc = ttk.Frame(frm_top); row_misc.pack(fill='x', pady=6)
        ttk.Checkbutton(row_misc, text="递归子目录", variable=self.var_recursive).grid(row=0, column=0, sticky='w')
        ttk.Checkbutton(row_misc, text="成功后删除源压缩包（含分卷）", variable=self.var_delete).grid(row=0, column=1, sticky='w', padx=16)
        ttk.Checkbutton(row_misc, text="解压后递归处理二次压缩包", variable=self.var_nested).grid(row=0, column=2, sticky='w')
        ttk.Checkbutton(row_misc, text="先测试再解压（更快发现损坏）", variable=self.var_pretest).grid(row=0, column=3, sticky='w', padx=16)
        ttk.Checkbutton(row_misc, text="失败自动切换解压器", variable=self.var_cross_try).grid(row=0, column=4, sticky='w')

        row_end = ttk.Frame(frm_top); row_end.pack(fill='x', pady=6)
        ttk.Label(row_end, text="完成后动作：").grid(row=0, column=0, sticky='w')
        ttk.Radiobutton(row_end, text="无", variable=self.var_end_action, value='none').grid(row=0, column=1, padx=6)
        ttk.Radiobutton(row_end, text="退出程序", variable=self.var_end_action, value='exit').grid(row=0, column=2, padx=6)
        ttk.Radiobutton(row_end, text="关机（Windows）", variable=self.var_end_action, value='shutdown').grid(row=0, column=3, padx=6)

        # 选项卡
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=12, pady=(0, 12))

        # Tab1: 扫描并解压（全量）
        tab1 = ttk.Frame(self.nb)
        self.nb.add(tab1, text="扫描并解压")
        self.t1_info = ttk.Label(tab1, text="此模式：扫描后立即按设置解压所有发现的压缩包。")
        self.t1_info.pack(anchor='w', padx=6, pady=6)
        t1_btns = ttk.Frame(tab1); t1_btns.pack(anchor='w', padx=6, pady=4)
        self.btn_start1 = ttk.Button(t1_btns, text="开始解压（全量）", command=self.on_start_full)
        self.btn_stop1 = ttk.Button(t1_btns, text="停止", command=self.on_stop, state='disabled')
        self.btn_start1.pack(side='left', padx=4)
        self.btn_stop1.pack(side='left', padx=4)

        # Tab2: 仅扫描 → 选择后解压
        tab2 = ttk.Frame(self.nb)
        self.nb.add(tab2, text="仅扫描（选择后解压）")

        t2_top = ttk.Frame(tab2); t2_top.pack(fill='x', padx=6, pady=6)
        self.btn_scan = ttk.Button(t2_top, text="扫描", command=self.on_scan_only); self.btn_scan.pack(side='left')
        self.lbl_filter = ttk.Label(t2_top, text="过滤关键词："); self.lbl_filter.pack(side='left', padx=(12,4))
        ent = ttk.Entry(t2_top, textvariable=self.var_filter, width=28); ent.pack(side='left')
        ttk.Label(t2_top, text="大小(MB)：").pack(side='left', padx=(12,4))
        ttk.Entry(t2_top, textvariable=self.var_size_min, width=6).pack(side='left')
        ttk.Label(t2_top, text="至").pack(side='left', padx=(4,4))
        ttk.Entry(t2_top, textvariable=self.var_size_max, width=6).pack(side='left')
        self.btn_apply_filter = ttk.Button(t2_top, text="应用过滤", command=self.apply_filter); self.btn_apply_filter.pack(side='left', padx=4)
        self.btn_export = ttk.Button(t2_top, text="导出列表", command=self.export_scan_list); self.btn_export.pack(side='left', padx=(10,4))
        self.lbl_workers = ttk.Label(t2_top, text="并发："); self.lbl_workers.pack(side='left', padx=(16,4))
        ttk.Spinbox(t2_top, from_=1, to=16, textvariable=self.var_workers, width=4).pack(side='left')
        self.btn_extract_sel = ttk.Button(t2_top, text="解压选中", command=self.on_extract_selected); self.btn_extract_sel.pack(side='left', padx=8)
        self.btn_select_all = ttk.Button(t2_top, text="全选", command=lambda: self._t2_select_all(True)); self.btn_select_all.pack(side='left', padx=6)
        self.btn_select_none = ttk.Button(t2_top, text="全不选", command=lambda: self._t2_select_all(False)); self.btn_select_none.pack(side='left', padx=6)
        self.lbl_t2_count = ttk.Label(t2_top, text="已列出：0")
        self.lbl_t2_count.pack(side='left', padx=12)

        # 列表
        cols = ('sel', 'fav', 'name', 'size', 'type', 'dir', 'pwd')
        self.tree = ttk.Treeview(tab2, columns=cols, show='headings', selectmode='extended', height=18)
        for c, text, w, anchor in [
            ('sel','✔',    40, 'center'),
            ('fav','★',    40, 'center'),
            ('name','文件名', 320, 'w'),
            ('size','大小',   90, 'e'),
            ('type','类型',   70, 'center'),
            ('dir', '所在目录', 400, 'w'),
            ('pwd', '推断密码', 200, 'w'),
        ]:
            self.tree.heading(c, text=text, command=lambda col=c: self.sort_by(col))
            self.tree.column(c, width=w, anchor=anchor)
        self.tree.pack(fill='both', expand=True, padx=6, pady=(0,6))
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Control-c>", self._copy_selected_cell)
        self.ctx_iid = None
        self.last_cell = {'iid': None, 'col': None}

        # 右键菜单
        self.ctx_menu = tk.Menu(self, tearoff=0)
        self.ctx_menu.add_command(label="勾选/取消勾选", command=self._ctx_toggle_check)
        self.ctx_menu.add_command(label="打开所在目录", command=self._ctx_open_dir)
        self.ctx_menu.add_command(label="删除本地文件", command=self._ctx_delete_files)
        self.ctx_menu.add_command(label="从列表移除（不删文件）", command=self._ctx_remove_items)
        self.ctx_menu.add_command(label="复制选中到目录", command=self._ctx_copy_to_dir)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="加入/取消收藏", command=self._ctx_toggle_fav)
        self.ctx_menu.add_command(label="更正密码（批量）", command=self._ctx_correct_pwd)
        self.ctx_menu.add_command(label="复制单元格", command=self._ctx_copy_cell)

        # 说明页
        self.tab_help = ttk.Frame(self.nb)
        self.nb.add(self.tab_help, text="说明 / Help")
        self.help_text = tk.Text(self.tab_help, height=20, wrap='word')
        self.help_text.pack(fill='both', expand=True, padx=8, pady=8)
        self.help_text.configure(state='disabled')

        # 统一的进度与日志（两个模式公用）
        row6 = ttk.Frame(self); row6.pack(fill='x', padx=12, pady=(0,4))
        self.progress = ttk.Progressbar(row6, mode='determinate')
        self.progress.pack(fill='x', expand=True, side='left', padx=4)
        self.lbl_stat = ttk.Label(row6, text="待处理：0 / 0")
        self.lbl_stat.pack(side='left', padx=8)

        # 新增：当前任务/阶段
        row6b = ttk.Frame(self); row6b.pack(fill='x', padx=12, pady=(0,8))
        self.lbl_now = ttk.Label(row6b, text="当前：-")
        self.lbl_now.pack(side='left', padx=(4, 18))
        self.lbl_phase = ttk.Label(row6b, text="阶段：-")
        self.lbl_phase.pack(side='left')

        frm_log = ttk.LabelFrame(self, text="日志（两个模式共用）")
        frm_log.pack(fill='both', expand=True, padx=12, pady=(0, 12))
        self.txt = tk.Text(frm_log, height=14, wrap='none')
        self.txt.pack(fill='both', expand=True, side='left')
        scroll = ttk.Scrollbar(frm_log, command=self.txt.yview)
        scroll.pack(side='right', fill='y')
        self.txt.configure(yscrollcommand=scroll.set)

        self.lang.trace_add('write', self._apply_lang)
        self._apply_lang()

    # ---------- 公共小工具 ----------

    def choose_root(self):
        d = filedialog.askdirectory(title="选择扫描根目录")
        if d:
            self.var_root.set(d)

    def choose_out(self):
        d = filedialog.askdirectory(title="选择输出根目录")
        if d:
            self.var_out.set(d)

    def choose_exe(self, var: tk.StringVar, exe_name: str):
        p = filedialog.askopenfilename(
            title=f"选择 {exe_name}", filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if p:
            var.set(p)

    def post(self, msg: str):
        self.queue.put(msg)

    def log(self, msg: str):
        self.txt.insert('end', msg + '\n'); self.txt.see('end')

    def _drain_queue(self):
        while True:
            try:
                m = self.queue.get_nowait()
            except queue.Empty:
                break
            else:
                self.log(m)
        self.after(100, self._drain_queue)

    def _update_progress(self, done: int, total: int):
        self.progress['value'] = done
        self.lbl_stat.config(text=f"已处理：{done} / {total}")

    def _set_now(self, i: int, total: int, arc: Path):
        self.lbl_now.config(text=f"当前：{i}/{total} — {arc.name}")

    def _set_phase(self, s: str):
        self.lbl_phase.config(text=f"阶段：{s}")

    def _clear_phase(self):
        self.lbl_phase.config(text=f"阶段：-")

    def _apply_lang(self, *args):
        lang = self.lang.get()
        t = LANG_TEXT.get(lang, LANG_TEXT['zh'])
        self.title(t['title'])
        try:
            self.frm_top.configure(text=t['frame_basic'])
        except Exception:
            pass
        tabs = self.nb.tabs()
        if len(tabs) >= 3:
            self.nb.tab(tabs[0], text=t['tab_scan'])
            self.nb.tab(tabs[1], text=t['tab_list'])
            self.nb.tab(tabs[2], text=t['tab_help'])
        if hasattr(self, 'btn_start1'):
            self.btn_start1.configure(text=t['start_all'])
        if hasattr(self, 'btn_stop1'):
            self.btn_stop1.configure(text=t['stop'])
        if hasattr(self, 'btn_scan'):
            self.btn_scan.configure(text=t['scan'])
        if hasattr(self, 'lbl_filter'):
            self.lbl_filter.configure(text=t['filter_kw'])
        if hasattr(self, 'btn_apply_filter'):
            self.btn_apply_filter.configure(text=t['apply_filter'])
        if hasattr(self, 'btn_export'):
            self.btn_export.configure(text=t['export'])
        if hasattr(self, 'lbl_workers'):
            self.lbl_workers.configure(text="并发：" if lang == 'zh' else "Threads:")
        if hasattr(self, 'btn_extract_sel'):
            self.btn_extract_sel.configure(text=t['extract_sel'])
        if hasattr(self, 'btn_select_all'):
            self.btn_select_all.configure(text=t['select_all'])
        if hasattr(self, 'btn_select_none'):
            self.btn_select_none.configure(text=t['select_none'])
        if hasattr(self, 't1_info'):
            self.t1_info.configure(text=t['scan_desc'])
        try:
            self.lbl_t2_count.config(text=f"{t['listed']}{len(self.scan_rows)}")
        except Exception:
            pass
        if hasattr(self, 'help_text'):
            self.help_text.configure(state='normal')
            self.help_text.delete('1.0', 'end')
            self.help_text.insert('end', t['help_body'])
            self.help_text.configure(state='disabled')

    def _init_progress(self, total: int):
        self.progress['maximum'] = max(total, 1)
        self.lbl_stat.config(text=f"已处理：0 / {total}")

    def _finish_run(self, stopped: bool):
        self.btn_start1.configure(state='normal')
        self.btn_stop1.configure(state='disabled')
        if stopped:
            self.post("已停止，未执行完成后动作。")
            return
        self.after(800, self._do_end_action)

    # ---------- Tab1: 全量扫描并解压 ----------

    def on_start_full(self):
        root = Path(self.var_root.get().strip('" '))
        if not root.is_dir():
            messagebox.showerror("错误", "请先选择有效的扫描根目录")
            return
        self.stop_flag.clear()
        self.btn_start1.configure(state='disabled')
        self.btn_stop1.configure(state='normal')
        self.txt.delete('1.0', 'end')
        self.progress['value'] = 0
        self.lbl_stat.config(text="准备中...")
        self.worker = threading.Thread(target=self._work_full, args=(root,), daemon=True)
        self.worker.start()

    def on_stop(self):
        self.stop_flag.set()
        self.post("请求停止，正在结束当前任务...")
        self.btn_stop1.configure(state='disabled')

    def _work_full(self, root: Path):
        try:
            archives = gather_archives(root, self.var_recursive.get())
            total = len(archives); done = 0
            self.post(f"发现压缩包：{total} 个")
            self.after(0, self._init_progress, total)

            for idx, arc in enumerate(archives, 1):
                if self.stop_flag.is_set():
                    break
                # 进度标签 + 开始日志
                self.after(0, self._set_now, idx, total, arc)
                self.after(0, self._set_phase, "准备")
                self.post(f"== 开始：[{idx}/{total}] {arc}")
                self._handle_one_archive(arc, root)
                done += 1; self.after(0, self._update_progress, done, total)
                self.after(0, self._set_phase, "完成")

        finally:
            self.post("任务结束。")
        self.after(0, self._finish_run, self.stop_flag.is_set())

    # 处理单个压缩包（解压）
    def _handle_one_archive(self, arc: Path, root_for_rel: Path):
        sig = sniff_signature(arc)
        if sig in ('html', 'xml', 'pdf'):
            self.post(f"⚠ 不是压缩包（检测到 {sig.upper()} 头），可能下载的是网页/占位文件：{arc}")
            return

        password = infer_password(arc)

        out_base = Path(self.var_out.get().strip('" ')) if self.var_out.get().strip() else arc.parent
        if self.var_out.get().strip():
            try:
                rel = arc.parent.relative_to(root_for_rel) if root_for_rel in arc.parents else Path('')
                out_dir = out_base / rel / (arc.stem)
            except Exception:
                out_dir = out_base / (arc.stem)
        else:
            out_dir = arc.parent / (arc.stem)
        out_dir.mkdir(parents=True, exist_ok=True)

        bz = self.var_bz.get().strip('" ')
        sz = self.var_7z.get().strip('" ')
        policy = self.var_policy.get()
        quiet = max(10, int(self.var_quiet.get() or 30))

        first = None; second = None
        if bz and Path(bz).is_file():
            first = ('bandizip', bz)
        if sz and Path(sz).is_file():
            if first is None:
                first = ('7zip', sz)
            else:
                second = ('7zip', sz)

        if first is None and second is None:
            bz_auto = find_on_path(['bz.exe'])
            sz_auto = find_on_path(['7z.exe'])
            if bz_auto:
                first = ('bandizip', bz_auto)
                self.post(f"[提示] 已自动找到 Bandizip：{bz_auto}")
            if sz_auto and (first is None):
                first = ('7zip', sz_auto)
                self.post(f"[提示] 已自动找到 7-Zip：{sz_auto}")
            elif sz_auto:
                second = ('7zip', sz_auto)

        if first is None:
            self.post(f"!! 未找到解压程序，跳过：{arc}")
            return

        # 测试
        if self.var_pretest.get():
            self.after(0, self._set_phase, "测试")
            tester = self._test_archive(first, arc, password, quiet)
            if tester is False and self.var_cross_try.get() and second:
                self.post("  ↺ 测试失败，切换另一个解压器再测...")
                self.after(0, self._set_phase, "测试（切换）")
                if self._test_archive(second, arc, password, quiet) is False:
                    self.post("✖ 归类为不可用/损坏或分卷缺失，已跳过（可尝试重新下载/补齐分卷/修复）")
                    return

        # 解压
        self.after(0, self._set_phase, "解压")
        ok = self._extract_with(first, arc, out_dir, password, policy, quiet)
        if not ok and self.var_cross_try.get() and second:
            self.post("  ↺ 失败，切换另一个解压器重试...")
            self.after(0, self._set_phase, "解压（切换）")
            ok = self._extract_with(second, arc, out_dir, password, policy, quiet)

        if ok:
            if self.var_nested.get():
                nested = self._extract_nested(out_dir, password, policy,
                                              first[0] if ok else '7zip',
                                              first[1] if first and first[0]=='bandizip' else '',
                                              second[1] if second and second[0]=='7zip' else (first[1] if first and first[0]=='7zip' else ''))
                if nested:
                    self.post(f"  ✔ 二次解压完成（{nested} 个）")
            if self.var_delete.get():
                removed = 0
                for p in get_all_multipart_siblings(arc):
                    try: p.unlink(missing_ok=True); removed += 1
                    except Exception: pass
                self.post(f"  🗑 已删除源压缩包 {removed} 个")
        else:
            self.post(f"✖ 解压失败。建议：检查密码/文件完整性/分卷是否齐全/更换解压器版本")

    # ---------- Tab2: 仅扫描 & 列表/过滤/排序/并行 ----------

    def on_scan_only(self):
        root = Path(self.var_root.get().strip('" '))
        if not root.is_dir():
            messagebox.showerror("错误", "请先选择有效的扫描根目录")
            return
        self.tree.delete(*self.tree.get_children())
        self.scan_rows.clear(); self.bytes_map.clear(); self.checked_map.clear(); self.favorite_map.clear()
        paths = gather_archives(root, self.var_recursive.get())
        for p in paths:
            sig = sniff_signature(p)
            szb = file_size(p)
            row = {
                'path': p, 'name': p.name, 'sizeb': szb, 'sizes': human(szb),
                'type': sig, 'dir': str(p.parent), 'pwd': infer_password(p) or "",
                'checked': False, 'fav': False
            }
            self.scan_rows.append(row); self.bytes_map[str(p)] = szb
        self._reload_tree(self.scan_rows)
        self.lbl_t2_count.config(text=f"已列出：{len(self.scan_rows)}")

    def _reload_tree(self, rows: List[Dict]):
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            iid = str(r['path'])
            checked = bool(r.get('checked'))
            fav = bool(r.get('fav'))
            self.tree.insert('', 'end', iid=iid, values=(
                '✓' if checked else '',
                '★' if fav else '',
                r['name'], r['sizes'], r['type'], r['dir'], r['pwd']
            ))
            self.checked_map[iid] = checked
            self.favorite_map[iid] = fav

    def _set_checked(self, iid: str, flag: bool):
        self.checked_map[iid] = flag
        self._update_scan_row_state(iid, 'checked', flag)
        vals = list(self.tree.item(iid, 'values'))
        vals[0] = '✓' if flag else ''
        self.tree.item(iid, values=vals)

    def _set_favorite(self, iid: str, flag: bool):
        self.favorite_map[iid] = flag
        self._update_scan_row_state(iid, 'fav', flag)
        vals = list(self.tree.item(iid, 'values'))
        vals[1] = '★' if flag else ''
        self.tree.item(iid, values=vals)

    def _on_tree_click(self, event):
        row = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if row and col:
            self.last_cell = {'iid': row, 'col': col}
        if not row or col not in ('#1', '#2'):
            return
        if col == '#1':
            self._set_checked(row, not self.checked_map.get(row, False))
        elif col == '#2':
            self._set_favorite(row, not self.favorite_map.get(row, False))
        return "break"

    def _ctx_selected_iids(self) -> List[str]:
        sel = list(self.tree.selection())
        if not sel and self.ctx_iid:
            sel = [self.ctx_iid]
        return sel

    def _on_tree_right_click(self, event):
        row = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if row and col:
            self.last_cell = {'iid': row, 'col': col}
        if row:
            # 如果未选中则追加选中，已选则保留原有多选
            if row not in self.tree.selection():
                self.tree.selection_add(row)
            self.ctx_iid = row
        else:
            self.ctx_iid = None
        try:
            self.ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.ctx_menu.grab_release()

    def _on_tree_double_click(self, event):
        row = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row or not col:
            return
        self.last_cell = {'iid': row, 'col': col}
        # 仅允许在“推断密码”列编辑
        if col != '#7':
            return
        bbox = self.tree.bbox(row, col)
        if not bbox:
            return
        x, y, w, h = bbox
        vals = list(self.tree.item(row, 'values'))
        old = vals[6]
        entry = ttk.Entry(self.tree)
        entry.insert(0, old)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()

        def save_edit(event=None):
            new_val = entry.get()
            entry.destroy()
            vals[6] = new_val
            self.tree.item(row, values=vals)
            self._update_scan_row_state(row, 'pwd', new_val)

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    def _copy_selected_cell(self, event=None):
        cell = self.last_cell
        if not cell or not cell.get('iid'):
            return
        vals = self.tree.item(cell['iid'], 'values')
        try:
            idx = int(cell['col'].lstrip('#')) - 1
        except Exception:
            return
        if idx < 0 or idx >= len(vals):
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(str(vals[idx]))
        except Exception:
            pass

    def _ctx_copy_cell(self):
        self._copy_selected_cell()

    def _ctx_toggle_check(self):
        iids = self._ctx_selected_iids()
        if not iids:
            return
        target_flag = not all(self.checked_map.get(i, False) for i in iids)
        for iid in iids:
            self._set_checked(iid, target_flag)

    def _ctx_toggle_fav(self):
        iids = self._ctx_selected_iids()
        if not iids:
            return
        target_flag = not all(self.favorite_map.get(i, False) for i in iids)
        for iid in iids:
            self._set_favorite(iid, target_flag)

    def _ctx_open_dir(self):
        iids = self._ctx_selected_iids()
        if not iids:
            return
        dirs = set()
        for iid in iids:
            vals = self.tree.item(iid, 'values')
            if len(vals) >= 6:
                dirs.add(vals[5])
        for d in dirs:
            try:
                if os.name == 'nt':
                    os.startfile(d)
                else:
                    subprocess.Popen(['xdg-open', d])
            except Exception as e:
                self.post(f"!! 打开目录失败：{d} ({e})")

    def _remove_row(self, iid: str):
        self.tree.delete(iid)
        self.checked_map.pop(iid, None)
        self.favorite_map.pop(iid, None)
        self.bytes_map.pop(iid, None)
        self.scan_rows = [r for r in self.scan_rows if str(r.get('path')) != iid]

    def _ctx_delete_files(self):
        iids = self._ctx_selected_iids()
        if not iids:
            return
        if not messagebox.askyesno("确认", f"删除本地文件（共 {len(iids)} 个）并从列表移除？此操作不可恢复。"):
            return
        removed = 0
        for iid in iids:
            try:
                Path(iid).unlink(missing_ok=True)
                removed += 1
            except Exception as e:
                self.post(f"!! 删除失败：{iid} ({e})")
            self._remove_row(iid)
        self.lbl_t2_count.config(text=f"已列出：{len(self.scan_rows)}")
        self.post(f"已删除并移除 {removed} 个文件")

    def _ctx_remove_items(self):
        iids = self._ctx_selected_iids()
        if not iids:
            return
        for iid in iids:
            self._remove_row(iid)
        self.lbl_t2_count.config(text=f"已列出：{len(self.scan_rows)}")
        self.post(f"已从列表移除 {len(iids)} 条记录（未删除文件）")

    def _ctx_copy_to_dir(self):
        iids = self._ctx_selected_iids()
        if not iids:
            return
        target = filedialog.askdirectory(title="选择目标目录（复制选中压缩包）")
        if not target:
            return
        target_path = Path(target)
        copied = 0
        for iid in iids:
            src = Path(iid)
            if not src.exists():
                self.post(f"? 文件不存在，跳过：{src}")
                continue
            dst = target_path / src.name
            try:
                if dst.exists():
                    dst = target_path / f"{src.stem}_copy{src.suffix}"
                shutil.copy2(src, dst)
                copied += 1
            except Exception as e:
                self.post(f"!! 复制失败：{src} -> {dst} ({e})")
        self.post(f"复制完成：{copied}/{len(iids)} 个已放到 {target_path}")

    def _ctx_correct_pwd(self):
        iids = self._ctx_selected_iids()
        if not iids:
            return
        new_pwd = simpledialog.askstring("更正密码", "输入新的解压密码（留空则清除）：", parent=self)
        if new_pwd is None:
            return
        for iid in iids:
            self._update_scan_row_state(iid, 'pwd', new_pwd)
            vals = list(self.tree.item(iid, 'values'))
            vals[6] = new_pwd
            self.tree.item(iid, values=vals)
        self.post(f"已更新 {len(iids)} 条记录的密码")

    def apply_filter(self):
        kw = self.var_filter.get().strip().lower()
        min_mb = self.var_size_min.get().strip()
        max_mb = self.var_size_max.get().strip()
        min_b = None; max_b = None
        try:
            if min_mb:
                min_b = float(min_mb) * 1024 * 1024
        except ValueError:
            messagebox.showerror("错误", "最小大小请输入数字（MB）")
            return
        try:
            if max_mb:
                max_b = float(max_mb) * 1024 * 1024
        except ValueError:
            messagebox.showerror("错误", "最大大小请输入数字（MB）")
            return

        filt = []
        for r in self.scan_rows:
            blob = f"{r['name']} {r['dir']} {r['pwd']}".lower()
            if kw and kw not in blob:
                continue
            sz = r.get('sizeb', 0)
            if min_b is not None and sz < min_b:
                continue
            if max_b is not None and sz > max_b:
                continue
            filt.append(r)
        self._reload_tree(filt)
        tag = "（过滤后）" if kw or min_b is not None or max_b is not None else ""
        self.lbl_t2_count.config(text=f"已列出：{len(filt)} {tag}")

    def export_scan_list(self):
        """导出当前列表或选中项为 Excel（尊重过滤结果）"""
        items = list(self.tree.selection()) or list(self.tree.get_children())
        if not items:
            messagebox.showinfo("提示", "列表为空，无法导出。请先扫描或应用过滤后再导出。")
            return
        path = filedialog.asksaveasfilename(
            title="导出预览列表为 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.append(["勾选", "收藏", "文件名", "大小", "类型", "所在目录", "推断密码"])
            for iid in items:
                vals = self.tree.item(iid, 'values')
                ws.append([vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], vals[6]])
            wb.save(path)
            messagebox.showinfo("完成", f"已导出 {len(items)} 条记录到：\n{path}")
        except ImportError:
            messagebox.showerror("错误", "缺少 openpyxl，请先安装：pip install openpyxl")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{e}")

    def sort_by(self, col: str):
        items = list(self.tree.get_children(''))
        def keyfunc(iid):
            vals = self.tree.item(iid, 'values')
            if col == 'sel':
                return (self.checked_map.get(iid, False),)
            if col == 'fav':
                return (self.favorite_map.get(iid, False),)
            if col == 'name':
                return (vals[2].lower(),)
            elif col == 'size':
                return (self.bytes_map.get(iid, 0),)
            elif col == 'type':
                return (vals[4].lower(),)
            elif col == 'dir':
                return (vals[5].lower(),)
            elif col == 'pwd':
                return (vals[6].lower(),)
            return (vals[2].lower(),)
        if hasattr(self, 'sort_state') and self.sort_state.get('col') == col:
            self.sort_state['desc'] = not self.sort_state['desc']
        else:
            self.sort_state = {'col': col, 'desc': False}
        items.sort(key=keyfunc, reverse=self.sort_state['desc'])
        for idx, iid in enumerate(items):
            self.tree.move(iid, '', idx)

    def _t2_select_all(self, flag: bool):
        if flag:
            self.tree.selection_set(self.tree.get_children())
        else:
            self.tree.selection_remove(self.tree.get_children())
        for iid in self.tree.get_children():
            self.checked_map[iid] = flag
            self._update_scan_row_state(iid, 'checked', flag)
            vals = list(self.tree.item(iid, 'values'))
            vals[0] = '✓' if flag else ''
            self.tree.item(iid, values=vals)

    def _update_scan_row_state(self, iid: str, key: str, val):
        for r in self.scan_rows:
            if str(r.get('path')) == iid:
                r[key] = val
                break

    def on_extract_selected(self):
        checked = [iid for iid, v in self.checked_map.items() if v]
        sel = checked or list(self.tree.selection())
        if not sel:
            messagebox.showinfo("提示", "请先勾选或选择要解压的项（支持多选）。")
            return
        workers = max(1, min(int(self.var_workers.get() or 1), 16))
        total = len(sel)
        self.stop_flag.clear()
        self.txt.delete('1.0', 'end')
        self.progress['value'] = 0
        self.progress['maximum'] = total
        self.lbl_stat.config(text=f"准备中...（并发 {workers}）")

        done_lock = threading.Lock()
        done_cnt = {'n': 0}

        def task(iid: str):
            if self.stop_flag.is_set():
                return
            arc = Path(iid)
            root = Path(self.var_root.get().strip('" '))
            if not arc.exists():
                self.post(f"⚠ 找不到文件，跳过：{arc}")
            else:
                root_for_rel = root if root.is_dir() else arc.parent
                self._handle_one_archive(arc, root_for_rel)
            with done_lock:
                done_cnt['n'] += 1
                n = done_cnt['n']
            self.after(0, self._update_progress, n, total)

        def worker_selected():
            try:
                with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="extract") as ex:
                    futures = [ex.submit(task, iid) for iid in sel]
                    for _ in as_completed(futures):
                        if self.stop_flag.is_set():
                            for f in futures:
                                f.cancel()
                            break
            finally:
                self.post("所选项处理完成。")
                self.after(0, self._finish_run, self.stop_flag.is_set())

        threading.Thread(target=worker_selected, daemon=True).start()

    # ---------- 解压子流程 ----------

    def _test_archive(self, tool_pair, arc: Path, pwd: Optional[str], quiet: int) -> Optional[bool]:
        name, exe = tool_pair
        if name == 'bandizip':
            cmd = bandizip_test_cmd(exe, arc, pwd)
        else:
            cmd = sevenzip_test_cmd(exe, arc, pwd)
        self.post(f"→ 测试：{arc}  使用：{name}")
        rc = run_cmd(cmd, self.post, self.stop_flag, monitor_dir=None, quiet_limit=quiet, phase_name="测试")
        if rc == 0:
            self.post("  ✔ 测试通过")
            return True
        if rc in (-1, 9001, 9002):
            return None
        return False

    def _extract_with(self, tool_pair, arc: Path, out_dir: Path, pwd: Optional[str], policy: str, quiet: int) -> bool:
        name, exe = tool_pair
        if name == 'bandizip':
            cmd = bandizip_cmd(exe, arc, out_dir, pwd, policy)
        else:
            cmd = sevenzip_cmd(exe, arc, out_dir, pwd, policy)
        self.post(f"→ 解压：{arc}  使用：{name}  输出：{out_dir}  策略：{policy}")
        rc = run_cmd(cmd, self.post, self.stop_flag, monitor_dir=out_dir, quiet_limit=quiet, phase_name="解压")
        return rc == 0

    def _extract_nested(self, root: Path, password: str, policy: str, exe_name: str, bz: str, sz: str) -> int:
        count = 0
        for dirpath, _, files in os.walk(root):
            if self.stop_flag.is_set():
                break
            for f in files:
                if self.stop_flag.is_set():
                    break
                arc = Path(dirpath) / f
                arc = normalize_extension(arc)
                low = arc.name.lower()
                if any([low.endswith('.zip'), low.endswith('.7z'), low.endswith('.rar'),
                        low.endswith('.001'), low.endswith('.z01'),
                        re.search(r'\.part\d+\.rar$', low) is not None]):
                    is_multi, is_first = is_multipart_first(arc)
                    if is_multi and not is_first:
                        continue
                    out_dir = arc.parent / (arc.stem)
                    out_dir.mkdir(parents=True, exist_ok=True)
                    if exe_name == 'bandizip' and bz and Path(bz).is_file():
                        cmd = bandizip_cmd(bz, arc, out_dir, password, policy)
                    elif sz and Path(sz).is_file():
                        cmd = sevenzip_cmd(sz, arc, out_dir, password, policy)
                    else:
                        continue
                    rc = run_cmd(cmd, self.post, self.stop_flag, monitor_dir=out_dir, quiet_limit=max(10, int(self.var_quiet.get() or 30)), phase_name="二次解压")
                    if rc == 0:
                        count += 1
                        if self.var_delete.get():
                            for p in get_all_multipart_siblings(arc):
                                try: p.unlink(missing_ok=True)
                                except: pass
        return count

    # ---------- 完成后动作 ----------

    def _do_end_action(self):
        if self.stop_flag.is_set():
            self.post("已停止，未执行完成后动作。")
            return
        action = self.var_end_action.get()
        if action == 'exit':
            self.post("已选择：完成后退出程序。")
            self.after(200, self.destroy)
        elif action == 'shutdown':
            self.post("已选择：完成后关机。将调用系统关机命令。")
            try:
                if os.name == 'nt':  # Windows
                    subprocess.Popen(['shutdown', '/s', '/t', '0'])
                elif sys.platform == 'darwin':  # macOS
                    subprocess.Popen(['osascript', '-e', 'tell application \"System Events\" to shut down'])
                else:  # Linux/Unix
                    subprocess.Popen(['shutdown', '-h', 'now'])
            except Exception as e:
                self.post(f"!! 执行关机命令失败：{e}")
        else:
            self.post("已选择：完成后不执行额外动作。")

if __name__ == '__main__':
    app = AutoExtractorApp()
    app.mainloop()
