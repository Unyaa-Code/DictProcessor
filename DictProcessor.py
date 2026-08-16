import os
import re
import ctypes
import io
import itertools
import contextlib
import webbrowser
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

# wintypes 仅 Windows 存在，跨平台适配时条件导入（Linux/macOS 上无此模块）
if sys.platform == 'win32':
    from ctypes import wintypes
else:
    wintypes = None

# 跨平台 CJK 字体：Windows 用雅黑、macOS 用苹方、Linux 用 Noto Sans CJK
# 各平台默认字体不存在时 Tk 会自动回退，不会崩溃
if sys.platform == 'win32':
    CJK_FONT = 'Microsoft YaHei UI'
elif sys.platform == 'darwin':
    CJK_FONT = 'PingFang SC'
else:
    CJK_FONT = 'Noto Sans CJK SC'

# 应用版本号（关于页、GitHub Actions 打包均引用此值）
APP_VERSION = '1.4.0'

# Windows 高分屏 DPI 感知，避免 GUI 模糊
# ctypes.windll 仅 Windows 存在，try/except 捕获后其他平台静默跳过
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# 可选编码列表（用于 GUI 手动选择；'auto' 走 detect_encoding 自动检测）
ENCODING_CHOICES = [
    ('自动检测', 'auto'),
    ('UTF-8', 'utf-8'),
    ('UTF-8 with BOM', 'utf-8-sig'),
    ('GBK', 'gbk'),
    ('GB18030', 'gb18030'),
    ('Big5', 'big5'),
    ('UTF-16 LE', 'utf-16-le'),
    ('UTF-16 BE', 'utf-16-be'),
    ('UTF-32 LE', 'utf-32-le'),
    ('UTF-32 BE', 'utf-32-be'),
    ('ISO-8859-1', 'iso-8859-1'),
]

# 词组编码功能的忽略设置
# - 忽略标点符号：独立复选框（默认勾选）；「忽略表」输入框默认同时含标点 + 编辑符号
# - 忽略表：单行输入框，直接填需剔除的字符；取消「忽略标点符号」时其中标点部分不生效
IGNORE_PUNCT = '，。、；：！？…—～·（）《》【】「」『』〈〉〔〕“”‘’'
DEFAULT_IGNORE_EXTRA = '·#@%&*+=_~^$|\\/<>{}'
# 「忽略表」输入框默认内容（标点 + 编辑符号，保证默认不被「变少」）
DEFAULT_IGNORE_BOX = IGNORE_PUNCT + DEFAULT_IGNORE_EXTRA


# 词组编码：内置规则预设（一键载入到「生成规则」框）
# 每行格式：范围 = 规则；范围如 2 / 3 / 4,99，规则用 [字序][码序] 拼接
PHRASE_PRESETS = {
    '五笔规则': (
        '2 = [0][:2] + [1][:2]\n'
        '3 = [0][0] + [1][0] + [-1][:2]\n'
        '4,99 = [0:3][0] + [-1][0]'
    ),
    '两笔规则': (
        '2 = [0][:2] + [1][:2]\n'
        '3 = [0][:2] + [1][0] + [2][0]\n'
        '4,99 = [0:3][0] + [-1][0]'
    ),
    '拼音规则': (
        '2,99 = [:][:]'
    ),
    '速成规则': (
        '2,99 = [:][0,-1]'
    ),
}


def resolve_encoding(path, choice):
    """根据用户选择解析最终使用的编码名

    choice='auto' 时调用 detect_encoding 自动识别，否则直接返回 choice。
    """
    if choice == 'auto' or not choice:
        return detect_encoding(path, sample_size=64 * 1024)
    return choice


def open_folder_and_select(path):
    """在系统默认文件管理器中打开并选中指定文件（跨平台）

    - Windows: Shell API SHOpenFolderAndSelectItems（不依赖 explorer.exe，支持第三方资源管理器）
    - macOS:   open -R（在 Finder 中打开并选中文件）
    - Linux:   xdg-open 打开父目录（Linux 桌面环境无统一的"选中文件"标准，仅打开所在目录）
    返回 True 表示成功，False 表示失败（调用方可回退到其他方式）。
    """
    if not path or not os.path.exists(path):
        return False

    if sys.platform == 'win32':
        return _open_folder_and_select_win(path)
    elif sys.platform == 'darwin':
        return _open_folder_and_select_mac(path)
    else:
        return _open_folder_and_select_linux(path)


def _open_folder_and_select_win(path):
    """Windows: 使用 Shell API SHOpenFolderAndSelectItems 打开并选中文件"""
    # 初始化 COM（STA 模式，与 GUI 线程一致）
    try:
        ctypes.windll.ole32.CoInitializeEx(None, 2)  # COINIT_APARTMENTTHREADED
    except Exception:
        # 已初始化过，忽略 RPC_E_CHANGED_CONTEXT
        pass

    shell32 = ctypes.windll.shell32

    # 配置函数签名
    # PCIDLIST_ABSOLUTE ILCreateFromPathW(LPCWSTR pszPath)
    shell32.ILCreateFromPathW.restype = ctypes.c_void_p
    shell32.ILCreateFromPathW.argtypes = [wintypes.LPCWSTR]

    # void ILFree(PCIDLIST pidl)
    shell32.ILFree.restype = None
    shell32.ILFree.argtypes = [ctypes.c_void_p]

    # HRESULT SHOpenFolderAndSelectItems(
    #   PCIDLIST_ABSOLUTE pidlFolder, UINT cidl,
    #   PCUITEMID_CHILD_ARRAY apidl, DWORD dwFlags)
    shell32.SHOpenFolderAndSelectItems.restype = ctypes.c_long  # HRESULT
    shell32.SHOpenFolderAndSelectItems.argtypes = [
        ctypes.c_void_p,  # pidlFolder
        ctypes.c_uint,    # cidl
        ctypes.c_void_p,  # apidl
        wintypes.DWORD,   # dwFlags
    ]

    pidl = shell32.ILCreateFromPathW(os.path.abspath(path))
    if not pidl:
        return False

    try:
        # 传入文件的 PIDL + cidl=0：Windows 自动打开父文件夹并选中该文件
        hr = shell32.SHOpenFolderAndSelectItems(pidl, 0, None, 0)
        return hr == 0
    finally:
        shell32.ILFree(pidl)


def _open_folder_and_select_mac(path):
    """macOS: 使用 open -R 在 Finder 中打开并选中文件"""
    try:
        subprocess.Popen(['open', '-R', path])
        return True
    except Exception:
        return False


def _open_folder_and_select_linux(path):
    """Linux: 使用 xdg-open 打开父目录（Linux 无统一的"选中文件"标准）"""
    try:
        parent = os.path.dirname(os.path.abspath(path))
        subprocess.Popen(['xdg-open', parent])
        return True
    except Exception:
        return False


# 预览框滚动条样式是否已初始化（避免重复 element_create 报错）
_preview_sb_style_inited = False


def _init_preview_scrollbar_style():
    """为 ttk.Scrollbar 创建基于 clam 主题的样式，确保滚动条颜色配置生效

    问题背景：Windows 默认主题（vista/xpnative）下 ttk.Scrollbar 与 tk.Scrollbar
    均使用系统原生渲染，troughcolor/background 等颜色参数被忽略，滑块(thumb)颜色
    由系统主题决定。在某些 Windows 视觉样式下 thumb 与滑槽对比度不足，导致看不到
    滑块位置（只能看到整个滚动条槽，鼠标拖动仍能滚动）。

    解决方案：从 clam 主题（Tk 内置纯绘制主题，不依赖系统）导入 scrollbar 元素，
    创建自定义样式 Preview.Vertical.TScrollbar，颜色配置完全生效，thumb 始终可见。
    模块级只初始化一次，所有 _Preview 实例共享。
    """
    global _preview_sb_style_inited
    if _preview_sb_style_inited:
        return
    _preview_sb_style_inited = True
    style = ttk.Style()
    # 从 clam 主题导入垂直滚动条的 trough、thumb、grip 元素
    # grip 是 thumb 中央的小装饰方块，导入后配置同色让它隐藏，避免"中间浅灰色小块"
    for name, src in [
        ('Preview.Scrollbar.trough', 'Vertical.Scrollbar.trough'),
        ('Preview.Scrollbar.thumb', 'Vertical.Scrollbar.thumb'),
        ('Preview.Scrollbar.grip', 'Vertical.Scrollbar.grip'),
    ]:
        try:
            style.element_create(name, 'from', 'clam', src)
        except Exception:
            pass  # 元素已存在
    try:
        # 布局：trough 外层 + thumb 内层（含 grip 子元素，同色后不可见）
        style.layout('Preview.Vertical.TScrollbar', [
            ('Preview.Scrollbar.trough', {
                'sticky': 'ns',
                'children': [
                    ('Preview.Scrollbar.thumb', {
                        'sticky': 'nswe',
                        'children': [
                            ('Preview.Scrollbar.grip', {'sticky': ''})
                        ]
                    })
                ]
            })
        ])
        # 颜色统一：lightcolor/darkcolor/bordercolor 都设为 thumb 色，消除 3D 高光边框；
        # foreground（grip 用色）也设为 thumb 色，隐藏中央装饰方块。
        # thumb 用中浅灰 #c0c0c0，trough 用更浅的 #ececec，对比柔和不过深。
        style.configure('Preview.Vertical.TScrollbar',
                        troughcolor='#ececec',
                        background='#c0c0c0',
                        foreground='#c0c0c0',     # grip 颜色 = thumb 颜色，隐藏 grip
                        bordercolor='#c0c0c0',    # 边框 = thumb 色，无对比
                        lightcolor='#c0c0c0',     # 高光 = thumb 色，消除顶部浅色条
                        darkcolor='#c0c0c0',      # 阴影 = thumb 色，消除底部深色条
                        arrowsize=14)
        style.map('Preview.Vertical.TScrollbar',
                  background=[('active', '#a0a0a0'), ('pressed', '#909090')],
                  foreground=[('active', '#a0a0a0'), ('pressed', '#909090')])
    except Exception:
        pass


def read_preview(path, n=100, enc_choice='auto'):
    """读取文件前 n 行内容用于预览，返回 (预览文本, 总行数)

    大文件优化：行数统计改为按 1MB 块读取字节流并计数 b'\\n'，避免逐行解码。
    enc_choice: 编码选择，'auto' 时自动检测；其他值（如 'utf-8'、'gbk'）直接使用。
    """
    try:
        # 1. 读取前 n 行（容错：根据用户选择解析编码，非法字节替换为 U+FFFD）
        enc = resolve_encoding(path, enc_choice)
        shown = []
        with open(path, 'r', encoding=enc, errors='replace', buffering=1 << 20) as f:
            for i, line in enumerate(f):
                if i >= n:
                    break
                shown.append(line.rstrip('\n'))

        # 2. 快速统计行数（按 1MB 块读取字节流，统计换行符数量）
        total = 0
        with open(path, 'rb') as f:
            while True:
                chunk = f.read(1 << 20)
                if not chunk:
                    break
                total += chunk.count(b'\n')

        if total == 0 and not shown:
            return '（空文件）', 0
        result = '\n'.join(shown)
        if total > n:
            result += f'\n…（省略 {total - n} 行）'
        return result, total
    except Exception as e:
        return f'读取失败: {e}', 0


def build_key_set(source, ignore_col, enc_choice='auto'):
    """加载参考文件，返回用于比对的 key 集合（ignore_col 控制是否忽略第二列）

    使用 resolve_encoding 解析编码（自动或手动），支持非 UTF-8 文件中的
    Unicode 新字符（如 CJK Ext-B 的 "𰻝"）。source 为 list 时直接来自内存。
    """
    keys = set()
    try:
        with _iter_source(source, enc_choice) as f:
            for line in f:
                stripped = line.strip()
                if not stripped or stripped.startswith('#'):
                    continue
                key = stripped.split('\t', 1)[0] if ignore_col else stripped
                keys.add(key)
    except Exception as e:
        messagebox.showerror('错误', f'读取参考文件出错: {e}')
    return keys


def get_unique_output_path(out_path):
    """若输出文件已存在，仿 Windows 自动追加 (1)、(2)... 序号，避免覆盖"""
    if not os.path.exists(out_path):
        return out_path
    base, ext = os.path.splitext(out_path)
    i = 1
    while True:
        candidate = f"{base}({i}){ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


def detect_encoding(path, sample_size=None):
    """启发式检测文件编码

    sample_size: 若指定，仅读取文件头部指定字节数进行检测（适合大文件流式处理）；
                 未指定时读取整个文件。
    注意：UTF-16/UTF-32 仅通过 BOM 识别，候选列表中不再包含 'utf-16'，
          因为 Python 的 'utf-16' 编解码器要求流必须以 BOM 开头，
          采样检测时容易误判并在 open() 时抛出 'Stream does not start with BOM'。
    """
    try:
        with open(path, 'rb') as f:
            data = f.read(sample_size) if sample_size is not None else f.read()
    except Exception:
        return 'utf-8'

    # BOM 识别（覆盖 UTF-16/UTF-32 各端序，以及 UTF-8 BOM）
    if data.startswith(b'\x00\x00\xfe\xff'): return 'utf-32-be'
    if data.startswith(b'\xff\xfe\x00\x00'): return 'utf-32-le'
    if data.startswith(b'\xfe\xff'): return 'utf-16-be'
    if data.startswith(b'\xff\xfe'): return 'utf-16-le'
    if data.startswith(b'\xef\xbb\xbf'): return 'utf-8-sig'

    # 无 BOM 时按候选依次尝试严格解码
    # 注意：不包含 'utf-16'，避免无 BOM 的数据被误判后在 open 时失败
    candidates = ['utf-8', 'gb18030', 'gbk', 'big5', 'iso-8859-1']
    for enc in candidates:
        try:
            data.decode(enc, errors='strict')
            return enc
        except Exception:
            continue
    return 'iso-8859-1'


def _manual_lines(content):
    """将手动输入的预览文本转成『每行带换行符』的列表，供内存模式直接迭代。

    手动输入上限 1 万行、体积很小，直接在内存中处理即可，无需落盘临时文件。
    """
    return [ln + '\n' for ln in content.split('\n')]


@contextlib.contextmanager
def _iter_source(source, enc_choice='auto'):
    """统一迭代输入源的上下文管理器。

    source 为 list/tuple（手动输入，已在内存中）时直接迭代；
    为文件路径字符串时按编码流式打开文件（支持千万行级数据）。
    """
    if isinstance(source, (list, tuple)):
        yield iter(source)
    else:
        enc = resolve_encoding(source, enc_choice)
        with open(source, 'r', encoding=enc, errors='replace', buffering=1 << 20) as f:
            yield f


def process(target_path, ref_path, ignore_target, ignore_ref, operation, output_path,
            progress_callback=None, target_enc='auto', ref_enc='auto', out_stream=None):
    """执行差集/交集处理（流式，支持千万行级数据）

    使用 resolve_encoding 解析编码（自动或手动），支持非 UTF-8 文件中的 Unicode 新字符
    （如 CJK Ext-B 的 "𰻝"）。输出统一使用 UTF-8。
    out_stream 不为 None 时直接写入该内存流（手动输入模式），否则写入 output_path 文件。
    """
    ref_set = build_key_set(ref_path, ignore_ref, enc_choice=ref_enc)
    if not ref_set and operation == 'intersection':
        messagebox.showwarning('提示', '参考文件未读取到有效内容。')
        return 0

    count = 0
    processed = 0
    PROGRESS_INTERVAL = 100_000  # 每 10 万行更新一次进度
    f_out = out_stream if out_stream is not None else open(
        output_path, 'w', encoding='utf-8', buffering=1 << 20)
    own = out_stream is None
    try:
        with _iter_source(target_path, target_enc) as f_in:
            for line in f_in:
                stripped = line.strip()
                # 保留注释与空行
                if not stripped or stripped.startswith('#'):
                    f_out.write(line)
                    continue
                processed += 1
                key = stripped.split('\t', 1)[0] if ignore_target else stripped
                if operation == 'difference':
                    if key not in ref_set:
                        f_out.write(line)
                        count += 1
                else:
                    if key in ref_set:
                        f_out.write(line)
                        count += 1

                if progress_callback and processed % PROGRESS_INTERVAL == 0:
                    progress_callback(f'已处理 {processed:,} 行，保留 {count:,} 行...')
        return count
    except Exception as e:
        messagebox.showerror('错误', f'处理过程中发生错误: {e}')
        return None
    finally:
        if own:
            f_out.close()


def sort_collected_by_freq(collected):
    """按词频对收集表进行原地排序（上词频 / 简词提取共用）

    collected 中每项为 (orig, result, is_num, num_val)：
    - is_num=True 的行按 -num_val 升序（即数字降序）排在前
    - is_num=False 的行保持原顺序（稳定排序）置后
    - 若全部非数字则不排序（避免无意义重排）

    说明：简词提取中每个正则对应的桶（bucket）即是一个收集表，
    与上词频共用本函数完成排序，保证行为一致。
    """
    if any(r[2] for r in collected):
        # 排序键：(0, -num) 表示数字组按降序；(1, 0) 表示非数字组保持原顺序
        # Python sort 稳定，非数字组组内顺序不变
        collected.sort(key=lambda r: (0 if r[2] else 1, -r[3] if r[2] else 0))


def process_lookup(target_path, ref_path, output_path, unmatched_value='#N/A',
                   progress_callback=None, target_enc='auto', ref_enc='auto',
                   sort_by_freq=False, out_stream=None):
    """上词频：在参考文件中查找待处理文件每行的 key，将匹配值追加到行尾

    流式处理，支持千万行级数据：仅匹配字典常驻内存，输入/输出文件逐行处理。

    - 参考文件格式：'key\\tvalue'（仅保留首次出现的 key）
    - 待处理文件每行用 '\\t' 前部分作为 key
    - 输出：原行 + '\\t' + 匹配值（未命中输出 unmatched_value，默认 '#N/A'）
    - sort_by_freq=True 时：按匹配值（数字）降序排序输出，非数字行置后；
      若所有值都非数字则不排序。注意：排序需将所有结果载入内存。
    - 返回 (总行数, 匹配成功数, 重复 Key 列表)
    """
    # 1. 构建匹配字典（流式读取参考文件）
    match_dict = {}
    duplicates = []
    try:
        with _iter_source(ref_path, ref_enc) as f:
            for lineno, raw in enumerate(f, start=1):
                line = raw.strip()
                if not line or line.startswith('#'):
                    continue
                parts = line.split('\t', 1)
                if len(parts) != 2:
                    continue
                key = parts[0].strip()
                val = parts[1].strip()
                if not key:
                    continue
                if key in match_dict:
                    duplicates.append((lineno, key))
                    continue
                match_dict[key] = val
    except Exception as e:
        messagebox.showerror('错误', f'读取参考文件出错: {e}')
        return None

    # 2. 流式处理待处理文件
    total = matched = 0
    PROGRESS_INTERVAL = 100_000
    # 排序模式需要收集所有结果；非排序模式直接边读边写
    collected = [] if sort_by_freq else None

    def _match_line(raw):
        """单行匹配：返回 (orig_line, match_value, is_num, num_val)"""
        nonlocal total, matched
        line = raw.strip()
        if not line or line.startswith('#'):
            return None
        total += 1
        key = line.split('\t', 1)[0].strip()
        if key and key in match_dict:
            result = match_dict[key]
            matched += 1
        else:
            result = unmatched_value
        orig = raw.rstrip('\n').rstrip('\r')
        return orig, result

    fout = out_stream if out_stream is not None else open(
        output_path, 'w', encoding='utf-8', buffering=1 << 20)
    own = out_stream is None
    try:
        if sort_by_freq:
            # 排序模式：先收集所有行
            with _iter_source(target_path, target_enc) as fin:
                for raw in fin:
                    item = _match_line(raw)
                    if item is None:
                        continue
                    orig, result = item
                    # 尝试把匹配值解析为数字
                    try:
                        num_val = float(result)
                        is_num = True
                    except (ValueError, TypeError):
                        num_val = 0.0
                        is_num = False
                    collected.append((orig, result, is_num, num_val))

                    if progress_callback and total % PROGRESS_INTERVAL == 0:
                        progress_callback(f'已处理 {total:,} 行，匹配 {matched:,} 行...')

            # 复用公共排序函数（与简词提取共用同一逻辑）
            sort_collected_by_freq(collected)

            # 写入排序后的结果
            for orig, result, _, _ in collected:
                fout.write(orig)
                fout.write('\t')
                fout.write(result)
                fout.write('\n')
        else:
            # 非排序模式：边读边写
            with _iter_source(target_path, target_enc) as fin:
                for raw in fin:
                    item = _match_line(raw)
                    if item is None:
                        continue
                    orig, result = item
                    fout.write(orig)
                    fout.write('\t')
                    fout.write(result)
                    fout.write('\n')

                    if progress_callback and total % PROGRESS_INTERVAL == 0:
                        progress_callback(f'已处理 {total:,} 行，匹配 {matched:,} 行...')
    except Exception as e:
        messagebox.showerror('错误', f'处理过程中发生错误: {e}')
        return None
    finally:
        if own:
            fout.close()

    return total, matched, duplicates


def _is_subsequence(part, full):
    """判断 part 是否按顺序包含于 full 中（对应 DictTool 的 Util.isPermutationPart）

    如 'gw' 按顺序包含于 'ggwh'（g→…→w），但不包含于 'yg'。
    """
    it = iter(full)
    return all(ch in it for ch in part)


def _shortcode_in_scope(scope_int, char, is_phrase, ref_chars):
    """根据「处理范围」判断某词条是否纳入出简不出全处理（范围外词条原样保留）

    1: 删除参考列表中全部词条全码（参考内词条出简，单字+词组）
    2: 保留参考列表中全部词条全码（参考外词条出简，单字+词组）
    3: 删除参考列表中单字词条全码（参考内单字出简）
    4: 保留参考列表中单字全码（参考外单字出简，不处理词组）
    5: 仅处理单字（忽略参考）
    6: 仅处理词组（忽略参考）
    7: 处理全部词条（忽略参考）
    """
    if scope_int == 1:
        return char in ref_chars
    if scope_int == 2:
        return char not in ref_chars
    if scope_int == 3:
        return char in ref_chars and not is_phrase
    if scope_int == 4:
        return char not in ref_chars and not is_phrase
    if scope_int == 5:
        return not is_phrase
    if scope_int == 6:
        return is_phrase
    if scope_int == 7:
        return True
    return False


def _shortcode_filter_codes(codes, rule_int):
    """按「出简模式」计算某词条应保留的编码列表（算法与 DictTool-master 一致）

    rule_int:
    1: 仅通过编码长度判断简码 —— 只保留编码最短的（可能多个同长）
    2: 通过前部编码是否相同逐步判断简码 —— 按顺序逐条判断：
       若已保留的某简码是当前编码的前缀，则当前编码为全码（删除）；
       否则保留当前编码，并删除已保留的以当前编码为前缀的更长编码。
    3: 通过编码是否完全包含判断简码 —— 按顺序逐条判断：
       若已保留的某简码按字符顺序完全包含于当前编码，则当前编码为全码（删除）；
       否则保留当前编码，并删除已保留的完全包含当前编码的更长编码。
    """
    if rule_int == 1:
        min_len = min(len(c) for c in codes)
        return [c for c in codes if len(c) == min_len]
    kept = []
    if rule_int == 2:
        for code in codes:
            # 当前编码的任一真前缀已被保留 → 当前编码是全码
            has_simple = any(code[:j] in kept for j in range(1, len(code)))
            if not has_simple:
                # 删除已保留的以当前编码为前缀的更长编码，再保留当前编码
                kept = [k for k in kept if not k.startswith(code)]
                kept.append(code)
    else:  # rule_int == 3
        for code in codes:
            # 已保留的某简码按顺序完全包含于当前编码 → 当前编码是全码
            has_simple = any(_is_subsequence(k, code) for k in kept)
            if not has_simple:
                # 删除已保留的完全包含当前编码的更长编码，再保留当前编码
                kept = [k for k in kept if not _is_subsequence(code, k)]
                kept.append(code)
    return kept


def process_shortcode(target_path, ref_path, output_path, rule, scope,
                      progress_callback=None, target_enc='auto', ref_enc='auto',
                      out_stream=None):
    """出简不出全：删除范围内词条的全码，仅保留简码（判定规则与 DictTool-master 一致）

    rule:  出简模式 1=仅通过编码长度 2=前部编码逐步判断 3=编码完全包含判断
    scope: 处理范围 1-4 需要参考文件（按词条是否在参考列表中筛选），5-7 忽略参考。
    第一遍收集范围内各词条的编码列表并按规则计算保留集合（常驻内存），
    第二遍逐行判定输出，支持千万行级数据。
    返回 (写出行数, 冲突词条列表)；冲突仅在规则 1 下统计（同长最短编码多于一个）。
    """
    rule_int = int(rule)
    scope_int = int(scope)
    ref_chars = set()
    if scope_int in (1, 2, 3, 4):
        try:
            with _iter_source(ref_path, ref_enc) as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith('#'):
                        continue
                    # 参考行若含制表符，仅取第一列作为词条
                    ref_chars.add(s.split('\t', 1)[0])
        except Exception as e:
            messagebox.showerror('错误', f'读取参考文件出错: {e}')
            return None

    # 第一遍：按文件顺序收集范围内各词条的编码列表
    char_codes = {}
    total = 0
    try:
        with _iter_source(target_path, target_enc) as f:
            for line in f:
                total += 1
                s = line.rstrip('\n').rstrip('\r')
                if '\t' not in s:
                    continue
                char, code = s.split('\t', 1)
                is_phrase = len(char) > 1
                if _shortcode_in_scope(scope_int, char, is_phrase, ref_chars):
                    char_codes.setdefault(char, []).append(code)
                if progress_callback and total % 100_000 == 0:
                    progress_callback(f'第一遍：已扫描 {total:,} 行...')
    except Exception as e:
        messagebox.showerror('错误', f'读取待处理文件出错: {e}')
        return None

    # 按出简模式计算各词条保留的编码集合
    kept_map = {}
    conflicts = []
    for char, codes in char_codes.items():
        kept = _shortcode_filter_codes(codes, rule_int)
        if rule_int == 1 and len(kept) > 1:
            conflicts.append(char)
        kept_map[char] = set(kept)

    # 第二遍：逐行判定并写出（范围外词条原样保留）
    written = 0
    processed = 0
    fout = out_stream if out_stream is not None else open(
        output_path, 'w', encoding='utf-8', buffering=1 << 20)
    own = out_stream is None
    try:
        with _iter_source(target_path, target_enc) as fin:
            for line in fin:
                processed += 1
                s = line.rstrip('\n').rstrip('\r')
                if '\t' not in s:
                    fout.write(s + '\n')
                    written += 1
                    continue
                char, code = s.split('\t', 1)
                is_phrase = len(char) > 1
                keep = True
                if _shortcode_in_scope(scope_int, char, is_phrase, ref_chars):
                    keep = code in kept_map.get(char, ())
                if keep:
                    fout.write(s + '\n')
                    written += 1
                if progress_callback and processed % 100_000 == 0:
                    progress_callback(f'第二遍：已处理 {processed:,} 行...')
    except Exception as e:
        messagebox.showerror('错误', f'处理过程中发生错误: {e}')
        return None
    finally:
        if own:
            fout.close()

    return written, conflicts


def process_fullcode(target_path, output_path, scope,
                     progress_callback=None, target_enc='auto', out_stream=None):
    """提取全码：删除范围内词条的简码，仅保留最长的编码（出简不出全的反向版本）

    与 process_shortcode 互为反向：
    - 出简不出全：保留最短编码（简码）
    - 提取全码：保留最长编码（全码），同为最长长度的多条均保留

    scope: 处理范围 1=仅处理单字 2=仅处理词组 3=处理全部词条（均忽略参考文件）
    第一遍收集范围内各词条的编码列表并计算保留集合（最长编码），第二遍逐行判定输出。
    返回写出行数。
    """
    # 提取全码序号 1/2/3 映射到 _shortcode_in_scope 的 5/6/7（复用同一范围判定逻辑）
    _scope_map = {1: 5, 2: 6, 3: 7}
    scope_int = _scope_map.get(int(scope), 7)
    # 提取全码无需参考文件，统一传空集合给 _shortcode_in_scope
    ref_chars = set()

    # 第一遍：按文件顺序收集范围内各词条的编码列表
    char_codes = {}
    total = 0
    try:
        with _iter_source(target_path, target_enc) as f:
            for line in f:
                total += 1
                s = line.rstrip('\n').rstrip('\r')
                if '\t' not in s:
                    continue
                char, code = s.split('\t', 1)
                is_phrase = len(char) > 1
                if _shortcode_in_scope(scope_int, char, is_phrase, ref_chars):
                    char_codes.setdefault(char, []).append(code)
                if progress_callback and total % 100_000 == 0:
                    progress_callback(f'第一遍：已扫描 {total:,} 行...')
    except Exception as e:
        messagebox.showerror('错误', f'读取待处理文件出错: {e}')
        return None

    # 计算各词条保留的编码集合：只保留最长的（同长均保留）
    kept_map = {}
    for char, codes in char_codes.items():
        max_len = max(len(c) for c in codes)
        kept_map[char] = set(c for c in codes if len(c) == max_len)

    # 第二遍：逐行判定并写出（范围外词条原样保留）
    written = 0
    processed = 0
    fout = out_stream if out_stream is not None else open(
        output_path, 'w', encoding='utf-8', buffering=1 << 20)
    own = out_stream is None
    try:
        with _iter_source(target_path, target_enc) as fin:
            for line in fin:
                processed += 1
                s = line.rstrip('\n').rstrip('\r')
                if '\t' not in s:
                    fout.write(s + '\n')
                    written += 1
                    continue
                char, code = s.split('\t', 1)
                is_phrase = len(char) > 1
                keep = True
                if _shortcode_in_scope(scope_int, char, is_phrase, ref_chars):
                    keep = code in kept_map.get(char, ())
                if keep:
                    fout.write(s + '\n')
                    written += 1
                if progress_callback and processed % 100_000 == 0:
                    progress_callback(f'第二遍：已处理 {processed:,} 行...')
    except Exception as e:
        messagebox.showerror('错误', f'处理过程中发生错误: {e}')
        return None
    finally:
        if own:
            fout.close()

    return written


def process_dupcode(target_path, output_path, progress_callback=None,
                    target_enc='auto', tab_separated=False, out_stream=None):
    """出重码号：给每行重复出现的编码追加序号后缀（abc→abc1、abc2…），无需参考文件

    流式逐行处理：匹配「非空白 + 空白分隔 + 非空白」的行，对编码做出现计数并追加后缀；
    不匹配的行（注释、空行、多于两列等）原样保留。支持千万行级数据。
    返回写出行数。out_stream 不为 None 时直接写入该内存流（手动输入模式）。
    tab_separated=False（默认）：序号直接拼在编码后（如 abc1）；
    tab_separated=True：序号用 Tab 分隔为第三列（如 abc\\t1），编码与序号不再相连。
    """
    code_counts = {}
    written = 0
    fout = out_stream if out_stream is not None else open(
        output_path, 'w', encoding='utf-8', buffering=1 << 20)
    own = out_stream is None
    try:
        with _iter_source(target_path, target_enc) as fin:
            for line in fin:
                s = line.rstrip('\n').rstrip('\r')
                m = re.match(r'^(\S+)(\s+)(\S+)$', s)
                if m:
                    char, sep, code = m.groups()
                    code_counts[code] = code_counts.get(code, 0) + 1
                    seq = code_counts[code]
                    new_line = (f'{char}\t{code}\t{seq}'
                                if tab_separated
                                else f'{char}{sep}{code}{seq}')
                else:
                    new_line = s
                fout.write(new_line + '\n')
                written += 1
                if progress_callback and written % 100_000 == 0:
                    progress_callback(f'已处理 {written:,} 行...')
    except Exception as e:
        messagebox.showerror('错误', f'处理过程中发生错误: {e}')
        return None
    finally:
        if own:
            fout.close()
    return written


def process_wordextract(target_path, ref_path, regex_list, output_dir,
                        extract_limit=0, merge_results=False,
                        include_regex_header=True, progress_callback=None,
                        target_enc='auto', ref_enc='auto'):
    """简词提取：用多个正则从待处理文件提取行，按词频降序排序，输出到文件

    - 待处理文件格式：word\\tcode（正则对 'word\\tcode' 做 fullmatch）
    - 词频文件格式：word\\tfreq（freq 为数字时用于降序排序；无词频文件按原始顺序）
    - merge_results=False：每个正则输出一个文件（1.txt、2.txt…），output_dir 为目录
    - merge_results=True：所有结果合并到一个文件，output_dir 为文件路径
    - include_regex_header：结果中是否写入「# 匹配正则: ...」标题行
    - extract_limit > 0 时，每个正则只保留排序后的前 extract_limit 行
    返回 (总匹配行数, 各正则行数列表)；失败返回 None。
    """
    # 1. 加载词频字典（无词频文件则跳过，按原始顺序输出）
    freq_dict = {}
    if ref_path:
        try:
            with _iter_source(ref_path, ref_enc) as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith('#'):
                        continue
                    parts = s.split('\t', 1)
                    if len(parts) == 2:
                        freq_dict[parts[0].strip()] = parts[1].strip()
        except Exception as e:
            messagebox.showerror('错误', f'读取词频文件出错: {e}')
            return None

    # 2. 预编译正则（空行跳过，重复正则去重）
    compiled = []
    seen_patterns = set()
    for pattern in regex_list:
        p = pattern.strip()
        if not p or p in seen_patterns:
            continue
        try:
            compiled.append((p, re.compile(p)))
            seen_patterns.add(p)
        except re.error as e:
            messagebox.showerror('错误', f'正则表达式错误: {p}\n{e}')
            return None
    if not compiled:
        messagebox.showwarning('提示', '请输入至少一个有效的正则表达式。')
        return None

    # 3. 收集每个正则匹配的行（word\tcode, freq）
    buckets = [[] for _ in compiled]
    total = 0
    PROGRESS_INTERVAL = 100_000
    try:
        with _iter_source(target_path, target_enc) as fin:
            for line in fin:
                s = line.rstrip('\n').rstrip('\r')
                if not s or s.startswith('#'):
                    continue
                parts = s.split('\t', 1)
                if len(parts) < 2:
                    continue
                data_line = f'{parts[0]}\t{parts[1]}'
                freq = freq_dict.get(parts[0], '')
                for idx, (_, regex) in enumerate(compiled):
                    if regex.fullmatch(data_line):
                        buckets[idx].append((data_line, freq))
                        total += 1
                if progress_callback and total % PROGRESS_INTERVAL == 0:
                    progress_callback(f'已匹配 {total:,} 行...')
    except Exception as e:
        messagebox.showerror('错误', f'处理过程中发生错误: {e}')
        return None

    # 4. 对每个桶（收集表）按词频排序 + 截取
    #    每个正则对应一个收集表，复用上词频的排序逻辑（sort_collected_by_freq）
    #    未匹配词频（空字符串）当作 0 处理，归入数字组
    for idx in range(len(compiled)):
        # 将 (data_line, freq) 转成收集表格式 (orig, result, is_num, num_val)
        collected = []
        for data_line, freq in buckets[idx]:
            if not freq:  # 空字符串 = 未匹配，当作 0
                is_num, num_val = True, 0.0
            else:
                try:
                    num_val = float(freq)
                    is_num = True
                except (ValueError, TypeError):
                    is_num, num_val = False, 0.0
            collected.append((data_line, freq, is_num, num_val))
        sort_collected_by_freq(collected)
        # 截取前 N 行（若启用），再还原为 (data_line, freq) 形式
        if extract_limit > 0:
            collected = collected[:extract_limit]
        buckets[idx] = [(orig, result) for orig, result, _, _ in collected]
    counts = [len(buckets[i]) for i in range(len(compiled))]

    try:
        if merge_results:
            # 合并模式：生成结果文本，由调用方决定是否写文件（支持结果行数过滤）
            lines = []
            for idx, (pattern, _) in enumerate(compiled):
                if include_regex_header:
                    lines.append(f'# 匹配正则: {pattern}')
                for data_line, freq in buckets[idx]:
                    lines.append(f'{data_line}\t{freq}' if freq else data_line)
                if idx < len(compiled) - 1:
                    lines.append('')  # 块间空行
            result_text = '\n'.join(lines) + '\n' if lines else ''
            return total, counts, result_text
        else:
            # 非合并模式：每个正则一个文件
            os.makedirs(output_dir, exist_ok=True)
            for idx, (pattern, _) in enumerate(compiled):
                out_file = os.path.join(output_dir, f'{idx + 1}.txt')
                with open(out_file, 'w', encoding='utf-8',
                          buffering=1 << 20) as fout:
                    if include_regex_header:
                        fout.write(f'# 匹配正则: {pattern}\n')
                    for data_line, freq in buckets[idx]:
                        fout.write(f'{data_line}\t{freq}\n' if freq
                                   else f'{data_line}\n')
    except Exception as e:
        messagebox.showerror('错误', f'写入输出文件失败: {e}')
        return None

    return total, counts, None


class _OptionRow:
    """常驻占位的选项行：始终占用固定高度，切换操作时仅在真实内容与等高占位间切换，
    不会引起下方布局上下移动。新增选项只需在 App._build_ui 的 self.option_rows 中
    登记一行 build 函数（build_fn(parent) 在该行显示时填充控件）。"""

    def __init__(self, parent, build_fn):
        self.frame = ttk.Frame(parent)
        # 所有子选项行都放在同一格 (0,0)，切换时仅显示当前操作的一行，保证始终在同一行
        self.frame.grid(row=0, column=0, sticky='ew', pady=(6, 0))
        # 先构建真实内容（暂不测量高度）：所有选项行建好后由 _build_ui 统一刷新一次几何布局，
        # 避免每行各自 update_idletasks 造成的多次全量重排，缩短启动时间
        self.content = ttk.Frame(self.frame)
        build_fn(self.content)
        self.content.pack(fill='x')
        self.h = 22  # 占位高度，_measure() 时替换为真实高度

    def _measure(self):
        """统一刷新后测量真实内容高度（必须在一次 update_idletasks 之后调用）"""
        self.h = max(self.content.winfo_reqheight(), 22)
        # 固定行高，禁止内容撑开框架（占位与真实内容等高，切换不跳动）
        self.frame.pack_propagate(False)
        self.frame.config(height=self.h)
        self.spacer = ttk.Frame(self.frame, height=self.h)
        self.spacer.pack(fill='x')
        self.content.pack_forget()  # 默认显示占位

    def show(self, visible):
        if visible:
            self.frame.grid(row=0, column=0, sticky='ew')
            self.spacer.pack_forget()
            self.content.pack(fill='x')
        else:
            self.content.pack_forget()
            self.spacer.pack(fill='x')
            self.frame.grid_remove()  # 整行移出布局，仅当前操作的子选项占位同一行

    def set_height(self, h):
        """统一所有可选行到同一高度（视觉一致、无跳动）"""
        self.h = h
        self.frame.config(height=h)
        self.spacer.config(height=h)


class _Preview:
    """带行号栏的预览框。

    - line_numbers=True 时左侧显示行号（随滚动同步）。
    - editable=True 时文本框可手动编辑，并限制最多 MAX_LINES 行，超出弹窗提示后截断。
    新增预览框只需调用 _make_preview(...) 并传入对应参数。
    """

    MAX_LINES = 10000  # 手动输入行数上限

    def __init__(self, parent, text_font, height=8, width=80, editable=False,
                 line_numbers=True, on_overlimit=None):
        self.line_numbers = line_numbers
        self.on_overlimit = on_overlimit
        self.editable = editable

        self.frame = ttk.Frame(parent)
        self.frame.pack(fill='x', pady=(1, 0))
        self.inner = ttk.Frame(self.frame, relief='sunken', borderwidth=1)
        self.inner.pack(fill='x')

        if line_numbers:
            self.gutter = tk.Text(
                self.inner, width=6, height=height, state='disabled',
                wrap='none', font=text_font, padx=3, pady=2,
                takefocus=0, bg='#eef0f2', fg='#9098a3',
                relief='flat', borderwidth=0, cursor='arrow',
                spacing1=0, spacing2=0, spacing3=0)
        else:
            self.gutter = None

        self.text = tk.Text(
            self.inner, height=height, width=width, wrap='none',
            state='disabled' if not editable else 'normal',
            font=text_font, padx=4, pady=2,
            relief='flat', borderwidth=0,
            spacing1=0, spacing2=0, spacing3=0)

        # 滚动条：使用 ttk.Scrollbar + clam 主题样式，颜色配置生效，thumb 始终可见。
        # Windows 默认主题下 tk.Scrollbar 与 ttk.Scrollbar 均用原生渲染，
        # 颜色参数被忽略，某些视觉样式下 thumb 与 trough 对比度不足导致看不到滑块。
        # clam 是 Tk 内置纯绘制主题，troughcolor/background 等参数完全生效。
        _init_preview_scrollbar_style()
        self.sb = ttk.Scrollbar(
            self.inner, orient='vertical',
            command=self._on_sb_scroll,
            style='Preview.Vertical.TScrollbar')
        self.text.config(yscrollcommand=self._on_text_scroll)

        # 关键 pack 顺序：先固定侧边的 gutter(left) 与 sb(right)，
        # 最后 pack text(side='left', expand=True) 占据剩余空间。
        # 若先 pack expand 的 text，小窗口下 text 会抢占所有水平空间，
        # 导致后 pack 的 sb 被挤出可见区（滚动条整个被裁切看不到）。
        if self.gutter:
            self.gutter.pack(side='left', fill='y')
        self.sb.pack(side='right', fill='y')
        self.text.pack(side='left', fill='both', expand=True)

        # ---- 右键上下文菜单 ----
        self.ctx_menu = tk.Menu(self.text, tearoff=0)
        self.ctx_menu.add_command(label='复制  Ctrl+C', command=self._ctx_copy)
        self.ctx_menu.add_command(label='粘贴  Ctrl+V', command=self._ctx_paste)
        self.ctx_menu.add_command(label='剪切  Ctrl+X', command=self._ctx_cut)
        self.text.bind('<Button-3>' if sys.platform != 'darwin' else '<Button-2>',
                        self._show_context_menu)

        if editable:
            self.text.bind('<KeyRelease>', self._on_change)
            # 粘贴后内容可能瞬间超出上限，粘贴事件后再校验一次
            self.text.bind('<<Paste>>', lambda e: self.text.after(1, self._on_change))
            self.text.bind('<<Cut>>', lambda e: self.text.after(1, self._on_change))

    # ---- 滚动与行号同步 ----
    def _on_sb_scroll(self, *args):
        self.text.yview(*args)
        self._sync_gutter()

    def _on_text_scroll(self, *args):
        self.sb.set(*args)
        self._sync_gutter()

    def _sync_gutter(self):
        if self.gutter:
            self.gutter.yview_moveto(self.text.yview()[0])

    def _update_line_numbers(self):
        if not self.gutter:
            return
        n = int(self.text.index('end-1c').split('.')[0])
        nums = '\n'.join(str(i) for i in range(1, n + 1))
        self.gutter.config(state='normal')
        self.gutter.delete('1.0', tk.END)
        self.gutter.insert('1.0', nums)
        self.gutter.config(state='disabled')
        self._sync_gutter()

    # ---- 内容编辑与限制 ----
    def _on_change(self, event=None):
        self._enforce_limit()
        self._update_line_numbers()

    def _enforce_limit(self):
        # 仅在可编辑且超出上限时截断并提示
        if not self.editable:
            return
        n = int(self.text.index('end-1c').split('.')[0])
        if n > self.MAX_LINES:
            self.text.config(state='normal')
            self.text.delete(f'{self.MAX_LINES + 1}.0', tk.END)
            self.text.config(state='normal')
            self.text.see('1.0')
            if self.on_overlimit:
                self.on_overlimit(n)

    # ---- 右键上下文菜单 ----
    def _show_context_menu(self, event):
        """在鼠标位置弹出右键菜单，根据选中/编辑状态启用对应项"""
        # 判断是否有选中内容
        try:
            has_sel = bool(self.text.tag_ranges(tk.SEL))
        except tk.TclError:
            has_sel = False
        # 复制：有选中则可用
        self.ctx_menu.entryconfigure(0, state='normal' if has_sel else 'disabled')
        # 粘贴：可编辑则可用
        self.ctx_menu.entryconfigure(1, state='normal' if self.editable else 'disabled')
        # 剪切：可编辑且有选中则可用
        self.ctx_menu.entryconfigure(2, state='normal' if (self.editable and has_sel) else 'disabled')
        self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def _ctx_copy(self):
        try:
            self.text.event_generate('<<Copy>>')
        except tk.TclError:
            pass

    def _ctx_paste(self):
        try:
            self.text.event_generate('<<Paste>>')
        except tk.TclError:
            pass

    def _ctx_cut(self):
        try:
            self.text.event_generate('<<Cut>>')
        except tk.TclError:
            pass

    def set_content(self, content):
        """写入内容并刷新行号；保持当前 editable 状态"""
        self.text.config(state='normal')
        self.text.delete('1.0', tk.END)
        self.text.insert('1.0', content)
        self.text.config(state='normal' if self.editable else 'disabled')
        self._update_line_numbers()

    def get_content(self):
        return self.text.get('1.0', 'end-1c')

    def set_editable(self, editable):
        self.editable = editable
        self.text.config(state='normal' if editable else 'disabled')
        self._update_line_numbers()


class App:
    # 各操作对应的默认输出文件名（仅文件名，不含目录）
    DEFAULT_OUTPUT_NAMES = {
        'difference': '差集-处理结果.txt',
        'intersection': '交集-处理结果.txt',
        'lookup': '上词频-处理结果.txt',
        'shortcode': '出简不出全-处理结果.txt',
        'fullcode': '提取全码-处理结果.txt',
        'dupcode': '重码号-处理结果.txt',
        'wordextract': '简词提取结果',
        'phrasecode': '词组编码结果.txt',
    }

    # 操作类型定义：新增操作只需在此登记
    # needs_ref: True 需要参考文件 / False 不需要 / 'short_mode' 由出简处理范围决定
    # 另需在 OP_RADIO_LAYOUT、self.option_rows、_run 中补充对应逻辑
    OPERATIONS = [
        ('difference', '差集（保留不在参考中的行）', True),
        ('intersection', '交集（保留在参考中的行）', True),
        ('lookup', '上词频（查找并追加匹配值）', True),
        ('shortcode', '出简不出全（删除全码保留简码）', 'short_mode'),
        ('fullcode', '提取全码（保留最长编码删除简码）', False),
        ('dupcode', '出重码号（重复编码加序号后缀）', False),
        ('wordextract', '简词提取（按正则提取并按词频排序）', True),
        ('phrasecode', '词组编码（给无编码词条编码）', False),
    ]
    # 操作单选按钮在 3 行 3 列网格中的 (operation_id, 行, 列) 布局（左对齐）
    # 列号从 1 开始（0 列留空，原「操作类型:」标签已删除）
    # 顺序调整：简词提取描述最长，单独放第三行第一列，避免挤压第三列导致描述被截断
    OP_RADIO_LAYOUT = [
        ('difference', 0, 1), ('intersection', 0, 2), ('lookup', 0, 3),
        ('dupcode', 1, 1), ('shortcode', 1, 2), ('fullcode', 1, 3),
        ('wordextract', 2, 1),
        ('phrasecode', 2, 2),
    ]

    def __init__(self, root):
        self.root = root
        self.root.title('词库处理工具')
        # 窗口图标（Base64 内嵌，无需外部文件）
        self._icon = tk.PhotoImage(data=(
            'iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwY'
            'AAADCUlEQVR4nO2aX4gNURzHP9e/FGLt5s+iZVfhWRTyQinUWskLEQ+kbHmTJxIPS1FeFG/+'
            'ZUtqPSjRhl0P5IFVREhhL2KXsOvuWjs69b11mu7uzJ2Zu3dmm2/9mnvvzPmd+cz5nTnn/M6F'
            'VKlSpZLuAjmgW/YVeAHcAy4Bh4CNQCXFqwrYABwGLgLtwBvgG9ALODJTV2g5Pu0f8Ehgc4bx'
            'NxM4CDxUGb/+QyvvaAJQAcwGFgFrgd1Ak1rtl3XtgFqr1vJjyp0H+qzrfgP3gZPytUa+q1RX'
            'phQgXjKg6xUiOZUxcPuABoWL+a0faAbqgYkR1l8SR7P09Add4XEBmFekL6ecIHltVj8wobZt'
            'pOuv0NN8CbyN4Ink1BeCypHl35g31N881RzxWyMqEMeyFj8FTaf8C1S7HJUbxGi61Tqe6pER'
            'QxCs17anUpAhlIJIKUhPxCAdowXEaBmjAGRqiLKxAWnSXMtM83cBU5IKctw1tTB+LwPrgLFJ'
            'Allslf/jmtY/9rk8duIAgpa/pnwjMF8Lr7zPJz5gnLiA7Ff5m8AkoFXf+3zCOHEBqdRNmxn1'
            'F8vfTq15vGCcuIAYXbf8mMXaLSuR4QXjxAlkAfBeftqAydY5LxgnLiALdawLCNMXF5BnwBJ9r'
            'gHeyd8D1wBpEnfPde6pclsoE5mLEuQOcA24ApwDTgPHgD1K1pnwGVfAX1ZpUHy0TPUQLbMpC'
            'pCuIRIAhaxX05GjwEplCrOqOGPdWBCY0CAmxre6bAdwQDd8FritkHHncl9rRB/Q+R/A8pAw/'
            'UFBitE0xfQZ15iRty6FZCYEzBY9kJKC2Bqv2DY38QHYrt/csmHafcLUjySIrTEe54PAeOq74'
            'jDIxk0YRQ7TogL2LlUprRNY7QNmb5GzZmYAV4HPIwRi7COwYhiYBmuE7w4SZuVSnQXTYUGcU'
            'Ji9sgbp2KvWgnH02kZbc53WmmYVCYNpA5ZaEFkdzQ4zSYMZ1LFVfeenvgfJzsQizFqAU/r8i'
            'YRpLnCkwDa4mTEkUjX6p0RTyFRsKkqp/77aJnRGCooMAAAAAElFTkSuQmCC'
        ))
        self.root.iconphoto(True, self._icon)
        # 先隐藏窗口：后续 _build_ui 中的 _OptionRow 会调用 update_idletasks() 量高度，
        # 若此时窗口可见会强制绘制造成启动闪屏；全部布局算好后再 deiconify()
        self.root.withdraw()

        # 提升高分屏文字清晰度
        try:
            default_font = tkfont.nametofont('TkDefaultFont')
            default_font.configure(size=11, family=CJK_FONT)
            # 预览使用与界面一致的标准 CJK 字体，避免中文触发等宽字体回退导致显示异常
            text_font = default_font
        except Exception:
            text_font = None

        # 参考文件区「无需参考」时置灰用的 LabelFrame 样式（标题变灰）
        try:
            ttk.Style().configure('RefGrey.TLabelFrame', foreground='#808080')
        except Exception:
            pass

        self.target_path = tk.StringVar()
        self.ref_path = tk.StringVar()
        self.operation = tk.StringVar(value='difference')
        # 忽略第二列默认不勾选（避免用户未注意时误丢第二列信息）
        self.ignore_target = tk.BooleanVar(value=False)
        self.ignore_ref = tk.BooleanVar(value=False)
        self.output_name = tk.StringVar(value=self.DEFAULT_OUTPUT_NAMES['difference'])
        # 出简不出全：出简模式（判定简码的算法，与 DictTool-master 一致）
        # Combobox 选中项的完整文本，取首字符数字作为编号
        self.shortcode_rule = tk.StringVar(
            value='1. 仅通过编码长度判断简码')
        # 出简不出全：处理范围（哪些词条参与出简，其余原样保留）
        self.shortcode_scope = tk.StringVar(
            value='7. 处理全部词条（忽略参考）')
        # 提取全码：处理范围（只有 1/2/3 三项，均忽略参考文件）
        self.fullcode_scope = tk.StringVar(
            value='3. 处理全部词条')
        # 出重码号：序号是否用 Tab 分隔（默认不勾选，数字直接拼在编码后）
        self.dupcode_tab = tk.BooleanVar(value=False)
        self.lines_t = tk.StringVar()
        self.lines_r = tk.StringVar()
        self.save_path_var = tk.StringVar()
        self.unmatched_value = tk.StringVar(value='#N/A')
        # 上词频：是否按匹配值（数字）降序排序输出
        self.sort_by_freq = tk.BooleanVar(value=False)
        # 简词提取：提取前N行（勾选生效，默认勾选，默认值100）
        self.extract_limit_enabled = tk.BooleanVar(value=True)
        self.extract_limit = tk.StringVar(value='100')
        # 简词提取：将所有正则结果合并成一个文件（勾选时默认仅提取前1行）
        self.merge_results = tk.BooleanVar(value=False)
        # 简词提取：结果包含正则标题行（默认勾选）
        self.include_regex_header = tk.BooleanVar(value=True)
        # 编码选择（待处理 / 参考文件各一），'auto' 走自动检测
        self.target_enc = tk.StringVar(value='auto')
        self.ref_enc = tk.StringVar(value='auto')
        self.preset_var = tk.StringVar()  # 预设下拉框选中项
        # 预设目录解析：优先取脚本所在目录；PyInstaller onefile 下 __file__ 指向临时解压目录，
        # 此时回退到 _MEIPASS 解压根，保证源码 / onedir / onefile 三种运行方式都能找到 预设 文件夹
        self.preset_dir = self._resolve_preset_dir()
        self.last_output_path = None  # 最近一次输出文件路径，供「打开输出位置」使用
        self._about_win = None  # 关于窗口引用，避免重复弹出（已存在则聚焦）
        self.preview_only_enabled = tk.BooleanVar(value=True)  # 结果行数过滤：默认勾选
        self.preview_only_threshold = tk.StringVar(value='100')  # 行数阈值，默认 100

        # ---- 词组编码功能状态 ----
        # 文件区直接复用标准 frm_files（待处理文件=待编码词组 / 参考文件=单字表），
        # 仅通过 _on_operation_change 改标签，不新建独立文件区。
        self._ui_pad = {'padx': 8, 'pady': 1}
        self.allow_predef = tk.BooleanVar(value=True)    # 允许预定义编码（默认勾选）
        self.ignore_punct = tk.BooleanVar(value=True)     # 忽略标点符号（默认勾选）
        self.ignore_extra = tk.StringVar(value=DEFAULT_IGNORE_BOX)   # 忽略表输入框（默认含全部忽略字符）

        self._build_ui(text_font)

        # 输出文件名变化时同步刷新保存位置提示
        self.output_name.trace_add('write', lambda *a: self._update_save_label())
        # 操作类型变化时联动「忽略第二列」复选框状态
        self.operation.trace_add('write', lambda *a: self._on_operation_change())
        # 出简不出全处理范围变化时，影响是否需要参考文件（整体变灰/恢复正常）
        self.shortcode_scope.trace_add(
            'write', lambda *a: (self._update_reference_state(), self._update_status()))
        self._on_operation_change()
        self._update_reference_state()
        self._update_save_label()
        self._update_status()
        # 初始化复选框关联输入框的置灰状态
        self._update_preview_only_state()
        self._update_extract_limit_state()

        # 状态栏用 side='bottom' 钉在底部，会让 winfo_reqheight 不可靠、固定大尺寸又留白。
        # 上词频行与出简不出全行已常驻占位（即使不显示也保留一行高度，切换时不再上下移动），
        # 这里临时把状态栏改回 side='top'，并把这两行的真实内容都显示出来测出最大高度，
        # 从而精确设定窗口高度：既不隐藏也不截断，切到任意模式底部提示都不会被挡。
        # 宽度适当放大以确保各操作括号描述完整显示。
        # 窗口已在 __init__ 开头 withdraw()，此处测量在隐藏状态下完成，定好尺寸后再显示。
        content_h, w = self._measure_full()
        # 状态栏钉在底部、始终可见；先以最长示例文本测出其最大高度，
        # 避免处理完成后（含长路径）的多行提示被固定窗口高度裁掉
        self.status.pack(side='bottom', pady=4, fill='x')
        self.status.config(wraplength=w - 20)
        self.status.config(
            text=('完成！共处理 9,999,999 行，匹配成功 9,999,999 行，失败 0 行。'
                  ' 已按匹配值降序排序。 参考文件有 99 处重复 Key（已保留首次值）。'
                  '\n结果已保存至: C:\\很长的目录名\\更长的子目录名\\输出结果文件.txt'))
        self.root.update_idletasks()
        status_h = self.status.winfo_reqheight()
        self._update_status()  # 还原为当前真实的初始提示

        # 窗口宽度：以内容自然宽度为准（保证描述完整显示），并限定合理区间；
        # 额外 +18 像素留给滚动条。内容过宽时窗口也不会无限变宽（封顶 1400）。
        w = min(max(w, 1080), 1400)
        # 窗口高度：平时贴合（内容高度 + 状态栏高度）；但封顶到舒适高度 WIN_MAX_H，
        # 超出则整页滚动，避免窗口被撑到接近整屏（之前“高度异常/顶部留白”的根因）。
        WIN_MAX_H = 1400
        max_h = max(560, self.root.winfo_screenheight() - 48 - 20)
        h = content_h + status_h
        h = min(max(h, 560), WIN_MAX_H, max_h)
        # 窗口水平居中；顶部留 40px 余量（不再紧贴屏幕顶边），同时不进入底部任务栏范围
        x = max(0, (self.root.winfo_screenwidth() - (w + 18)) // 2)
        y = 40
        self.root.geometry(f'{w + 18}x{h}+{x}+{y}')
        self.root.minsize(980, 560)
        self.root.deiconify()  # 尺寸已全部定好，再显示窗口，避免启动抖动
        # 显式让内容区宽度跟随画布宽度（横向铺满、无右侧空隙），并刷新滚动区域
        self.root.update_idletasks()
        self._sync_content_width(self.canvas.winfo_width())
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
        # 路径标签随窗口尺寸自动换行
        self.root.bind('<Configure>', lambda e: self._update_path_wrap())
        self._update_path_wrap()

    @staticmethod
    def _resolve_preset_dir():
        """解析 预设 文件夹路径，兼容源码 / onedir / onefile 三种运行方式。

        - 源码态：脚本与 预设 同目录。
        - onedir（PyInstaller 6.x）：依赖与 datas 被收集到 exe 同级的 _internal 子目录，
          故需额外探测 <根>/_internal/预设；同时以 sys.executable 所在目录作为可靠的兜底根。
        - onefile：依赖解压到 _MEIPASS，仅当打包时通过 --add-data 包含 预设 才会存在。
        """
        # 候选根目录（去重、保序）
        roots = []
        # 模块文件所在目录：源码态=脚本目录；冻结态通常指向 exe 或 _internal 内的脚本
        roots.append(os.path.dirname(os.path.abspath(__file__)))
        if getattr(sys, 'frozen', False):
            # 可执行文件所在目录最可靠：onedir 下预设可能被放在这里或其 _internal 子目录
            roots.append(os.path.dirname(os.path.abspath(sys.executable)))
            meipass = getattr(sys, '_MEIPASS', '')
            if meipass:
                roots.append(meipass)

        for r in roots:
            # 直接在根目录下找 预设
            d = os.path.join(r, '预设')
            if os.path.isdir(d):
                return d
            # PyInstaller 6.x onedir 把收集内容放到 _internal 子目录
            d = os.path.join(r, '_internal', '预设')
            if os.path.isdir(d):
                return d
        # 都找不到时回退到模块目录（保持原有报错路径，便于定位）
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '预设')

    def _measure_full(self):
        """测量同时显示所有可选选项行真实内容时的最大高度与宽度，测量后还原为占位状态"""
        self.status.pack_forget()
        self.status.pack(side='top', fill='x', pady=4)
        for row in self.option_rows.values():
            row.show(True)
        self.root.update_idletasks()
        # 内容已包进可滚动容器，量其内在高度/宽度即可（与状态栏无关）
        h = self.content.winfo_reqheight()
        w = self.content.winfo_reqwidth()
        for row in self.option_rows.values():
            row.show(False)
        self.status.pack_forget()
        return h, w

    def _on_mousewheel(self, event):
        """鼠标滚轮滚动整页；弹窗（Toplevel）内不拦截，由其自身处理"""
        if event.widget.winfo_toplevel() is not self.root:
            return
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def _on_canvas_configure(self, event):
        """画布尺寸变化时：内容区宽度跟随画布宽度（横向铺满），并刷新滚动区域"""
        self._sync_content_width(event.width)
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _on_content_configure(self, event):
        """内容尺寸变化（如切换操作类型、显示/隐藏选项行）时刷新滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _sync_content_width(self, width):
        """让可滚动内容区宽度铺满画布宽度，避免出现右侧空白"""
        if width and width > 1:
            self.canvas.itemconfigure(self._content_win, width=width)

    def _build_ui(self, text_font):
        pad = {'padx': 8, 'pady': 1}
        self._text_font = text_font
        self._ui_pad = pad

        # ---- 整页滚动容器：除底部状态栏外的内容都放进可滚动区域，
        #      窗口保持舒适高度，内容过高时整页纵向滚动，不再截断底部 ----
        # 用一个容器框同时承载 canvas + 滚动条，避免 canvas 直接 expand 时
        # 与底部状态栏争夺高度（会把状态栏压成 0 高、或令窗口高度异常）。
        self.frm_scroll = ttk.Frame(self.root)
        self.frm_scroll.pack(fill='both', expand=True)
        self.canvas = tk.Canvas(self.frm_scroll, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(
            self.frm_scroll, orient='vertical', command=self.canvas.yview,
            style='Preview.Vertical.TScrollbar')
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.content = ttk.Frame(self.canvas)
        self._content_win = self.canvas.create_window(
            (0, 0), window=self.content, anchor='nw')
        self.scrollbar.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)
        # 画布尺寸变化：让内部内容区宽度跟随画布宽度（横向铺满，不留右侧空隙），
        # 并刷新滚动区域；内容尺寸变化也刷新滚动区域。
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.content.bind('<Configure>', self._on_content_configure)
        # 鼠标滚轮滚动整页（弹窗内不拦截，避免影响弹窗自身滚动）
        self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)

        # ---- 文件选择区（两个并排，预览框不互相挤压）----
        self.frm_files = ttk.Frame(self.content, padding=4)
        frm_files = self.frm_files
        # fill='x' 不 expand：文件区按内容高度紧凑显示，避免窗口变大时
        # 预览框被拉高导致下方出现大片无用空白；剩余空间交给结果预览区吸收。
        frm_files.pack(fill='x', **pad)
        frm_files.columnconfigure(0, weight=1)
        frm_files.columnconfigure(1, weight=1)

        # 待处理文件（词组编码模式复用为「待编码词组」）
        self.frm_t = ttk.LabelFrame(frm_files, text='待处理文件', padding=2)
        self.frm_t.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        frm_t_btn = ttk.Frame(self.frm_t)
        frm_t_btn.pack(anchor='w')
        self.btn_import_target = ttk.Button(
            frm_t_btn, text='导入待处理文件', command=self._import_target)
        self.btn_import_target.pack(side='left')
        ttk.Button(frm_t_btn, text='清空', width=6, command=self._clear_target).pack(
            side='left', padx=6)
        self.lbl_target_path = ttk.Label(self.frm_t, textvariable=self.target_path,
                                         foreground='blue', wraplength=380,
                                         font=(CJK_FONT, 8))
        self.lbl_target_path.pack(anchor='w', pady=2)
        ttk.Label(self.frm_t, textvariable=self.lines_t, foreground='gray').pack(anchor='w')
        self.chk_target = ttk.Checkbutton(self.frm_t, text='忽略第二列（编码）',
                                          variable=self.ignore_target)
        self.chk_target.pack(anchor='w', pady=2)
        # 编码选择行：手动指定编码可解决自动检测误判导致的乱码
        frm_t_enc = ttk.Frame(self.frm_t)
        frm_t_enc.pack(fill='x', pady=2)
        ttk.Label(frm_t_enc, text='编码:').pack(side='left')
        self.cmb_enc_t = ttk.Combobox(
            frm_t_enc, textvariable=self.target_enc, width=10, state='readonly',
            values=[v for _, v in ENCODING_CHOICES])
        self.cmb_enc_t.pack(side='left', padx=4)
        # 切换编码时使用所选编码重新加载预览
        self.cmb_enc_t.bind('<<ComboboxSelected>>', lambda e: self._reload_target_preview())
        # 未导入文件时，待处理预览可手动编辑（带行号、限 1 万行）；导入文件后转只读
        self.preview_t = self._make_preview(
            self.frm_t, text_font, editable=True, on_overlimit=self._on_preview_overlimit)

        # 参考文件
        self.frm_r = ttk.LabelFrame(frm_files, text='参考文件', padding=2)
        self.frm_r.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        frm_r_btn = ttk.Frame(self.frm_r)
        frm_r_btn.pack(anchor='w')
        self.btn_import_ref = ttk.Button(frm_r_btn, text='导入参考文件',
                                         command=self._import_ref)
        self.btn_import_ref.pack(side='left')
        self.btn_clear_ref = ttk.Button(frm_r_btn, text='清空', width=6,
                                        command=self._clear_ref)
        self.btn_clear_ref.pack(side='left', padx=6)
        self.lbl_ref_path = ttk.Label(self.frm_r, textvariable=self.ref_path,
                                      foreground='blue', wraplength=380,
                                      font=(CJK_FONT, 8))
        self.lbl_ref_path.pack(anchor='w', pady=2)
        self.lbl_ref_lines = ttk.Label(self.frm_r, textvariable=self.lines_r,
                                       foreground='gray')
        self.lbl_ref_lines.pack(anchor='w')
        self.chk_ref = ttk.Checkbutton(self.frm_r, text='忽略第二列（编码）',
                                       variable=self.ignore_ref)
        self.chk_ref.pack(anchor='w', pady=2)
        # 编码选择行
        frm_r_enc = ttk.Frame(self.frm_r)
        frm_r_enc.pack(fill='x', pady=2)
        ttk.Label(frm_r_enc, text='编码:').pack(side='left')
        self.cmb_enc_r = ttk.Combobox(
            frm_r_enc, textvariable=self.ref_enc, width=10, state='readonly',
            values=[v for _, v in ENCODING_CHOICES])
        self.cmb_enc_r.pack(side='left', padx=4)
        self.cmb_enc_r.bind('<<ComboboxSelected>>', lambda e: self._reload_ref_preview())
        # 预设下拉框：从程序文件夹\预设 载入选中预设内容（UTF-8，可能含 BOM），置于编码右侧
        ttk.Label(frm_r_enc, text='载入预设:').pack(side='left')
        self.cmb_preset = ttk.Combobox(
            frm_r_enc, textvariable=self.preset_var, width=22, state='readonly')
        self.cmb_preset.pack(side='left', padx=4)
        self.cmb_preset.bind(
            '<<ComboboxSelected>>',
            lambda e: self._load_preset(self.preset_var.get()))
        if os.path.isdir(self.preset_dir):
            self.cmb_preset['values'] = [''] + sorted(
                f for f in os.listdir(self.preset_dir)
                if os.path.isfile(os.path.join(self.preset_dir, f)))
        else:
            # 诊断：预设目录未找到时给出明确提示与解析路径，便于定位打包问题
            messagebox.showwarning(
                '预设目录未找到',
                f'未能定位 预设 文件夹，载入预设功能暂不可用。\n\n'
                f'程序解析到的路径：\n{self.preset_dir}\n\n'
                f'请确认该路径下存在「预设」文件夹（打包后需与 exe 放在一起）。')
        # 未导入文件时，参考预览可手动编辑（带行号、限 1 万行）；导入文件后转只读
        self.preview_r = self._make_preview(
            self.frm_r, text_font, editable=True, on_overlimit=self._on_preview_overlimit)

        # ---- 操作按钮区 ----
        # 互换按钮居中于整行：用内层 frame + grid 单列 weight=1 让按钮在 cell 内居中；
        # 关于按钮用 place 定位到右侧，不参与 pack/grid 布局流，
        # 避免占用布局空间导致互换按钮居中计算偏移。
        # 关于按钮 width=6 与「清空」按钮等宽；行高由互换按钮决定，与原 pack 一致。
        frm_act = ttk.Frame(self.content, padding=0)
        frm_act.pack(fill='x', padx=8, pady=1)
        frm_act_inner = ttk.Frame(frm_act)
        frm_act_inner.pack(fill='x')
        frm_act_inner.columnconfigure(0, weight=1)
        ttk.Button(frm_act_inner, text='⇄ 互换两个文件',
                   command=self._swap).grid(row=0, column=0)
        ttk.Button(frm_act, text='关于', width=6,
                   command=self._show_about).place(relx=1.0, rely=0.5, anchor='e')

        # ---- 选项 ----
        self.frm_opt = ttk.LabelFrame(self.content, text='处理选项', padding=8)
        self.frm_opt.pack(fill='x', **pad)

        # 操作类型：3 行 3 列网格（无「操作类型:」标签，上方「处理选项」已提示）
        # 新增操作只需在 App.OPERATIONS / App.OP_RADIO_LAYOUT 登记
        frm_op_row = ttk.Frame(self.frm_opt)
        frm_op_row.pack(fill='x')
        op_labels = {oid: lbl for oid, lbl, _ in self.OPERATIONS}
        for oid, r, c in self.OP_RADIO_LAYOUT:
            ttk.Radiobutton(frm_op_row, text=op_labels[oid],
                            variable=self.operation, value=oid).grid(
                row=r, column=c, sticky='w', padx=4)
        # 列宽按内容自适应（不用 uniform 等宽，避免长描述被压缩截断）
        frm_op_row.columnconfigure((1, 2, 3), weight=1)

        # 可选选项行：所有操作的子选项叠放在同一容器格中，切换时只显示当前操作的子选项，
        # 始终位于同一行（不会上下错开）；各行统一高度，切换不跳动。
        # 新增选项：在 self.option_rows 中登记一行 build 函数即可。
        self.frm_opt_rows = ttk.Frame(self.frm_opt)
        self.frm_opt_rows.pack(fill='x', pady=(6, 0))
        self.frm_opt_rows.columnconfigure(0, weight=1)
        self.option_rows = {
            'lookup': _OptionRow(self.frm_opt_rows, self._build_lookup_options),
            'shortcode': _OptionRow(self.frm_opt_rows, self._build_shortcode_options),
            'fullcode': _OptionRow(self.frm_opt_rows, self._build_fullcode_options),
            'dupcode': _OptionRow(self.frm_opt_rows, self._build_dupcode_options),
            'wordextract': _OptionRow(self.frm_opt_rows, self._build_wordextract_options),
            'phrasecode': _OptionRow(self.frm_opt_rows, self._build_phrase_options),
        }
        # 所有可选行统一到同一高度，视觉一致且无跳动。
        # 先统一刷新一次几何布局，再测量各行真实高度（合并多次 update_idletasks 为一次，启动更快）
        self.frm_opt_rows.update_idletasks()
        for row in self.option_rows.values():
            row._measure()
        _max_row_h = max(row.h for row in self.option_rows.values())
        for row in self.option_rows.values():
            row.set_height(_max_row_h)

        # ---- 输出文件名（wordextract 模式下文案改为「输出文件夹名」）----
        self.frm_out = ttk.Frame(self.content, padding=8)
        frm_out = self.frm_out
        frm_out.pack(fill='x', **pad)
        self.lbl_out_name = ttk.Label(frm_out, text='输出文件名:')
        self.lbl_out_name.pack(side='left')
        ttk.Entry(frm_out, textvariable=self.output_name, width=64).pack(side='left', padx=8)
        ttk.Button(frm_out, text='浏览…', command=self._browse_output).pack(side='left')
        # 保存位置提示：无内容时整体隐藏，避免残留空行
        self.lbl_save_path = ttk.Label(self.content, textvariable=self.save_path_var,
                                       foreground='gray')

        # ---- 执行 ----
        self.frm_run = ttk.Frame(self.content)
        self.frm_run.pack(pady=8)
        ttk.Button(self.frm_run, text='开始处理', command=self._run).pack(side='left', padx=4)
        self.btn_open_output = ttk.Button(
            self.frm_run, text='打开输出位置', command=self._open_output_location,
            state='disabled')
        self.btn_open_output.pack(side='left', padx=4)

        # ---- 结果行数过滤 ----
        self.frm_filter = ttk.Frame(self.content)
        frm_filter = self.frm_filter
        frm_filter.pack(fill='x', padx=4, pady=(0, 4))
        vcmd = (self.root.register(self._validate_threshold), '%P')
        self.chk_preview_only = ttk.Checkbutton(
            frm_filter, text='结果行数少于', variable=self.preview_only_enabled,
            command=self._update_preview_only_state)
        self.chk_preview_only.pack(side='left')
        self.ent_threshold = ttk.Entry(
            frm_filter, textvariable=self.preview_only_threshold, width=7,
            validate='key', validatecommand=vcmd)
        self.ent_threshold.pack(side='left', padx=4)
        ttk.Label(frm_filter, text='行时，结果只输出到预览框中（不保存为文件）').pack(side='left')
        self.ent_threshold.bind('<FocusOut>', self._clamp_threshold)

        # ---- 处理结果预览（固定 6 行高，自带滚动条；整页滚动由外层画布负责）----
        self.frm_result = ttk.LabelFrame(self.content, text='处理结果预览', padding=8)
        frm_result = self.frm_result
        frm_result.pack(fill='x', **pad)
        self.lbl_result_lines = ttk.Label(frm_result, text='', foreground='gray')
        # 初始无结果，先隐藏，避免占位留白（有结果后由 _set_result_lines 显示）
        self.lbl_result_lines.pack_forget()
        self.preview_result = self._make_preview(frm_result, text_font, height=6)

        # 状态提示用 side='bottom' 从底部装载，确保始终可见
        # wraplength 让长路径自动换行，justify='left' 多行左对齐；
        # anchor='w' + fill='x' 让状态栏占满宽度，避免路径被截断
        self.status = ttk.Label(self.root, text='请先导入两个文件', foreground='gray',
                                wraplength=1380, justify='left', anchor='w')
        self.status.pack(side='bottom', pady=4, fill='x')

    def _set_preview(self, pv, content):
        """向一个 _Preview 写入内容（保留其 editable 状态）"""
        pv.set_content(content)

    def _set_result_lines(self, text):
        """设置结果行数提示；无内容时隐藏标签，避免残留空行"""
        self.lbl_result_lines.config(text=text)
        if text:
            self.lbl_result_lines.pack(anchor='w')
        else:
            self.lbl_result_lines.pack_forget()

    def _make_preview(self, parent, text_font, height=8, editable=False,
                      line_numbers=True, on_overlimit=None):
        """创建一个带行号栏的预览框（返回 _Preview 对象）

        - editable=True：可手动编辑（用于未导入文件时的待处理预览）
        - line_numbers=True：左侧显示行号
        - on_overlimit：超出 MAX_LINES 行时的回调（用于弹窗提示）
        """
        return _Preview(parent, text_font, height=height, editable=editable,
                        line_numbers=line_numbers, on_overlimit=on_overlimit)

    def _on_preview_overlimit(self, count):
        messagebox.showwarning(
            '提示',
            f'待处理预览最多支持 {_Preview.MAX_LINES:,} 行，\n'
            f'当前已输入 {count:,} 行，超出部分已自动截断。')

    @staticmethod
    def _fmt_lines(path, total):
        if not path:
            return ''
        return f'共 {total} 行' if total else '（空文件）'

    def _needs_reference(self, op=None, short_scope=None):
        """判断当前操作是否需要参考文件"""
        op = op or self.operation.get()
        if op in ('difference', 'intersection', 'lookup', 'wordextract', 'phrasecode'):
            return True
        if op == 'shortcode':
            # 处理范围 1-4 需要参考列表，5-7 忽略参考
            sc = short_scope if short_scope is not None else int(self.shortcode_scope.get()[0])
            return sc in (1, 2, 3, 4)
        # 提取全码 / 出重码号 无需参考文件
        return False

    def _update_path_wrap(self):
        """路径标签按当前父容器宽度自动换行，随窗口缩放调整"""
        for lbl in (self.lbl_target_path, self.lbl_ref_path):
            parent = lbl.master
            w = parent.winfo_width()
            if w > 20:
                lbl.config(wraplength=max(60, w - 18))

    def _update_status(self):
        t = self.target_path.get()
        r = self.ref_path.get()
        need_ref = self._needs_reference()
        # 简词提取的词频文件可选（不导入则不按词频排序，按原始顺序）
        if self.operation.get() == 'wordextract':
            need_ref = False
        # 已导入文件 或 已在预览框手动输入内容，均视为「已就绪」
        t_present = bool(t) or bool(self.preview_t.get_content().strip())
        r_present = bool(r) or bool(self.preview_r.get_content().strip())

        # 词组编码：左框=单字表、右框=待编码词组，给出更准确的就绪提示
        if self.operation.get() == 'phrasecode':
            if not t_present:
                self.status.config(
                    text='请先导入单字表（左框），或在左侧预览框中手动输入内容。',
                    foreground='gray')
            elif not r_present:
                self.status.config(
                    text='请导入待编码词组（右框），或在右侧预览框中手动输入内容。',
                    foreground='gray')
            else:
                self.status.config(text='已就绪，可点击「开始处理」。', foreground='green')
            return

        if not t_present:
            self.status.config(
                text='请先导入待处理文件，或在待处理预览框中手动输入内容。',
                foreground='gray')
        elif need_ref and not r_present:
            self.status.config(
                text='请导入参考文件，或在参考预览框中手动输入内容。', foreground='gray')
        else:
            self.status.config(text='已就绪，可点击「开始处理」。', foreground='green')

    def _update_save_label(self):
        t = self.target_path.get()
        if not t:
            # 手动输入模式：结果保存到当前工作目录
            if self.preview_t.get_content().strip():
                out = self.output_name.get().strip() or self.DEFAULT_OUTPUT_NAMES[self.operation.get()]
                p = out if os.path.isabs(out) else os.path.join(os.getcwd(), out)
                if not p.lower().endswith(('.txt', '.yaml')):
                    p += '.txt'
                self.save_path_var.set(f'保存位置：{p}（手动输入模式）')
            else:
                self.save_path_var.set('')
            # 有保存位置说明才显示标签，否则隐藏以避免残留空行
            if self.save_path_var.get():
                self.lbl_save_path.pack(anchor='w', padx=18, pady=(0, 4))
            else:
                self.lbl_save_path.pack_forget()
            return
        out = self.output_name.get().strip() or self.DEFAULT_OUTPUT_NAMES[self.operation.get()]
        if os.path.isabs(out):
            p = out
        else:
            p = os.path.join(os.path.dirname(t) or '.', out)
        if not p.lower().endswith(('.txt', '.yaml')):
            p += '.txt'
        self.save_path_var.set(f'保存位置：{p}')
        self.lbl_save_path.pack(anchor='w', padx=18, pady=(0, 4))

    def _browse_output(self):
        """浏览保存位置：wordextract 非合并选文件夹，合并模式和其他模式选文件"""
        op = self.operation.get()
        if op == 'wordextract' and not self.merge_results.get():
            # 简词提取非合并模式：输出为文件夹
            d = filedialog.askdirectory(title='选择输出文件夹')
            if d:
                self.output_name.set(d)
                self._update_save_label()
        else:
            # 其他模式或合并模式：输出为文件
            initial = self.output_name.get().strip() or '处理结果.txt'
            p = filedialog.asksaveasfilename(
                title='选择保存文件',
                initialfile=initial,
                filetypes=[('文本文件', '*.txt'), ('YAML 文件', '*.yaml'),
                           ('所有文件', '*.*')])
            if p:
                self.output_name.set(p)
                self._update_save_label()

    def _import_target(self):
        path = filedialog.askopenfilename(
            title='选择待处理文件',
            filetypes=[('文本/码表文件', '*.txt *.yaml'), ('所有文件', '*.*')])
        if not path:
            return
        self.target_path.set(path)
        # 新导入文件时重置编码为自动检测
        self.target_enc.set('auto')
        text, total = read_preview(path, 100, enc_choice='auto')
        self._set_preview(self.preview_t, text)
        self.preview_t.set_editable(False)  # 导入文件后，待处理预览转为只读
        self.lines_t.set(self._fmt_lines(path, total))
        # 自动填充输出路径：仅当当前文件名为空或属于任一默认名时
        if self._is_default_output_name():
            default_name = self.DEFAULT_OUTPUT_NAMES[self.operation.get()]
            self.output_name.set(os.path.join(os.path.dirname(path), default_name))
        self._update_status()
        self._update_save_label()

    def _clear_target(self):
        """清空待处理：回到手动输入模式（预览框可编辑、清空内容）"""
        self.target_path.set('')
        self.target_enc.set('auto')
        self.preview_t.set_content('')
        self.preview_t.set_editable(True)
        self.lines_t.set('')
        self._update_status()
        self._update_save_label()

    def _import_ref(self):
        path = filedialog.askopenfilename(
            title='选择参考文件',
            filetypes=[('文本/码表文件', '*.txt *.yaml'), ('所有文件', '*.*')])
        if not path:
            return
        self.ref_path.set(path)
        self.ref_enc.set('auto')
        self.preset_var.set('')
        text, total = read_preview(path, 100, enc_choice='auto')
        self._set_preview(self.preview_r, text)
        self.preview_r.set_editable(False)  # 导入文件后，参考预览转为只读
        self.lines_r.set(self._fmt_lines(path, total))
        self._update_status()

    def _clear_ref(self):
        """清空参考：回到手动输入模式（预览框可编辑、清空内容）"""
        self.ref_path.set('')
        self.ref_enc.set('auto')
        self.preset_var.set('')
        self.preview_r.set_content('')
        self.preview_r.set_editable(True)
        self.lines_r.set('')
        self._update_status()

    def _load_preset(self, name):
        """从程序文件夹\预设 载入选中预设文件内容到参考预览框（UTF-8，自动处理 BOM）"""
        if not name:
            return
        preset_path = os.path.join(self.preset_dir, name)
        if not os.path.exists(preset_path):
            messagebox.showwarning('提示', f'预设文件不存在:\n{preset_path}')
            return
        self.ref_path.set(preset_path)
        self.ref_enc.set('auto')
        text, total = read_preview(preset_path, 100, enc_choice='auto')
        self._set_preview(self.preview_r, text)
        self.preview_r.set_editable(False)  # 载入预设后，参考预览转为只读
        self.lines_r.set(self._fmt_lines(preset_path, total))
        self._update_status()

    def _is_default_output_name(self):
        """判断当前 output_name 是否为空或属于任一操作的默认名（仅比较文件名部分）"""
        cur = self.output_name.get().strip()
        if not cur:
            return True
        base = os.path.basename(cur)
        return base in self.DEFAULT_OUTPUT_NAMES.values()

    def _reload_target_preview(self):
        """使用当前选择的编码重新加载待处理文件预览"""
        path = self.target_path.get()
        if not path:
            return
        text, total = read_preview(path, 100, enc_choice=self.target_enc.get())
        self._set_preview(self.preview_t, text)
        self.lines_t.set(self._fmt_lines(path, total))

    def _reload_ref_preview(self):
        """使用当前选择的编码重新加载参考文件预览"""
        path = self.ref_path.get()
        if not path:
            return
        text, total = read_preview(path, 100, enc_choice=self.ref_enc.get())
        self._set_preview(self.preview_r, text)
        self.lines_r.set(self._fmt_lines(path, total))

    def _swap(self):
        """互换两个文件的路径、编码与预览"""
        t_path = self.target_path.get()
        r_path = self.ref_path.get()
        t_enc = self.target_enc.get()
        r_enc = self.ref_enc.get()
        self.target_path.set(r_path)
        self.ref_path.set(t_path)
        self.target_enc.set(r_enc)
        self.ref_enc.set(t_enc)
        text_t, total_t = read_preview(r_path, 100, enc_choice=r_enc) if r_path else ('', 0)
        text_r, total_r = read_preview(t_path, 100, enc_choice=t_enc) if t_path else ('', 0)
        self._set_preview(self.preview_t, text_t)
        self._set_preview(self.preview_r, text_r)
        self.lines_t.set(self._fmt_lines(r_path, total_t))
        self.lines_r.set(self._fmt_lines(t_path, total_r))
        self._update_status()

    def _on_operation_change(self):
        """根据操作类型更新「忽略第二列」复选框、未匹配值输入框、默认输出文件名

        - 上词频 / 出简不出全 / 提取全码 模式下「忽略第二列」固定禁用并取消勾选
          （上词频参考文件是 key->value 表；出简/提取全码需用到第二列编码）
        - 上词频模式显示未匹配值输入框与排序复选框；
          出简不出全模式显示出简模式与处理范围下拉框；
          提取全码模式显示处理范围下拉框（仅 全部词条/单字/词组 三项）
        - 切换操作类型时若用户未自定义文件名，则自动更新为该操作的默认名
        """
        op = self.operation.get()
        # 上词频模式下右侧按钮文本改为「导入词频文件」；词组编码复用标准文件区，
        # 右框改为「单字表」、左框改为「待编码词组」（见下方标签切换）；其余恢复「导入参考文件」
        if op == 'phrasecode':
            self.btn_import_ref.config(text='导入待编码词组')
        elif op in ('lookup', 'wordextract'):
            self.btn_import_ref.config(text='导入词频文件')
        else:
            self.btn_import_ref.config(text='导入参考文件')
        # 简词提取输出为文件夹（合并模式除外），标签文案相应调整
        if op == 'wordextract' and not self.merge_results.get():
            self.lbl_out_name.config(text='输出文件夹名:')
        else:
            self.lbl_out_name.config(text='输出文件名:')
        # 「结果行数过滤」：简词提取非合并模式（输出文件夹）不适用，禁用并取消勾选
        # 切换离开时还原进入前记忆的状态（与合并复选框对 extract_limit 的处理一致）
        if op == 'wordextract' and not self.merge_results.get():
            if not hasattr(self, '_saved_preview_only'):
                self._saved_preview_only = self.preview_only_enabled.get()
            self.preview_only_enabled.set(False)
            self.chk_preview_only.state(['disabled'])
            self.ent_threshold.state(['disabled'])
        else:
            # 还原记忆的状态（仅还原一次，避免用户在新页面的手动勾选被覆盖）
            if hasattr(self, '_saved_preview_only'):
                self.preview_only_enabled.set(self._saved_preview_only)
                del self._saved_preview_only
            self.chk_preview_only.state(['!disabled'])
            self._update_preview_only_state()
        # 忽略第二列：上词频 / 简词提取 / 出简不出全 / 提取全码 / 出重码号 不需要，固定禁用并取消勾选
        no_col_ops = {'lookup', 'wordextract', 'shortcode', 'fullcode', 'dupcode', 'phrasecode'}
        if op in no_col_ops:
            self.ignore_target.set(False)
            self.ignore_ref.set(False)
            self.chk_target.state(['disabled'])
            self.chk_ref.state(['disabled'])
        else:
            self.chk_target.state(['!disabled'])
            self.chk_ref.state(['!disabled'])
        # 显示 / 隐藏对应的可选选项行（其余行以等高占位保留，不引起页面上下移动）
        for key, row in self.option_rows.items():
            row.show(key == op)
        # 整体变灰 / 恢复正常（如出简处理范围 5-7 无需参考文件）
        self._update_reference_state()

        # 若用户未自定义输出文件名，则更新为当前操作对应的默认名
        if self._is_default_output_name() and op in self.DEFAULT_OUTPUT_NAMES:
            cur = self.output_name.get().strip()
            dir_part = os.path.dirname(cur) if cur else ''
            new_name = self.DEFAULT_OUTPUT_NAMES[op]
            if dir_part:
                self.output_name.set(os.path.join(dir_part, new_name))
            else:
                self.output_name.set(new_name)

        # 词组编码直接复用标准文件区（与简词提取一致：不新建/不替换文件区），
        # 仅按操作类型改标签：左框=待编码词组（主输入）、右框=单字表（取码字典）。
        if op == 'phrasecode':
            self.frm_t.config(text='单字表')
            self.frm_r.config(text='待编码词组')
            self.btn_import_target.config(text='导入单字表')
        else:
            self.frm_t.config(text='待处理文件')
            self.frm_r.config(text='参考文件')
            self.btn_import_target.config(text='导入待处理文件')

    def _build_lookup_options(self, parent):
        """上词频选项行内容：未匹配填充值 + 排序复选框"""
        ttk.Label(parent, text='未匹配时的填充值:').pack(side='left')
        self.ent_unmatch = ttk.Entry(parent, textvariable=self.unmatched_value,
                                     width=24)
        self.ent_unmatch.pack(side='left', padx=8)
        self.chk_sort_freq = ttk.Checkbutton(
            parent, text='按匹配值（数字）降序排序（非数字置后）',
            variable=self.sort_by_freq)
        self.chk_sort_freq.pack(side='left', padx=12)

    def _build_dupcode_options(self, parent):
        """出重码号选项行内容：序号是否用 Tab 分隔的复选框"""
        self.chk_dupcode_tab = ttk.Checkbutton(
            parent, text='序号用 Tab 分隔（不拼在编码后）',
            variable=self.dupcode_tab)
        self.chk_dupcode_tab.pack(side='left', padx=8)

    def _build_wordextract_options(self, parent):
        """简词提取选项行：左侧正则输入框(2行高)+生成器，右侧两行选项"""
        # 整体水平布局：左侧正则区跨2行，右侧选项分2行
        frm_main = ttk.Frame(parent)
        frm_main.pack(fill='x')
        frm_main.columnconfigure(0, weight=0)
        frm_main.columnconfigure(1, weight=1)

        # 左侧：正则 + 生成器（跨2行，与右侧两行选项等高）
        frm_left = ttk.Frame(frm_main)
        frm_left.grid(row=0, column=0, rowspan=2, sticky='ns', padx=(0, 12))
        # 「正则:」做成可点击标签（外观同普通标签，悬停变蓝、显手型光标）
        # 点击后弹出正则语法说明页
        lbl_regex = ttk.Label(frm_left, text='正则:', cursor='hand2')
        lbl_regex.pack(side='left')
        lbl_regex.bind('<Button-1>', lambda e: self._show_regex_help())
        lbl_regex.bind('<Enter>', lambda e: lbl_regex.config(foreground='#0066cc'))
        lbl_regex.bind('<Leave>', lambda e: lbl_regex.config(foreground=''))
        self.txt_regex = _Preview(frm_left, (CJK_FONT, 9), height=3, width=24,
                                  editable=True, line_numbers=True)
        self.txt_regex.frame.pack_forget()
        self.txt_regex.frame.pack(side='left', fill='y', padx=4)
        ttk.Button(frm_left, text='生成器', width=5,
                   command=self._show_regex_generator).pack(side='left', padx=4)

        # 右侧第1行：仅提取前N行
        frm_r1 = ttk.Frame(frm_main)
        frm_r1.grid(row=0, column=1, sticky='w')
        self.chk_extract_limit = ttk.Checkbutton(
            frm_r1, text='仅提取前', variable=self.extract_limit_enabled,
            command=self._update_extract_limit_state)
        self.chk_extract_limit.pack(side='left')
        self.ent_extract_limit = ttk.Entry(
            frm_r1, textvariable=self.extract_limit, width=7)
        self.ent_extract_limit.pack(side='left', padx=4)
        ttk.Label(frm_r1, text='行').pack(side='left')

        # 右侧第2行：合并 + 标题行
        frm_r2 = ttk.Frame(frm_main)
        frm_r2.grid(row=1, column=1, sticky='w')
        self.chk_merge = ttk.Checkbutton(
            frm_r2, text='结果合并为一个文件',
            variable=self.merge_results,
            command=self._on_merge_change)
        self.chk_merge.pack(side='left', padx=8)
        self.chk_regex_header = ttk.Checkbutton(
            frm_r2, text='结果包含正则标题行',
            variable=self.include_regex_header)
        self.chk_regex_header.pack(side='left', padx=8)

    def _update_extract_limit_state(self):
        """仅提取前N行：不勾选时输入框置灰"""
        self.ent_extract_limit.state(
            ['!disabled'] if self.extract_limit_enabled.get() else ['disabled'])

    def _on_merge_change(self):
        """勾选「结果合并为一个文件」时：输出变文件、记忆并设提取前1行；取消时还原"""
        if self.merge_results.get():
            # 记住当前提取前N行的值，设为1
            self._saved_extract_limit = self.extract_limit.get()
            self.extract_limit_enabled.set(True)
            self.extract_limit.set('1')
            self._update_extract_limit_state()
            # 输出改为文件：若无 .txt/.yaml 后缀则补上
            name = self.output_name.get().strip()
            if name and not name.lower().endswith(('.txt', '.yaml')):
                self.output_name.set(name + '.txt')
            self.lbl_out_name.config(text='输出文件名:')
            # 合并模式输出单文件，自动勾选并启用「结果行数过滤」
            self.preview_only_enabled.set(True)
            self.chk_preview_only.state(['!disabled'])
            self._update_preview_only_state()
        else:
            # 还原提取前N行的值
            if hasattr(self, '_saved_extract_limit'):
                self.extract_limit.set(self._saved_extract_limit)
            # 输出改回文件夹：去掉 .txt/.yaml 后缀
            name = self.output_name.get().strip()
            if name.lower().endswith(('.txt', '.yaml')):
                self.output_name.set(os.path.splitext(name)[0])
            self.lbl_out_name.config(text='输出文件夹名:')
            # 非合并模式输出文件夹，禁用并取消「结果行数过滤」
            self.preview_only_enabled.set(False)
            self.chk_preview_only.state(['disabled'])
            self.ent_threshold.state(['disabled'])
        self._update_save_label()

    def _show_regex_help(self):
        """弹出正则语法说明窗口（模态，grab_set 抢占焦点）

        展示简词提取所支持的正则语法及其意义，包含：
        - 基础元字符、量词、字符类（Python re 模块）
        - 简词提取常用模式（四码方案，与生成器一致）
        - 匹配示例
        """
        win = tk.Toplevel(self.root)
        win.title('正则语法说明')
        win.transient(self.root)
        win.grab_set()
        win.resizable(True, True)

        # 内容文本（等宽字体对齐）
        content = [
            '简词提取使用 Python re 模块的 fullmatch 匹配整行，',
            '匹配格式为「词组\\t编码」（\\t 为 Tab 制表符）。',
            '',
            '  .      任意单个字符（不含换行）',
            '  ^      字符串开头',
            '  $      字符串结尾',
            '  \\t     Tab 制表符（分隔词组与编码）',
            '  \\d     数字 [0-9]',
            '  \\D     非数字',
            '  \\s     空白字符',
            '  \\S     非空白字符',
            '  a|b     a或b',
            '  *      前一项出现 0 次或多次',
            '  +      前一项出现 1 次或多次',
            '  ?      前一项出现 0 次或 1 次',
            '  {n}    前一项出现恰好 n 次',
            '  {n,m}  前一项出现 n 到 m 次',
            '  [abc]   a / b / c 中任意一个',
            '  [^abc]  非 a / b / c 的任意字符',
            '  [a-z]   a 到 z 中任意一个字符',
            '',
            '【详细介绍请查看】',
            'Python正则文档：https://docs.python.org/zh-cn/3/library/re.html',
            '标准正则教程：https://www.runoob.com/regexp/regexp-tutorial.html',
        ]

        frm = ttk.Frame(win, padding=8)
        frm.pack(fill='both', expand=True)

        txt = tk.Text(frm, wrap='none', font=(CJK_FONT, 10),
                      bg='#fafafa', relief='flat', padx=8, pady=6,
                      spacing1=0, spacing2=0, spacing3=2)
        txt.pack(side='left', fill='both', expand=True)
        _init_preview_scrollbar_style()
        sb = ttk.Scrollbar(frm, orient='vertical', command=txt.yview,
                           style='Preview.Vertical.TScrollbar')
        sb.pack(side='right', fill='y')
        txt.config(yscrollcommand=sb.set, state='normal')
        txt.insert('1.0', '\n'.join(content))
        txt.config(state='disabled')  # 只读

        # 关闭按钮
        ttk.Button(win, text='关闭', command=win.destroy).pack(side='bottom', pady=6)

        # 居中于主窗口
        win.update_idletasks()
        w, h = 960, 560
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        win.geometry(f'{w}x{h}+{max(0, x)}+{max(0, y)}')

    def _show_regex_generator(self):
        """弹出正则生成器窗口（模态，grab_set 抢占焦点，后方不可点击）

        根据词组长度、模式、输入码生成四码方案正则，追加到主界面正则输入框。
        左侧参数区 + 右侧输入框（带行号），每行输入超过限定字符自动换行。
        """
        win = tk.Toplevel(self.root)
        win.title('正则生成器')
        win.transient(self.root)
        win.grab_set()  # 模态：后方窗口不可点击
        win.resizable(True, True)

        # 主体：左侧参数区 + 右侧输入区（水平布局）
        frm_body = ttk.Frame(win, padding=8)
        frm_body.pack(fill='both', expand=True)

        # ---- 左侧参数区 ----
        frm_left = ttk.LabelFrame(frm_body, text='参数', padding=8)
        frm_left.pack(side='left', fill='y', padx=(0, 6))

        ttk.Label(frm_left, text='词组长度:').pack(anchor='w')
        cmb_length = ttk.Combobox(frm_left, width=6, state='readonly',
                                  values=['2', '3'])
        cmb_length.set('2')
        cmb_length.pack(anchor='w', pady=(2, 8))

        ttk.Label(frm_left, text='模式:').pack(anchor='w')
        cmb_mode = ttk.Combobox(
            frm_left, width=22, state='readonly',
            values=['简拼式（取每字首字母）', '联想式（取第一字前两码）'])
        cmb_mode.set('简拼式（取每字首字母）')
        cmb_mode.pack(anchor='w', pady=(2, 8))

        lbl_hint = ttk.Label(frm_left, text='每行最多 2 个字符',
                             foreground='gray')
        lbl_hint.pack(anchor='w')

        # 按钮放左侧底部
        ttk.Button(frm_left, text='生成并追加',
                   command=lambda: self._apply_generated_regex(
                       txt_input.get('1.0', 'end-1c'), cmb_length.get(),
                       cmb_mode.get(), win)).pack(side='bottom', fill='x', pady=(8, 2))
        ttk.Button(frm_left, text='关闭',
                   command=win.destroy).pack(side='bottom', fill='x')

        # ---- 右侧输入区（带行号）----
        frm_right = ttk.LabelFrame(frm_body, text='输入码（每行一个）', padding=4)
        frm_right.pack(side='left', fill='both', expand=True)

        # 行号栏
        txt_lines = tk.Text(frm_right, width=4, height=20, padx=3,
                            state='disabled', bg='#f0f0f0', takefocus=0,
                            font=(CJK_FONT, 10))
        txt_lines.pack(side='left', fill='y')
        # 输入框
        txt_input = tk.Text(frm_right, height=20, width=12, wrap='none',
                            font=(CJK_FONT, 10), undo=True)
        txt_input.pack(side='left', fill='both', expand=True)
        # 右侧滚动条（复用预览框同款样式）
        _init_preview_scrollbar_style()
        sb_input = ttk.Scrollbar(frm_right, orient='vertical',
                                 command=txt_input.yview,
                                 style='Preview.Vertical.TScrollbar')
        sb_input.pack(side='right', fill='y')
        txt_input.config(yscrollcommand=lambda *a: (sb_input.set(*a), txt_lines.yview_moveto(a[0])))

        def update_line_numbers(*_args):
            """根据输入框行数刷新左侧行号"""
            txt_lines.config(state='normal')
            txt_lines.delete('1.0', 'end')
            n = int(txt_input.index('end-1c').split('.')[0])
            txt_lines.insert('1.0', '\n'.join(str(i) for i in range(1, n + 1)))
            txt_lines.config(state='disabled')

        def auto_wrap(_evt=None):
            """每行超过限定字符自动换行（插入换行符），光标跟随下移到新行"""
            max_len = 2 if cmb_length.get() == '2' else 3
            cursor = txt_input.index('insert')
            cur_line = int(cursor.split('.')[0])
            cur_col = int(cursor.split('.')[1])
            line_text = txt_input.get(f'{cur_line}.0', f'{cur_line}.end')

            # 光标行超长：在 max_len 处插入换行符，光标移到新行
            if len(line_text) > max_len:
                txt_input.insert(f'{cur_line}.{max_len}', '\n')
                new_col = cur_col - max_len
                if new_col < 0:
                    new_col = 0
                txt_input.mark_set('insert', f'{cur_line + 1}.{new_col}')

            # 全量检查其他行（粘贴等情况）：超长切分，光标不跟随
            content = txt_input.get('1.0', 'end-1c')
            lines = content.split('\n')
            fixed = []
            for line in lines:
                if len(line) > max_len:
                    for i in range(0, len(line), max_len):
                        fixed.append(line[i:i + max_len])
                else:
                    fixed.append(line)
            new_content = '\n'.join(fixed)
            if new_content != content:
                txt_input.delete('1.0', 'end')
                txt_input.insert('1.0', new_content)
                try:
                    txt_input.mark_set('insert',
                                       f'{cur_line + 1}.{new_col}' if len(line_text) > max_len else cursor)
                except Exception:
                    txt_input.mark_set('insert', 'end')
            update_line_numbers()

        txt_input.bind('<KeyRelease>', auto_wrap)
        txt_input.bind('<<Paste>>', lambda e: win.after(10, auto_wrap))
        update_line_numbers()

        # 长度变化时联动：模式框禁用/启用 + 提示文字
        def _on_len_change(_evt=None):
            if cmb_length.get() == '2':
                cmb_mode.config(state='readonly')
                lbl_hint.config(text='每行最多 2 个字母')
            else:
                # 长度=3 无模式区分：禁用模式框并显示「简拼式」
                cmb_mode.set('简拼式（取每字首字母）')
                cmb_mode.config(state='disabled')
                lbl_hint.config(text='每行最多 3 个字母')
            auto_wrap()  # 长度变化后重新切分已有内容
        cmb_length.bind('<<ComboboxSelected>>', _on_len_change)

        # 底部提示
        ttk.Label(win, text='生成器只支持四码方案',
                  foreground='gray').pack(side='bottom', pady=4)

        # 居中于主窗口
        win.update_idletasks()
        w, h = 580, 460
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        win.geometry(f'{w}x{h}+{max(0, x)}+{max(0, y)}')

        # 焦点放到输入框，弹窗后可直接输入
        txt_input.focus_set()

    def _apply_generated_regex(self, codes_text, length_str, mode_str, win):
        """根据生成器输入生成正则列表，追加到主界面的正则输入框

        四码方案正则生成逻辑：
        - 词组前缀：长度2='..', 长度3='...'
        - 编码部分固定4字符
        - 简拼式（长度2）：输入字符交错 X.X. （每字首字母+任意第二码）
        - 联想式 / 长度3：输入字符放前，'.'补齐到4
        生成结果去重（保持首次出现顺序），追加时跳过正则框中已有的正则。
        """
        length = int(length_str)
        prefix = '.' * length
        max_len = 2 if length == 2 else 3
        is_simpinyin = (length == 2 and '简拼' in mode_str)
        regexes = []
        seen = set()
        for line in codes_text.split('\n'):
            code = line.strip()
            if not code:
                continue
            code = code[:max_len]  # 限制输入长度
            if is_simpinyin:
                # 简拼式：每字首字母 + 任意第二码 → X.X.
                enc = ''
                for i in range(2):
                    enc += (code[i] + '.') if i < len(code) else '..'
            else:
                # 联想式 / 长度3：输入字符放前，'.'补齐到4
                enc = code + '.' * (4 - len(code))
            regex = f'^{prefix}\\t{enc}$'
            if regex not in seen:  # 生成结果去重
                seen.add(regex)
                regexes.append(regex)

        if regexes:
            current = self.txt_regex.text.get('1.0', 'end-1c')
            existing = {l.strip() for l in current.split('\n') if l.strip()}
            new_lines = [r for r in regexes if r not in existing]  # 跳过已有
            if new_lines:
                if current.strip():
                    self.txt_regex.text.insert('end', '\n')
                self.txt_regex.text.insert('end', '\n'.join(new_lines))
                # 插入后手动刷新行号（insert 不触发 _on_change）
                self.txt_regex._update_line_numbers()
        win.destroy()

    def _build_fullcode_options(self, parent):
        """提取全码选项行内容：处理范围下拉框（仅 全部词条/单字/词组 三项）"""
        ttk.Label(parent, text='处理范围:').pack(side='left')
        self.cmb_full_scope = ttk.Combobox(
            parent, textvariable=self.fullcode_scope, width=36, state='readonly',
            values=[
                '1. 仅处理单字',
                '2. 仅处理词组',
                '3. 处理全部词条',
            ])
        self.cmb_full_scope.pack(side='left', padx=8)
        ttk.Label(parent, text='（保留最长编码，同为最长的均保留）',
                  foreground='gray').pack(side='left')

    def _build_shortcode_options(self, parent):
        """出简不出全选项行内容：出简模式下拉框 + 说明按钮 + 处理范围下拉框"""
        ttk.Label(parent, text='出简模式:').pack(side='left')
        self.cmb_short_rule = ttk.Combobox(
            parent, textvariable=self.shortcode_rule, width=34, state='readonly',
            values=[
                '1. 仅通过编码长度判断简码',
                '2. 通过前部编码是否相同逐步判断简码',
                '3. 通过编码是否完全包含判断简码',
            ])
        self.cmb_short_rule.pack(side='left', padx=(8, 2))
        ttk.Button(parent, text='说明', width=5,
                   command=self._show_shortcode_help).pack(side='left', padx=(0, 12))
        ttk.Label(parent, text='处理范围:').pack(side='left')
        self.cmb_short_scope = ttk.Combobox(
            parent, textvariable=self.shortcode_scope, width=36, state='readonly',
            values=[
                '1. 删除参考列表中全部词条全码',
                '2. 保留参考列表中全部词条全码',
                '3. 删除参考列表中单字词条全码',
                '4. 保留参考列表中单字全码（不处理词组）',
                '5. 仅处理单字（忽略参考）',
                '6. 仅处理词组（忽略参考）',
                '7. 处理全部词条（忽略参考）',
            ])
        self.cmb_short_scope.pack(side='left', padx=8)

    SHORTCODE_HELP = (
        '出简不出全判断简码规则说明\n'
        '词条格式：词条\\t编码。以下示例以「词条 编码」展示：\n'
        '\n'
        '【1. 仅通过编码长度判断简码】\n'
        '同一词条只保留编码最短的行（同为最短长度的都保留），\n'
        '其余更长编码全部视为全码删除。\n'
        '  一 g       （简码，保留）\n'
        '  一 gg      （非简码，删除）\n'
        '  一 ggll    （非简码，删除）\n'
        '  一 yi      （非简码，删除）\n'
        '  一个 gw    （简码，保留）\n'
        '  一个 ggwh  （非简码，删除）\n'
        '  一个 yg    （简码，保留，与 gw 同为最短）\n'
        '  一个 yige  （非简码，删除）\n'
        '\n'
        '【2. 通过前部编码是否相同逐步判断简码】\n'
        '按行序逐条判断：若之前已保留的某简码是当前编码的前缀，\n'
        '则当前编码视为全码删除；否则保留当前编码（并删除之前\n'
        '保留的以当前编码为前缀的更长编码）。\n'
        '  一 g       （简码，保留）\n'
        '  一 gg      （g 是其前缀 → 全码，删除）\n'
        '  一 ggll    （g 是其前缀 → 全码，删除）\n'
        '  一 yi      （无已保留前缀 → 简码，保留）\n'
        '  一个 gw    （简码，保留）\n'
        '  一个 ggwh  （gw 不是其前缀 → 简码，保留）\n'
        '  一个 yg    （简码，保留）\n'
        '  一个 yige  （yg 不是其前缀 → 简码，保留）\n'
        '\n'
        '【3. 通过编码是否完全包含判断简码】\n'
        '按行序逐条判断：若之前已保留的某简码的各字符按顺序完全\n'
        '出现在当前编码中，则当前编码视为全码删除；否则保留当前\n'
        '编码（并删除之前保留的完全包含当前编码的更长编码）。\n'
        '  一 g       （简码，保留）\n'
        '  一 gg      （包含 g → 全码，删除）\n'
        '  一 ggll    （包含 g → 全码，删除）\n'
        '  一 yi      （不包含 g → 简码，保留）\n'
        '  一个 gw    （简码，保留）\n'
        '  一个 ggwh  （按顺序包含 g、w → 全码，删除）\n'
        '  一个 yg    （不按顺序包含 gw → 简码，保留）\n'
        '  一个 yige  （按顺序包含 y、g → 全码，删除）\n'
        '\n'
        '处理范围：仅范围内的词条参与上述判断，范围外的词条原样\n'
        '保留。范围 1-4 需要参考列表（每行一个词条，若含制表符则\n'
        '取第一列），范围 5-7 忽略参考列表。'
    )

    def _show_shortcode_help(self):
        """弹出出简不出全的简码判定规则说明窗口（可滚动、可复制）"""
        win = tk.Toplevel(self.root)
        win.title('出简不出全 - 简码判定规则说明')
        win.transient(self.root)
        frm = ttk.Frame(win, padding=10)
        frm.pack(fill='both', expand=True)
        txt = tk.Text(frm, wrap='word', width=62, height=30,
                      font=(CJK_FONT, 10), padx=8, pady=6)
        sb = ttk.Scrollbar(frm, orient='vertical', command=txt.yview)
        txt.config(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        txt.pack(side='left', fill='both', expand=True)
        txt.insert('1.0', self.SHORTCODE_HELP)
        txt.config(state='disabled')
        ttk.Button(win, text='关闭', command=win.destroy).pack(pady=6)
        win.update_idletasks()
        # 居中于主窗口
        x = self.root.winfo_x() + (self.root.winfo_width() - win.winfo_reqwidth()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - win.winfo_reqheight()) // 2
        win.geometry(f'+{max(0, x)}+{max(0, y)}')

    def _show_about(self):
        """弹出关于窗口（内容暂为空，日后自行填写）

        窗口已预留内容容器 frm_body，直接在其中添加控件即可。
        若窗口已存在则聚焦到已有窗口，不重复创建。
        """
        # 若窗口已存在且未销毁，聚焦到它，不重复创建
        if self._about_win is not None:
            try:
                self._about_win.state()  # 触发 TclError 检查窗口是否仍存在
                self._about_win.lift()
                self._about_win.focus_force()
                return
            except tk.TclError:
                # 窗口已被销毁，重置引用
                self._about_win = None

        win = tk.Toplevel(self.root)
        self._about_win = win  # 保存引用，供重复点击时判断
        win.title('关于')
        win.transient(self.root)
        win.resizable(True, True)
        # 窗口关闭时清空引用（X 按钮与「关闭」按钮均走此回调）
        win.protocol('WM_DELETE_WINDOW', self._close_about)
        frm = ttk.Frame(win, padding=20)
        frm.pack(fill='both', expand=True)
        # 关于内容容器
        frm_body = ttk.Frame(frm)
        frm_body.pack(fill='both', expand=True)

        # GitHub 仓库地址（点击跳转浏览器）
        REPO_URL = "https://github.com/Unyaa-Code/DictProcessor"

        def open_repo(_event=None):
            webbrowser.open_new(REPO_URL)

        ttk.Label(frm_body, text='词库处理工具', font=('Microsoft YaHei', 14, 'bold')).pack(pady=(0, 12))
        ttk.Label(frm_body, text='作者：mono、铁圈、dsqm', font=('Microsoft YaHei', 10)).pack(anchor='w', pady=2)

        # GitHub 仓库：可点击的链接样式（蓝色 + 下划线 + 手型光标）
        row_repo = ttk.Frame(frm_body)
        row_repo.pack(anchor='w', pady=2)
        ttk.Label(row_repo, text='GitHub仓库：', font=('Microsoft YaHei', 10)).pack(side='left')
        link = tk.Label(row_repo, text=REPO_URL, font=('Microsoft YaHei', 10),
                        fg='#1a66ff', cursor='hand2')
        link.pack(side='left')
        link.bind('<Button-1>', open_repo)
        # 悬停时加下划线，离开时取消，模拟网页链接
        link.bind('<Enter>', lambda e: link.configure(font=('Microsoft YaHei', 10, 'underline')))
        link.bind('<Leave>', lambda e: link.configure(font=('Microsoft YaHei', 10)))

        ttk.Label(frm_body, text=f'当前版本：{APP_VERSION}', font=('Microsoft YaHei', 10)).pack(anchor='w', pady=(2, 0))

        # 底部图标来源标注（可点击跳转 icons8）
        row_icon = ttk.Frame(frm)
        row_icon.pack(side='bottom', fill='x', pady=(8, 0))
        link_icon = tk.Label(row_icon, text='Icons by icons8.com',
                             font=('Microsoft YaHei', 9), fg='#888888', cursor='hand2')
        link_icon.pack(side='right')
        link_icon.bind('<Button-1>', lambda e: webbrowser.open_new('https://icons8.com'))
        link_icon.bind('<Enter>', lambda e: link_icon.configure(fg='#1a66ff'))
        link_icon.bind('<Leave>', lambda e: link_icon.configure(fg='#888888'))

        ttk.Button(frm, text='关闭', command=self._close_about).pack(side='bottom', pady=(10, 0))
        win.update_idletasks()
        win.geometry('420x320')
        # 居中于主窗口
        x = self.root.winfo_x() + (self.root.winfo_width() - win.winfo_reqwidth()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - win.winfo_reqheight()) // 2
        win.geometry(f'+{max(0, x)}+{max(0, y)}')

    def _close_about(self):
        """关闭关于窗口并清空引用，避免下次点击误判窗口仍存在"""
        if self._about_win is not None:
            try:
                self._about_win.destroy()
            except tk.TclError:
                pass
            self._about_win = None

    def _update_reference_state(self):
        """当前操作不需要参考文件时（出简处理范围 5-7），将参考文件区整体置灰禁用"""
        need = self._needs_reference()
        grey = '#808080'
        if need:
            self.btn_import_ref.state(['!disabled'])
            self.btn_clear_ref.state(['!disabled'])
            # 上词频/简词提取/出简/提取全码/出重码号 不需要忽略第二列，保持禁用
            # （由 _on_operation_change 设置，此处不覆盖）
            if self.operation.get() not in {'lookup', 'wordextract', 'shortcode', 'fullcode', 'dupcode'}:
                self.chk_ref.state(['!disabled'])
            self.cmb_enc_r.state(['!disabled'])
            self.cmb_preset.state(['!disabled'])
            try:
                self.frm_r.config(style='TLabelframe')
            except Exception:
                pass
            # 词组编码下右框是「待编码词组」，不要被这里重置为「参考文件」
            if self.operation.get() != 'phrasecode':
                self.frm_r.config(text='参考文件')
            self.lbl_ref_path.config(foreground='#0000ff')
            self.lbl_ref_lines.config(foreground=grey)
            self.preview_r.text.config(fg='#000000', bg='#ffffff')
            # 未导入参考文件时，预览框可手动编辑；导入后转只读
            self.preview_r.set_editable(not bool(self.ref_path.get()))
        else:
            self.btn_import_ref.state(['disabled'])
            self.btn_clear_ref.state(['disabled'])
            self.chk_ref.state(['disabled'])
            self.cmb_enc_r.state(['disabled'])
            self.cmb_preset.state(['disabled'])
            try:
                self.frm_r.config(style='RefGrey.TLabelFrame')
            except Exception:
                pass
            self.frm_r.config(text='参考文件（本模式无需参考）')
            self.lbl_ref_path.config(foreground=grey)
            self.lbl_ref_lines.config(foreground=grey)
            self.preview_r.text.config(fg=grey, bg='#f0f0f0')
            self.preview_r.set_editable(False)

    # ------------------------------------------------------------------
    # 结果行数过滤校验
    # ------------------------------------------------------------------
    def _validate_threshold(self, value):
        """校验行数阈值输入：仅允许数字，最大 10000"""
        if value == '':
            return True
        if not value.isdigit():
            return False
        return int(value) <= 10000

    def _update_preview_only_state(self):
        """结果行数过滤：不勾选时输入框置灰"""
        self.ent_threshold.state(
            ['!disabled'] if self.preview_only_enabled.get() else ['disabled'])

    def _clamp_threshold(self, event=None):
        """焦点离开时修正阈值：空值设为 100，超出 10000 截断至 10000，小于 1 设为 1"""
        val = self.preview_only_threshold.get()
        if val == '' or not val.isdigit():
            self.preview_only_threshold.set('100')
            return
        v = int(val)
        if v > 10000:
            self.preview_only_threshold.set('10000')
        elif v < 1:
            self.preview_only_threshold.set('1')

    def _open_output_location(self):
        """在系统默认文件管理器中打开并选中最近一次输出的文件

        优先调用跨平台的 open_folder_and_select；失败时按平台回退到命令行方式。
        """
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            messagebox.showwarning('提示', '输出文件不存在，请先执行处理。')
            return
        path = os.path.normpath(self.last_output_path)
        try:
            if open_folder_and_select(path):
                return
            # 回退方案：按平台选择命令
            if sys.platform == 'win32':
                subprocess.Popen(f'explorer /select,"{path}"', shell=True)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', '-R', path])
            else:
                subprocess.Popen(['xdg-open', os.path.dirname(path)])
        except Exception as e:
            messagebox.showerror('错误', f'打开文件管理器失败: {e}')

    def _progress(self, msg):
        """处理过程中的进度回调（每 10 万行触发一次）"""
        self.status.config(text=msg, foreground='black')
        self.root.update_idletasks()

    def _run(self):
        target_path = self.target_path.get()
        ref_path = self.ref_path.get()
        # 出简不出全的处理范围 5-7 不需要参考文件，其余操作均需要
        need_ref = self._needs_reference()
        # 简词提取的词频文件可选（不导入则不按词频排序，按原始顺序）
        op = self.operation.get()
        if op == 'wordextract':
            need_ref = False
        # 手动输入模式：未导入文件，内容来自对应的预览框（已限 1 万行，直接在内存中处理）
        manual = not target_path
        if manual:
            content = self.preview_t.get_content()
            if not content.strip():
                messagebox.showwarning(
                    '提示', '请在「待处理」预览框中手动输入内容，或导入待处理文件。')
                return
            # 手动内容已限 1 万行，直接在内存中处理，无需落盘临时文件
            target_path = _manual_lines(content)
            target_enc = 'utf-8'
        else:
            if not target_path:
                messagebox.showwarning('提示', '请先导入待处理文件。')
                return
            target_enc = self.target_enc.get()

        if not ref_path:
            if need_ref:
                r_content = self.preview_r.get_content()
                if not r_content.strip():
                    messagebox.showwarning(
                        '提示', '请在「参考」预览框中手动输入内容，或导入参考文件。')
                    return
                # 手动内容已限 1 万行，直接在内存中处理，无需落盘临时文件
                ref_path = _manual_lines(r_content)
                ref_enc = 'utf-8'
            else:
                ref_enc = self.ref_enc.get()  # 本模式无需参考，不会被使用
        else:
            ref_enc = self.ref_enc.get()

        # 简词提取输出为文件夹（多文件），走独立流程，不适用常规单文件/缓冲逻辑
        if op == 'wordextract':
            self._run_wordextract(target_path, ref_path, target_enc, ref_enc, manual)
            return

        # 词组编码走独立流程（读取单字表/词组预览，输出到文件并复用标准结果预览）
        if op == 'phrasecode':
            self._run_phrasecode()
            return

        # 手动模式：结果直接生成在内存并呈现到预览，不写文件；
        # 文件模式：结果写入输出文件。
        # 「结果行数过滤」启用时，文件模式也先缓冲到内存，按行数决定是否写盘。
        use_buffer = manual or self.preview_only_enabled.get()
        if use_buffer:
            result_buf = io.StringIO()
        if manual:
            out_path = None
            dest_note = '结果已直接在预览中呈现（上限 1 万行）'
        else:
            out = self.output_name.get().strip() or self.DEFAULT_OUTPUT_NAMES[self.operation.get()]
            if os.path.isabs(out):
                out_path = out
            else:
                save_dir = os.path.dirname(target_path) or '.'
                out_path = os.path.join(save_dir, out)
            if not out_path.lower().endswith(('.txt', '.yaml')):
                out_path += '.txt'
            if use_buffer:
                dest_note = ''  # 缓冲模式下暂不确定输出路径
            else:
                out_path = get_unique_output_path(out_path)
                dest_note = f'结果已保存至: {out_path}'

        # 处理前清空上一次的结果预览
        self.preview_result.set_editable(False)
        self._set_preview(self.preview_result, '')
        self._set_result_lines('（处理中…）')

        if op == 'lookup':
            op_name = '上词频'
            self.status.config(text=f'正在执行{op_name}处理...', foreground='black')
            self.root.update_idletasks()
            unmatched = self.unmatched_value.get()
            result = process_lookup(target_path, ref_path, out_path, unmatched_value=unmatched,
                                    progress_callback=self._progress,
                                    target_enc=target_enc,
                                    ref_enc=ref_enc,
                                    sort_by_freq=self.sort_by_freq.get(),
                                    out_stream=result_buf if use_buffer else None)
            if result is None:
                return
            total, matched, duplicates = result
            # 第一行：统计信息；第二行：路径（单独一行避免长路径被截断）
            line1 = (f'{op_name}完成！共处理 {total:,} 行，匹配成功 {matched:,} 行，'
                     f'失败 {total - matched:,} 行。')
            if self.sort_by_freq.get():
                line1 += ' 已按匹配值降序排序。'
            if duplicates:
                line1 += f' 参考文件有 {len(duplicates)} 处重复 Key（已保留首次值）。'
            msg = f'{line1}\n{dest_note}'
            self.status.config(text=msg, foreground='green')
        elif op == 'shortcode':
            op_name = '出简不出全'
            rule = int(self.shortcode_rule.get()[0])
            scope = int(self.shortcode_scope.get()[0])
            self.status.config(
                text=f'正在执行{op_name}处理（模式{rule}，范围{scope}）...',
                foreground='black')
            self.root.update_idletasks()
            result = process_shortcode(
                target_path, ref_path, out_path, rule, scope,
                progress_callback=self._progress,
                target_enc=target_enc,
                ref_enc=ref_enc,
                out_stream=result_buf if use_buffer else None)
            if result is None:
                return
            written, conflicts = result
            line1 = f'{op_name}完成！共写出 {written:,} 行（模式{rule}，范围{scope}）。'
            if conflicts:
                line1 += f' 有 {len(conflicts)} 个词条存在相同长度的最短编码（均已保留）。'
            msg = f'{line1}\n{dest_note}'
            self.status.config(text=msg, foreground='green')
        elif op == 'fullcode':
            op_name = '提取全码'
            scope = int(self.fullcode_scope.get()[0])
            self.status.config(
                text=f'正在执行{op_name}处理（范围{scope}）...',
                foreground='black')
            self.root.update_idletasks()
            result = process_fullcode(
                target_path, out_path, scope,
                progress_callback=self._progress,
                target_enc=target_enc,
                out_stream=result_buf if use_buffer else None)
            if result is None:
                return
            written = result
            msg = f'{op_name}完成！共写出 {written:,} 行（范围{scope}）。\n{dest_note}'
            self.status.config(text=msg, foreground='green')
        elif op == 'dupcode':
            op_name = '出重码号'
            self.status.config(text=f'正在执行{op_name}处理（重复编码加后缀）...',
                               foreground='black')
            self.root.update_idletasks()
            result = process_dupcode(
                target_path, out_path,
                progress_callback=self._progress,
                target_enc=target_enc,
                tab_separated=self.dupcode_tab.get(),
                out_stream=result_buf if use_buffer else None)
            if result is None:
                return
            written = result
            msg = f'{op_name}完成！共写出 {written:,} 行。\n{dest_note}'
            self.status.config(text=msg, foreground='green')
        else:
            op_name = '差集' if op == 'difference' else '交集'
            self.status.config(text=f'正在执行{op_name}处理...', foreground='black')
            self.root.update_idletasks()
            count = process(target_path, ref_path, self.ignore_target.get(), self.ignore_ref.get(),
                            op, out_path, progress_callback=self._progress,
                            target_enc=target_enc, ref_enc=ref_enc,
                            out_stream=result_buf if use_buffer else None)
            if count is None:
                return
            # 完成提示分两行：统计行 + 路径行
            msg = f'{op_name}完成！保存 {count:,} 行。\n{dest_note}'
            self.status.config(text=msg, foreground='green')

        # 处理成功后：检验「结果行数过滤」条件
        if not manual and self.preview_only_enabled.get():
            result_text = result_buf.getvalue()
            total_lines = result_text.count('\n')
            thr = int(self.preview_only_threshold.get() or '0')
            if total_lines < thr:
                # 行数不足阈值：仅输出到预览，不保存文件
                self.last_output_path = None
                self.btn_open_output.state(['disabled'])
                self._show_manual_result(result_text)
                # 更新状态栏提示
                last_text = self.status.cget('text').rstrip()
                self.status.config(
                    text=(f'{last_text}\n'
                          f'（共 {total_lines:,} 行，少于阈值 {thr:,} 行，'
                          f'仅输出到结果预览框，未保存文件）'),
                    foreground='green')
                return
            else:
                # 行数达标：写入文件后按常规呈现
                out_path = get_unique_output_path(out_path)
                with open(out_path, 'w', encoding='utf-8') as f:
                    f.write(result_text)
                # 更新状态栏中的路径提示
                last_text = self.status.cget('text')
                line0 = last_text.split('\n')[0] if '\n' in last_text else last_text
                self.status.config(
                    text=f'{line0}\n结果已保存至: {out_path}',
                    foreground='green')

        # 文件模式启用「打开输出位置」；手动模式直接在预览呈现
        if manual:
            self.last_output_path = None
            self.btn_open_output.state(['disabled'])
            # 直接将内存中的结果呈现到预览（上限 1 万行，可编辑 / 复制）
            self._show_manual_result(result_buf.getvalue())
        else:
            self.last_output_path = out_path
            self.btn_open_output.state(['!disabled'])
            # 从输出文件读取前若干行刷新预览
            self._preview_result_file(out_path, manual=False)

    def _run_wordextract(self, target_path, ref_path, target_enc, ref_enc, manual):
        """简词提取：按正则提取并按词频降序排序，输出到文件夹或合并文件"""
        # 1. 读取正则列表（每行一个，空行跳过）
        regex_text = self.txt_regex.text.get('1.0', 'end-1c')
        regex_list = [l.strip() for l in regex_text.split('\n') if l.strip()]
        if not regex_list:
            messagebox.showwarning('提示', '请输入至少一个正则表达式。')
            return

        # 2. 提取前N行（未勾选则不限制）
        extract_limit = 0
        if self.extract_limit_enabled.get():
            try:
                extract_limit = int(self.extract_limit.get() or '0')
            except ValueError:
                extract_limit = 0

        merge = self.merge_results.get()
        include_header = self.include_regex_header.get()

        # 3. 确定输出路径
        out = self.output_name.get().strip() or '简词提取结果'
        if os.path.isabs(out):
            output_path = out
        else:
            save_dir = os.getcwd() if manual else (os.path.dirname(target_path) or '.')
            output_path = os.path.join(save_dir, out)
        if not merge:
            # 非合并模式：output_path 是目录，去掉可能的 .txt/.yaml 后缀
            if output_path.lower().endswith(('.txt', '.yaml')):
                output_path = os.path.splitext(output_path)[0]
        # 合并模式：output_path 是文件路径，process_wordextract 自动补 .txt

        # 4. 执行处理
        op_name = '简词提取'
        self.status.config(text=f'正在执行{op_name}处理...', foreground='black')
        self.root.update_idletasks()
        self.preview_result.set_editable(False)
        self._set_preview(self.preview_result, '')
        self._set_result_lines('（处理中…）')

        result = process_wordextract(
            target_path, ref_path, regex_list, output_path,
            extract_limit=extract_limit, merge_results=merge,
            include_regex_header=include_header,
            progress_callback=self._progress,
            target_enc=target_enc, ref_enc=ref_enc)

        if result is None:
            self.status.config(text=f'{op_name}处理失败。', foreground='red')
            return
        total, counts, result_text = result

        # 5. 显示结果
        if merge:
            # 合并模式：process_wordextract 返回结果文本，由这里决定是否写文件
            out_file = output_path
            if not out_file.lower().endswith(('.txt', '.yaml')):
                out_file += '.txt'

            # 检查「结果行数过滤」
            total_lines = result_text.count('\n')
            if not manual and self.preview_only_enabled.get():
                thr = int(self.preview_only_threshold.get() or '0')
                if total_lines < thr:
                    # 行数不足阈值：仅输出到预览，不保存文件
                    self._show_manual_result(result_text)
                    detail = '、'.join(
                        f'正则{i + 1}:{c}行' for i, c in enumerate(counts))
                    self.status.config(
                        text=(f'{op_name}完成！共匹配 {total:,} 行（{detail}）。\n'
                              f'少于阈值 {thr:,} 行，仅输出到结果预览框，未保存文件'),
                        foreground='green')
                    self.last_output_path = None
                    self.btn_open_output.state(['disabled'])
                    self._set_result_lines(
                        f'共 {len(counts)} 个正则，总匹配 {total:,} 行')
                    return

            # 行数达标或未启用过滤：写入文件
            try:
                with open(out_file, 'w', encoding='utf-8') as f:
                    f.write(result_text)
            except Exception as e:
                messagebox.showerror('错误', f'写入文件失败: {e}')
                self.status.config(text=f'{op_name}处理失败。', foreground='red')
                return

            detail = '、'.join(f'正则{i + 1}:{c}行' for i, c in enumerate(counts))
            msg = (f'{op_name}完成！共匹配 {total:,} 行，合并为 1 个文件'
                   f'（{detail}）。\n结果已保存至: {out_file}')
            self.status.config(text=msg, foreground='green')
            self.last_output_path = out_file
            self.btn_open_output.state(['!disabled'])
            self._preview_result_file(out_file, manual=False)
        else:
            # 非合并模式：多个文件
            detail = '、'.join(f'{i + 1}.txt:{c}' for i, c in enumerate(counts))
            msg = (f'{op_name}完成！共匹配 {total:,} 行，生成 {len(counts)} 个文件'
                   f'（{detail}）。\n结果已保存至文件夹: {output_path}')
            self.status.config(text=msg, foreground='green')
            first_file = os.path.join(output_path, '1.txt')
            if os.path.exists(first_file):
                self.last_output_path = first_file
                self.btn_open_output.state(['!disabled'])
                self._preview_result_file(first_file, manual=False)
            else:
                self.last_output_path = output_path
                self.btn_open_output.state(['!disabled'])
        self._set_result_lines(
            f'共 {len(counts)} 个正则，总匹配 {total:,} 行')

    def _show_manual_result(self, text):
        """手动输入模式：把处理结果显示到结果预览（上限 1 万行，可直接编辑 / 复制）

        结果只在内存中生成，不写文件；超出 1 万行的部分截断并提示。
        """
        lines = text.splitlines()
        total = len(lines)
        truncated = total > self.preview_result.MAX_LINES
        if truncated:
            lines = lines[:self.preview_result.MAX_LINES]
        self.preview_result.set_editable(True)
        self._set_preview(self.preview_result, '\n'.join(lines))
        note = '（可在此直接编辑 / 复制'
        note += '，超出 1 万行部分已截断）' if truncated else '）'
        self._set_result_lines(f'共 {total:,} 行{note}')

    def _preview_result_file(self, out_path, manual=False):
        """处理完成后把输出文件前若干行载入『处理结果预览』框（带行号）

        manual=True（待处理为手动输入）时，结果预览设为可编辑，便于直接修改/复制。
        """
        if out_path and os.path.exists(out_path):
            text, total = read_preview(out_path, 100, enc_choice='utf-8')
            self.preview_result.set_editable(manual)
            self._set_preview(self.preview_result, text)
            base = self._fmt_lines(out_path, total)
            self._set_result_lines(
                base + ('（可在此直接编辑 / 复制）' if manual else ''))
        else:
            self.preview_result.set_editable(False)
            self._set_preview(self.preview_result, '（无结果）')
            self._set_result_lines('')


    # =====================================================================
    #  词组编码功能（单字表 → 待编码词组，按切片规则生成词组码）
    # =====================================================================
    def _build_phrase_options(self, parent):
        """词组编码选项行：左侧规则输入框（多行可编辑，默认五笔规则）+ 右侧预设按钮 / 复选框 / 忽略表输入框，
        布局镜像「简词提取」的选项行（与正则框同位置、同结构，作为「处理选项」内的子选项）。"""
        pad = {'padx': 8, 'pady': 1}
        frm_rule = ttk.LabelFrame(
            parent, text='取码规则（生成规则，默认五笔规则）', padding=8)
        frm_rule.pack(fill='x', **pad)
        frm_left = ttk.Frame(frm_rule)
        frm_left.pack(side='left', fill='y')
        # 左侧：规则输入框（默认填充五笔规则），点击「词组生成规则:」标签查看语法说明
        lbl_rule = ttk.Label(frm_left, text='生成规则:', cursor='hand2')
        lbl_rule.pack(side='left')
        lbl_rule.bind('<Button-1>', lambda e: self._show_phrase_rule_help())
        lbl_rule.bind('<Enter>', lambda e: lbl_rule.config(foreground='#0066cc'))
        lbl_rule.bind('<Leave>', lambda e: lbl_rule.config(foreground=''))
        self.txt_phrase_rule = _Preview(
            frm_left, (CJK_FONT, 9), height=3, width=44,
            line_numbers=False, editable=True,
            on_overlimit=self._on_preview_overlimit)
        self.txt_phrase_rule.frame.pack(side='left', fill='y', padx=4)
        # 默认填充五笔规则
        self.txt_phrase_rule.set_content(PHRASE_PRESETS['五笔规则'])

        # 「?」帮助按钮：点击弹出规则语法说明（“生成规则”标签点击为其冗余入口）
        self.btn_rule_help = ttk.Button(
            frm_left, text='？', width=2,
            command=self._show_phrase_rule_help)
        self.btn_rule_help.pack(side='left', padx=(4, 0))

        # 右侧：两行控件（与左侧等高）
        frm_right = ttk.Frame(frm_left)
        frm_right.pack(side='left', fill='y', padx=(12, 0))
        frm_r1 = ttk.Frame(frm_right)
        frm_r1.pack(anchor='w')
        # 预设规则改为「加载默认规则」子菜单（五笔/两笔/拼音/速成；ttk 自带下拉箭头）
        mb_preset = ttk.Menubutton(frm_r1, text='加载默认规则')
        mb_preset.pack(side='left', padx=4)
        m_preset = tk.Menu(mb_preset, tearoff=0)
        for _name in PHRASE_PRESETS:
            m_preset.add_command(
                label=_name, command=lambda n=_name: self._phrase_load_preset(n))
        mb_preset.config(menu=m_preset)
        frm_r2 = ttk.Frame(frm_right)
        frm_r2.pack(anchor='w', pady=(2, 0))
        self.chk_allow_predef = ttk.Checkbutton(
            frm_r2, text='允许预定义编码', variable=self.allow_predef)
        self.chk_allow_predef.pack(side='left', padx=4)
        # 忽略标点符号：主页面复选框（默认勾选），不再放进二级菜单
        self.chk_ignore_punct = ttk.Checkbutton(
            frm_r2, text='忽略标点符号', variable=self.ignore_punct)
        self.chk_ignore_punct.pack(side='left', padx=4)
        # 忽略表：单行输入框，直接填需剔除的字符（不再用子菜单/弹窗）
        # 「忽略标点符号」取消勾选时，整张忽略表禁用（变灰、不生效）
        frm_r3 = ttk.Frame(frm_right)
        frm_r3.pack(anchor='w', pady=(2, 0))
        self.lbl_ignore_extra = ttk.Label(frm_r3, text='忽略表:')
        self.lbl_ignore_extra.pack(side='left')
        self.ent_ignore_extra = ttk.Entry(
            frm_r3, textvariable=self.ignore_extra, width=40)
        self.ent_ignore_extra.pack(side='left', padx=4, fill='x', expand=True)
        self._update_ignore_state()
        self.ignore_punct.trace_add('write', lambda *a: self._update_ignore_state())

    # ---- 忽略表状态联动 ----
    def _update_ignore_state(self):
        """「忽略标点符号」勾选时忽略表可用；取消勾选时整张忽略表变灰、不生效"""
        if self.ignore_punct.get():
            self.ent_ignore_extra.state(['!disabled'])
            self.lbl_ignore_extra.state(['!disabled'])
        else:
            self.ent_ignore_extra.state(['disabled'])
            self.lbl_ignore_extra.state(['disabled'])

    # ---- 规则预设 / 忽略表编辑 / 规则语法说明 ----
    def _phrase_load_preset(self, name):
        """把指定预设规则载入规则输入框（覆盖当前内容）"""
        if name in PHRASE_PRESETS:
            self.txt_phrase_rule.set_content(PHRASE_PRESETS[name])

    def _show_phrase_rule_help(self):
        """弹出词组生成规则语法说明窗口"""
        top = tk.Toplevel(self.root)
        top.title('词组生成规则语法')
        top.transient(self.root)
        top.grab_set()
        frm = ttk.Frame(top, padding=10)
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text='词组生成规则语法', font=(CJK_FONT, 11, 'bold')).pack(
            anchor='w', pady=(0, 6))
        help_text = (
            '【词组编码是做什么的】\n'
            '根据左侧「单字表」（每行：单字 + 它的编码）自动为右侧「待编码词组」生成编码。\n'
            '例：单字表有  中=khh / 国=lgd，词组「中国」就能按规则拼出 khh.lgd 之类的编码。\n\n'
            '【规则写在哪】\n'
            '主界面「生成规则」框，每行一条规则，格式：   字数范围 = 取码规则\n'
            '程序从上往下逐条匹配，第一个命中的规则就用于该词组。\n\n'
            '一、字数范围（决定这条规则管几个字的词组）\n'
            '  • 单个数字 N        如 2      → 只用于二字词\n'
            '  • A,B             如 4,99    → 用于 4 到 99 字词（含两端）\n'
            '  • 一端留空          如 ,99    → 1 到 99 字词；如 4, → 四字及以上\n'
            '  常见写法：2 管二字词，3 管三字词，4,99 管四字及更长的词。\n\n'
            '二、取码规则（用 [字序][码序] 从每个单字取码，多个区块用 + 连接）\n'
            '  [字序] 选「第几个字」：\n'
            '      0      第一个字        -1     最后一个字\n'
            '      :      所有字          0:3    前三个字        1,2    第2、第3个字\n'
            '      也支持任意切片：2:4（第3、4个字）、2:（第3个到末字）、:4（前4个字）\n'
            '  [码序] 选「取该字的哪几码」（索引从 0 开始，即 0=第1码/1=第2码…）：\n'
            '      :2     前两码          :      全部码\n'
            '      0      第一码          0,-1   首码+末码\n'
            '      也支持任意切片：2:4（第3、4码）、2:（第3码到末码）、:4（前4码）\n'
            '  + 把多个区块拼起来，如 [0][:2] + [1][:2] = 「首字前两码」连「次字前两码」。\n'
            '  例：[0][2:4] = 取首字编码里第3、4码（如 字=PBFF → 取 "FF"）；\n'
            '      [0:2][2:4] = 取前两个字的「第3、4码」再拼接。\n\n'
            '三、单字编码怎么取（重要）\n'
            '  单字表里一个字可能有多码（如五笔有两码/三码/全码）。取码时：\n'
            '  • 写 :2 只想取两码 → 优先用两码，没有两码就退回一码；\n'
            '  • 同长度有多码（如重码）会全部保留，用空格隔开。\n\n'
            '【一步步算给你看】（用五笔规则算三字词「计算机」）\n'
            '  单字表：计=YF / 算=THA / 机=SM\n'
            '  规则  3 = [0][0] + [1][0] + [-1][:2]\n'
            '   → [0][0]  首字「计」首码        = Y\n'
            '   → [1][0]  第2字「算」首码       = T\n'
            '   → [-1][:2] 末字「机」前两码      = SM\n'
            '   → 结果：YTSM\n\n'
            '【其他选项】\n'
            '  • 忽略标点符号（默认勾选）：编码前先把词组里的标点、空格删掉再处理；\n'
            '    取消勾选则标点不再被忽略（即使「忽略表」框里还显示着它们）。\n'
            '  • 忽略表：额外要删除的字符，直接填在框里（默认已含常见编辑符号）。\n'
            '  • 允许预定义编码（默认勾选）：词组行本身若已带编码，则原样保留、不再重算。\n\n'
            '【一键预设】点「加载默认规则」可选 五笔 / 两笔 / 拼音 / 速成 规则直接填入。'
        )
        txt = tk.Text(frm, wrap='word', width=82, height=30)
        txt.pack(fill='both', expand=True)
        txt.insert('1.0', help_text)
        txt.config(state='disabled')
        ttk.Button(frm, text='关闭', command=top.destroy).pack(pady=(6, 0))
        top.update_idletasks()
        top.minsize(520, 440)
        top.geometry('+%d+%d' % (self.root.winfo_rootx() + 60,
                                  self.root.winfo_rooty() + 60))

    # ---- 生成 & 保存 ----
    def _phrase_build_report(self, report):
        """把生成报告字典整理为可读文本（缺失字 / 未覆盖长度 / 计数等）"""
        rep = []
        rep.append(f'已生成: {report["generated"]:,} 行')
        rep.append(f'预定义保留: {report["predef"]:,} 行')
        if report['empty']:
            rep.append(f'空词组（忽略后为空）: {report["empty"]:,} 行')
        if report['failed']:
            rep.append(f'规则不适用（字序越界等）: {report["failed"]:,} 行')
        if report['unmatched']:
            rep.append('— 未覆盖长度（跳过）—')
            for length in sorted(report['unmatched']):
                cnt, samples = report['unmatched'][length]
                rep.append(f'长度 {length}: {cnt:,} 个  样例: {", ".join(samples[:5])}')
        if report['missing']:
            rep.append('— 含未收录单字（跳过）—')
            for ch in sorted(report['missing']):
                cnt, samples = report['missing'][ch]
                rep.append(f'缺字[{ch}]: {cnt:,} 次  样例: {", ".join(samples[:5])}')
        if not (report['unmatched'] or report['missing'] or report['empty'] or report['failed']):
            rep.append('全部词组均已成功处理。')
        return '\n'.join(rep)

    def _run_phrasecode(self):
        """词组编码：按规则框生成词组编码，输出到文件并刷新标准结果预览（与简词提取共用开始处理按钮）"""
        out_lines, report = self._phrase_generate()
        if out_lines is None:
            messagebox.showwarning('提示', report)
            return
        text = '\n'.join(out_lines)
        op_name = '词组编码'
        # 输出路径（与标准逻辑一致：相对路径落在当前工作目录）
        out = self.output_name.get().strip() or self.DEFAULT_OUTPUT_NAMES['phrasecode']
        if os.path.isabs(out):
            out_path = out
        else:
            out_path = os.path.join(os.getcwd(), out)
        if not out_path.lower().endswith(('.txt', '.yaml')):
            out_path += '.txt'
        # 清空上一次结果预览
        self.preview_result.set_editable(False)
        self._set_preview(self.preview_result, '')
        self._set_result_lines('（处理中…）')
        self.status.config(text=f'正在执行{op_name}处理...', foreground='black')
        self.root.update_idletasks()
        # 结果行数过滤：启用且行数不足阈值时仅输出到预览，不保存文件
        if self.preview_only_enabled.get():
            total_lines = len(out_lines)
            thr = int(self.preview_only_threshold.get() or '0')
            if total_lines < thr:
                self.last_output_path = None
                self.btn_open_output.state(['disabled'])
                self._show_manual_result(text)
                self.status.config(
                    text=(f'{op_name}完成：生成 {report["generated"]:,} 行，'
                          f'预定义保留 {report["predef"]:,} 行；未覆盖 '
                          f'{sum(v[0] for v in report["unmatched"].values()):,} 行，'
                          f'缺字 {sum(v[0] for v in report["missing"].values()):,} 行，'
                          f'空词组 {report["empty"]:,} 行。\n'
                          f'（共 {total_lines:,} 行，少于阈值 {thr:,} 行，'
                          f'仅输出到结果预览框，未保存文件）'),
                    foreground='green')
                self._phrase_show_report(self._phrase_build_report(report))
                return
        # 写文件（唯一文件名，避免覆盖）
        out_path = get_unique_output_path(out_path)
        try:
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(text + '\n')
        except Exception as e:
            messagebox.showerror('保存失败', str(e))
            return
        self.last_output_path = out_path
        self.btn_open_output.state(['!disabled'])
        self._preview_result_file(out_path, manual=False)
        unmatched_n = sum(v[0] for v in report['unmatched'].values())
        missing_n = sum(v[0] for v in report['missing'].values())
        self.status.config(
            text=(f'{op_name}完成：生成 {report["generated"]:,} 行，'
                  f'预定义保留 {report["predef"]:,} 行；未覆盖 {unmatched_n:,} 行，'
                  f'缺字 {missing_n:,} 行，空词组 {report["empty"]:,} 行。\n'
                  f'结果已保存至: {out_path}'),
            foreground='green')
        self._phrase_show_report(self._phrase_build_report(report))

    def _phrase_show_report(self, report_text):
        """生成完成后以弹窗展示处理报告（缺失字 / 未覆盖长度 / 计数等）"""
        top = tk.Toplevel(self.root)
        top.title('处理报告')
        top.transient(self.root)
        top.grab_set()
        frm = ttk.Frame(top, padding=10)
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text='处理报告', font=(CJK_FONT, 11, 'bold')).pack(
            anchor='w', pady=(0, 6))
        txt = tk.Text(frm, wrap='word', width=72, height=18)
        txt.pack(fill='both', expand=True)
        txt.insert('1.0', report_text)
        txt.config(state='disabled')
        ttk.Button(frm, text='关闭', command=top.destroy).pack(pady=(6, 0))
        top.update_idletasks()
        top.minsize(420, 200)
        # 点击窗口外不关闭，但允许用关闭按钮；定位到父窗口中心附近
        top.geometry('+%d+%d' % (self.root.winfo_rootx() + 60,
                                  self.root.winfo_rooty() + 60))

    # ---- 切片取码引擎 ----
    def _phrase_parse_char_table(self):
        """解析单字表：每行「单字 编码」，同字可多行 → {字: [码1, 码2, ...]}"""
        table = {}
        # 单字表复用标准文件区左框（待处理文件 → 词组编码下为「单字表」）
        text = self.preview_t.get_content()
        for line in text.splitlines():
            s = line.strip()
            if not s or s.startswith('#'):
                continue
            parts = s.split(None, 1)
            if len(parts) < 2:
                continue
            ch, code = parts[0], parts[1].strip()
            if not ch or not code:
                continue
            table.setdefault(ch, []).append(code)
        return table

    @staticmethod
    def _phrase_parse_rule(rule):
        """把规则字符串拆成 token 列表 [(字序表达式, 码序表达式), ...]"""
        tokens = []
        for m in re.finditer(r'\[([^\]]*)\]\[([^\]]*)\]', rule):
            tokens.append((m.group(1).strip(), m.group(2).strip()))
        if not tokens:
            raise ValueError('未找到 [字序][码序] 单元，请按说明书填写规则')
        return tokens

    @staticmethod
    def _phrase_match_range(length, range_str):
        """判断词组长度是否落在范围内：'A,B'闭区间 / 'N'精确 / 'A,'到无穷 / ',B'到B"""
        s = (range_str or '').strip()
        if not s:
            return False
        if ',' in s:
            a, _, b = s.partition(',')
            a, b = a.strip(), b.strip()
            lo = int(a) if a else 1
            hi = int(b) if b else None
            if length < lo:
                return False
            if hi is not None and length > hi:
                return False
            return True
        try:
            return length == int(s)
        except ValueError:
            return False

    @staticmethod
    def _phrase_parse_slice(expr, n):
        """把切片表达式（如 ':', ':2', '2:', '1:3'）解析为 slice 对象"""
        a, _, b = expr.partition(':')
        start = int(a) if a != '' else None
        stop = int(b) if b != '' else None
        return slice(start, stop)

    @staticmethod
    def _phrase_eval_char_idx(expr, chars):
        """解析字序表达式 → 字符索引列表（安全防越界；越界则整体返回空）"""
        n = len(chars)
        if ',' in expr:
            parts = [p for p in expr.split(',') if p != '']
            try:
                idxs = [int(p) for p in parts]
            except ValueError:
                return []
            res = []
            for i in idxs:
                if i < -n or i >= n:
                    return []
                res.append(i if i >= 0 else n + i)
            return res
        if ':' in expr:
            try:
                sl = App._phrase_parse_slice(expr, n)
            except (ValueError, TypeError):
                return []
            return list(range(n))[sl]
        try:
            i = int(expr)
        except ValueError:
            return []
        if i < -n or i >= n:
            return []
        return [i if i >= 0 else n + i]

    @staticmethod
    def _phrase_eval_code_idx(code, expr):
        """解析码序表达式 → 对单条码字符串取出一个片段（安全防越界；越界返回 None）"""
        n = len(code)
        if ',' in expr:
            parts = [p for p in expr.split(',') if p != '']
            try:
                idxs = [int(p) for p in parts]
            except ValueError:
                return None
            uniq = list(dict.fromkeys(idxs))  # 离散位置去重（保序），避免同一码被取两次
            frag = ''
            for i in uniq:
                if i < -n or i >= n:
                    return None
                frag += code[i if i >= 0 else n + i]
            # 提取结果再按字符去重（如单码字遇 0,-1 会得到 'aa'，应折叠为 'a'）
            return ''.join(dict.fromkeys(frag))
        if ':' in expr:
            try:
                sl = App._phrase_parse_slice(expr, n)
            except (ValueError, TypeError):
                return None
            positions = list(range(n))[sl]
            return ''.join(code[i] for i in positions)
        try:
            i = int(expr)
        except ValueError:
            return None
        if i < -n or i >= n:
            return None
        return code[i if i >= 0 else n + i]

    @staticmethod
    def _phrase_char_fragments(ch, codes, code_expr):
        """按 Q1 规则取单字候选片段：优先最长（两码优先、回退一码），同长多码全留"""
        frags = set()
        for code in codes:
            f = App._phrase_eval_code_idx(code, code_expr)
            if f is not None:
                frags.add(f)
        if not frags:
            return []
        maxlen = max(len(f) for f in frags)
        return [f for f in frags if len(f) == maxlen]

    @staticmethod
    def _phrase_apply_rule(chars, tokens, table):
        """对一组 token 求词组候选码（跨字笛卡尔积 + 整体去重保序）"""
        token_results = []
        for char_expr, code_expr in tokens:
            positions = App._phrase_eval_char_idx(char_expr, chars)
            if not positions:
                return []
            pos_frags = []
            for p in positions:
                ch = chars[p]
                frags = App._phrase_char_fragments(ch, table.get(ch, []), code_expr)
                if not frags:
                    return []
                pos_frags.append(frags)
            token_strs = [''.join(comb) for comb in itertools.product(*pos_frags)]
            token_results.append(token_strs)
        final = [''.join(comb) for comb in itertools.product(*token_results)]
        seen, deduped = set(), []
        for f in final:
            if f not in seen:
                seen.add(f)
                deduped.append(f)
        return deduped

    def _phrase_generate(self):
        """主流程：解析单字表/规则组 → 逐词组剔除忽略字 → 范围匹配 → 生成或跳过。

        返回 (out_lines, report)；若前置条件不满足返回 (None, 错误提示)。
        """
        table = self._phrase_parse_char_table()
        if not table:
            return None, '单字表为空或格式不正确（每行需为「单字 编码」，同字可多行）。'

        parsed_groups = []
        for line in self.txt_phrase_rule.get_content().splitlines():
            s = line.strip()
            if not s or s.startswith('#'):
                continue
            # 支持「范围 = 规则」（等号分隔，等号两侧空格可选）与旧式「范围 规则」（空格分隔）
            if '=' in s:
                r, rule = s.split('=', 1)
                r, rule = r.strip(), rule.strip()
            else:
                parts = s.split(None, 1)
                if len(parts) < 2:
                    return None, f'规则行格式错误（需「范围 = 规则」）: {s}'
                r, rule = parts[0], parts[1].strip()
            if not rule:
                return None, f'规则行格式错误（需「范围 = 规则」）: {s}'
            try:
                tokens = self._phrase_parse_rule(rule)
            except Exception as e:
                return None, f'规则解析失败: {rule}\n{e}'
            parsed_groups.append((r, tokens))
        if not parsed_groups:
            return None, '请至少填写一组有效规则（范围 = 规则）。'

        # 「忽略标点符号」勾选时忽略表生效（整框内容）；取消勾选则整张忽略表禁用、不忽略任何字符
        ignore_chars = self.ignore_extra.get() if self.ignore_punct.get() else ''
        translate_tbl = str.maketrans('', '', ignore_chars) if ignore_chars else None
        allow_predef = self.allow_predef.get()

        report = {'unmatched': {}, 'missing': {}, 'empty': 0,
                  'predef': 0, 'generated': 0, 'failed': 0}
        SAMPLE = 10
        out_lines = []

        # 待编码词组复用标准文件区右框（参考文件 → 词组编码下为「待编码词组」）
        for raw in self.preview_r.get_content().splitlines():
            stripped = raw.strip()
            if not stripped or stripped.startswith('#'):
                continue
            # 拆分「词组 + 可选预定义编码」（首个空白符），无论是否允许预定义都先隔离，
            # 不勾选时旧编码被忽略、按词组重新生成（即“替换原有”）。
            parts = stripped.split(None, 1)
            phrase_raw = parts[0]
            predef = parts[1].strip() if len(parts) > 1 else ''
            # 输出词组名保留原始文字（含标点）；忽略表只作用于"取码用的字数组"
            phrase_code = phrase_raw.translate(translate_tbl) if translate_tbl else phrase_raw
            if not phrase_code:
                report['empty'] += 1
                continue
            chars = list(phrase_code)  # 按 Unicode 字符切分（含扩展 B 正确）
            length = len(chars)

            # 允许预定义编码且本行带编码 → 原样保留（仅排序，不重新生成）
            if allow_predef and predef:
                out_lines.append(f'{phrase_raw}\t{predef}')
                report['predef'] += 1
                continue

            # 缺失字检测
            miss = [c for c in chars if c not in table]
            if miss:
                key = ''.join(sorted(set(miss)))
                bucket = report['missing'].setdefault(key, [0, []])
                bucket[0] += 1
                if len(bucket[1]) < SAMPLE:
                    bucket[1].append(phrase_raw)
                continue

            # 范围匹配（自上而下首个命中）
            matched = None
            for r, tokens in parsed_groups:
                if self._phrase_match_range(length, r):
                    matched = tokens
                    break
            if matched is None:
                bucket = report['unmatched'].setdefault(length, [0, []])
                bucket[0] += 1
                if len(bucket[1]) < SAMPLE:
                    bucket[1].append(phrase_raw)
                continue

            # 生成候选码
            candidates = self._phrase_apply_rule(chars, matched, table)
            if not candidates:
                report['failed'] += 1
                continue
            for cand in candidates:
                out_lines.append(f'{phrase_raw}\t{cand}')
            report['generated'] += 1

        return out_lines, report


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
