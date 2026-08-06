#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
catenary_cad — CAD COM 交互封装（GstarCAD / AutoCAD 通用）
============================================================

从主程序 catenary_app.py 抽取的 win32com 交互层，供主程序与辅助脚本
统一 import，消除代码重复（审查意见 Opt 2）。

包含：
  - 连接：get_gcad / _detect_cad_windows / foreground_cad_prefer
  - 图层：ensure_layer
  - 交互：pick_point / pick_entity / pick_angle
  - 提取：extract_points / get_entity_by_handle
  - 绘制：draw_catenary / draw_distance_line / draw_distance_label /
          draw_wind_curve / draw_wind_label / draw_bundle_line /
          draw_suspension_line

依赖：pywin32 (win32com / pythoncom / win32gui / win32con)。
⚠️ 惰性导入（审查 F1）：win32com 相关 import 只在真正连 CAD / 绘制时
按需触发（见 _win32()），模块顶层不依赖 pywin32——保证纯数学 selftest
与 CI / 无 CAD 环境仍可 import 本模块与 catenary_core。
图层/颜色常量统一在此定义（全项目唯一出处）。
"""

import math
import os
import subprocess
import time
import winreg

# ─── 图层与颜色（全项目唯一出处）──────────────────────────────────────
LAYER_NAME = "悬链线"
LAYER_DISTANCE = "最小距离示意"
LAYER_WIND = "风偏弧垂曲线"
COLOR_WIND = 5  # ACI: 5 = 蓝
LAYER_BUNDLE = "分裂导线"
COLOR_BUNDLE = 2  # ACI: 2 = 黄（用户指定：分裂线统一黄色）
LAYER_SUSPENSION = "悬垂串"
COLOR_SUSPENSION = 6  # ACI: 6 = 品红（竖直串 / 参照串）
LAYER_SUSPENSION_WIND = "悬垂串风偏"
COLOR_SUSPENSION_WIND = 4  # ACI: 4 = 浅蓝（风偏后的悬垂串）

# 风偏角度约定（模块③与模块⑤统一）：
#   以曲线起点→终点（点选先后顺序）为前进方向，**向右偏为正、向左偏为负**。
#   绕弦右手正旋转会把弧垂(向下)转向观察者的左侧，故取负号实现「正=右偏」。


# ─── 连接 ───────────────────────────────────────────────────────────

# CAD 程序 ProgID 偏好顺序：先 GstarCAD，后 AutoCAD（均可被 --cad 强制）
_CAD_PROGIDS = {
    "gstar": ["GCAD.Application", "GstarCAD.Application"],
    "autocad": ["AutoCAD.Application"],
}
_CAD_LABELS = {
    "GCAD.Application": "GstarCAD",
    "GstarCAD.Application": "GstarCAD",
    "AutoCAD.Application": "AutoCAD",
}


def _list_cad_progids(prefix: str) -> list[str]:
    """动态枚举注册表中以 prefix 开头的所有 ProgID（含版本号），版本新者在前。

    例如本机同时装了 AutoCAD 2020 与 2022 时，无版本别名
    `AutoCAD.Application` 只指向最后注册的 2022，但用户可能开着 2020 实例——
    仅靠别名永远连不上旧版实例（GetActiveObject 按 2022 的 CLSID 查 ROT 落空）。
    必须把 `AutoCAD.Application.23/.24` 等版本化 ProgID 全部枚举出来逐个尝试。

    排序：裸 ProgID（无后缀）最前（保持"最新版本优先"语义），
    其余按版本数字降序（.24 → .23 → …）。
    """
    found: list[str] = []
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as hk:
            i = 0
            while True:
                try:
                    name = winreg.EnumKey(hk, i)
                    i += 1
                    if name.lower().startswith(prefix.lower()):
                        found.append(name)
                except OSError:
                    break
    except Exception:
        pass

    def _ver(p: str) -> tuple[int, ...]:
        return tuple(int(x) for x in p.split(".")[2:] if x.isdigit())

    bare = [p for p in found if p.lower() == prefix.lower()]
    vers = [p for p in found if p.lower() != prefix.lower()]
    vers.sort(key=_ver, reverse=True)
    return bare + vers


def _cad_label(progid: str) -> str:
    """把 ProgID 映射为人类可读的 CAD 名（兼容带版本号后缀的 ProgID）。

    例如 'AutoCAD.Application.23.1' → 'AutoCAD'（2020 的版本化 ProgID）。
    """
    for base, label in _CAD_LABELS.items():
        if progid == base or progid.lower().startswith(base.lower()):
            return label
    return progid


# ─── COM 错误诊断 ───────────────────────────────────────────────────

def _hr_code(exc: Exception) -> int | None:
    """从 COM 异常中提取 HRESULT 错误码（十进制）。"""
    if exc is None:
        return None
    # pywin32 抛出的异常通常有 hresult 属性或 args[0] 是整数
    for attr in ("hresult", "errno", "winerror"):
        try:
            v = getattr(exc, attr, None)
            if isinstance(v, int):
                return v
        except Exception:
            pass
    for arg in exc.args:
        if isinstance(arg, int):
            return arg
    return None


def _classify_com_error(exc: Exception) -> tuple[str, str]:
    """把 COM 异常分类为 (简短类型, 用户可读说明+建议)。"""
    hr = _hr_code(exc)
    msg = str(exc)
    if hr is not None:
        # 统一转成 32 位有符号/无符号再判断
        hr_u = hr & 0xFFFFFFFF
        if hr_u == 0x800401F3 or hr_u == 0x80040154:
            return (
                "CAD 未注册",
                "当前 Windows 注册表中找不到 AutoCAD / GstarCAD 的 COM 标识。\n"
                "常见原因：\n"
                "  1) 本机未安装 AutoCAD 或 GstarCAD；\n"
                "  2) 安装的是绿色版/精简版，未写入注册表；\n"
                "  3) 操作系统为 64 位，但运行了 32 位 Python/COM，"
                "无法读取 64 位注册表。\n\n"
                "建议：用 64 位 Python 运行；或重装 CAD 并勾选“用于第三方应用程序的 COM 支持”。",
            )
        if hr_u == 0x800702EC:
            return (
                "需要管理员权限",
                "CAD 进程与当前工具权限不一致（通常 CAD 以管理员运行，而工具没有）。\n"
                "建议：右键本工具 →“以管理员身份运行”；或者把 CAD 也改成普通用户权限启动。",
            )
        if hr_u == 0x800401E3 or hr_u == 0x80040001:
            return (
                "CAD 服务器不可用",
                "COM 服务器没有响应。请确认 CAD 已正常启动并至少打开了一个图纸。",
            )
        if hr_u == 0x80010105:
            return (
                "CAD 内部错误",
                "CAD 进程内部出错（RPC_E_SERVERFAULT）。请尝试保存图纸后重启 CAD。",
            )
    if "无效的类字符串" in msg or "Class not registered" in msg:
        return (
            "CAD 未注册",
            "当前 Windows 注册表中找不到 AutoCAD / GstarCAD 的 COM 标识。\n"
            "建议：确认本机已安装 CAD；如已安装，尝试用 64 位 Python 运行本工具。",
        )
    if "需要提升" in msg or "elevation" in msg.lower():
        return (
            "需要管理员权限",
            "CAD 与工具权限不一致。建议右键本工具 →“以管理员身份运行”。",
        )
    return ("连接失败", msg)


def list_registered_cad() -> dict[str, list[str]]:
    """扫描注册表，返回本机已注册的 CAD ProgID 列表。

    Returns:
        {"AutoCAD": ["AutoCAD.Application", ...], "GstarCAD": [...]}
        若某类为空，表示注册表中未找到对应 COM 注册项。
    """
    result: dict[str, list[str]] = {"AutoCAD": [], "GstarCAD": []}
    for label, progids in [
        # 动态枚举（含版本化 ProgID，如 AutoCAD.Application.23/.24）
        ("GstarCAD", _list_cad_progids("GCAD.Application")
         + _list_cad_progids("GstarCAD.Application")),
        ("AutoCAD", _list_cad_progids("AutoCAD.Application")),
    ]:
        for progid in progids:
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid) as key:
                    # 存在 ProgID 键就算注册；进一步检查 CLSID 更严谨
                    try:
                        with winreg.OpenKey(key, "CLSID") as clsid_key:
                            winreg.QueryValueEx(clsid_key, "")  # 仅确认可读
                    except FileNotFoundError:
                        continue
                result[label].append(progid)
            except FileNotFoundError:
                pass
            except OSError:
                pass
    return result


def diagnose_cad_connection(prefer: str | None = None) -> str:
    """生成一段人类可读的 CAD 连接诊断报告。"""
    reg = list_registered_cad()
    lines = ["本机 CAD COM 注册状态："]
    for label in ("AutoCAD", "GstarCAD"):
        if reg[label]:
            lines.append(f"  {label}: 已注册 ({', '.join(reg[label])})")
        else:
            lines.append(f"  {label}: 未注册")
    total = sum(len(v) for v in reg.values())
    if total == 0:
        lines.append("\n结论：注册表中没有发现 AutoCAD / GstarCAD 的 COM 项。")
        lines.append("请先确认 CAD 已安装；如已安装，尝试用 64 位 Python 重新运行本工具。")
    else:
        lines.append("\n注册表项正常，但 GetActiveObject/Dispatch 仍可能因权限、CAD 未启动等失败。")
    return "\n".join(lines)


# ─── 提权处理（外部设备版移植）───────────────────────────────────────

def _is_elevation_error(e):
    """判断 COM 异常是否由 UAC 提权失败引起（非管理员环境最常见）。

    AutoCAD 的 COM LocalServer 注册为需要提权，非管理员进程通过
    Dispatch 激活时 Windows 拒绝启动 → CO_E_SERVER_EXEC_FAILURE (0x80010124)
    或返回"请求的操作需要提升"。
    """
    s = str(e).lower()
    return any(k in s for k in ("提升", "elevation", "server exec"))


def _get_cad_exe_path(progid):
    """从注册表 LocalServer32 查找 CAD 可执行文件路径。

    Dispatch 因 UAC 提权失败时，用此路径通过 subprocess 直接启动 CAD
    （不走 COM 激活，不触发提权），再用 GetActiveObject 连接。
    """
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID")
        clsid = winreg.QueryValue(key, "")
        winreg.CloseKey(key)
        server_key = winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32")
        raw = winreg.QueryValue(server_key, "")
        winreg.CloseKey(server_key)
        # LocalServer32 值形如 "D:\\...\\acad.exe /Automation" 或
        # "\"D:\\...\\acad.exe\" /Automation"。需要正确提取 exe 路径。
        raw = raw.strip()
        if raw.startswith('"'):
            end = raw.find('"', 1)
            if end > 0:
                candidate = raw[1:end]
                if os.path.isfile(candidate):
                    return candidate
        # 不带引号：路径可能含空格，逐段拼接直到找到存在的 exe 文件
        parts = raw.split()
        for i in range(len(parts), 0, -1):
            candidate = " ".join(parts[:i])
            if os.path.isfile(candidate):
                return candidate
        return None
    except Exception:
        return None


def _launch_cad_and_connect(progid, timeout=60):
    """当 Dispatch 因 UAC 提权失败时，直接启动 CAD 进程再连接。

    流程：
      1. 从注册表读取 CAD exe 路径
      2. subprocess.Popen 启动（不走 COM 激活，不需要提权）
      3. 轮询 GetActiveObject，等 CAD 完成 COM 注册（最多 timeout 秒）
    返回 COM 对象或 None。
    """
    win32com, _, _, _ = _win32()
    exe_path = _get_cad_exe_path(progid)
    if not exe_path or not os.path.isfile(exe_path):
        return None
    try:
        subprocess.Popen(
            [exe_path],
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )
    except Exception:
        return None
    # 轮询等待 CAD COM 注册完成
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(2)
        try:
            gc = win32com.client.GetActiveObject(progid)
            return gc
        except Exception:
            continue
    return None


def _check_runasadmin(exe_path):
    """检查 CAD exe 是否被设为"始终以管理员身份运行"（兼容性标志）。

    这是非管理员 Python 无法连接 CAD 的最常见原因：
    HKCU\\...\\AppCompatFlags\\Layers\\<exe_path> = "RUNASADMIN"
    设置后 CAD 以高完整性级别运行，COM 对象注册在高级别 ROT，
    非管理员进程无法通过 GetActiveObject 访问。
    """
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers")
        val = winreg.QueryValueEx(key, exe_path)
        winreg.CloseKey(key)
        return "RUNASADMIN" in str(val[0]).upper()
    except FileNotFoundError:
        return False
    except Exception:
        return False


# 进程检测缓存：同一 exe 名只跑一次 tasklist（多版本 ProgID 常指向同一 exe）
_proc_running_cache: dict[str, bool] = {}


def _is_cad_process_running(exe_path):
    """检查 CAD exe 是否有进程在运行（通过 tasklist）。

    当 GetActiveObject 失败但进程在运行时，通常意味着：
    1. CAD 以管理员身份运行（ROT 隔离），或
    2. CAD 尚未完成初始化（无打开的图纸）。

    带模块级缓存：本机 2020/2022 的 exe 都叫 acad.exe，多个版本化 ProgID
    会重复查询同一个进程名——缓存避免每次连接重复跑 tasklist（约 0.3s/次）。
    """
    exe_name = os.path.basename(exe_path).lower()
    if exe_name in _proc_running_cache:
        return _proc_running_cache[exe_name]
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {exe_name}"],
            capture_output=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
            encoding="gbk", errors="replace",
        )
        running = exe_name in result.stdout.lower()
    except Exception:
        running = False
    _proc_running_cache[exe_name] = running
    return running


def _collect_connect_diag(progids):
    """全部连接失败后收集场景诊断信息（每 exe 仅一次 tasklist，含缓存）。

    仅在所有 ProgID 都失败时调用——连接成功路径不触发任何慢速探测
    （tasklist/注册表扫描），保证已运行实例场景下秒连。
    """
    diag = []
    for progid in progids:
        exe_path = _get_cad_exe_path(progid)
        is_admin = _check_runasadmin(exe_path) if exe_path else False
        is_running = _is_cad_process_running(exe_path) if exe_path else False
        diag.append((progid, exe_path, is_admin, is_running))
    return diag


def _is_running_as_admin():
    """检查当前 Python 进程是否以管理员身份运行。"""
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _try_auto_elevate():
    """以管理员身份重启当前脚本（弹出 UAC 确认框）。

    使用 ShellExecuteW + "runas" 动词，用户确认后新进程以管理员权限运行，
    当前进程退出。如果用户拒绝 UAC 或调用失败，返回 False。

    使用场景：AutoCAD 以管理员身份运行，但 Python 脚本不是——
    COM 的完整性级别隔离导致 GetActiveObject 无法连接。
    提权后双方处于同一高级别，连接恢复正常。
    """
    try:
        import ctypes
        import sys

        params = " ".join(f'"{a}"' for a in sys.argv)
        # SW_SHOWNORMAL = 1
        result = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, params, None, 1
        )
        # ShellExecuteW 返回值 > 32 表示成功
        if result > 32:
            return True
        return False
    except Exception:
        return False


def _win32():
    """惰性导入 pywin32（审查 F1）。

    仅在实际连接 CAD / 交互拾取 / 绘制时调用，返回
    (win32com, pythoncom, win32gui, win32con) 四个模块。
    Python 会缓存已导入模块，重复调用无实际开销；
    缺 pywin32 时在此抛出 ImportError（带明确提示）。
    """
    try:
        import win32com.client
        import pythoncom
        import win32gui
        import win32con
    except ImportError as e:
        raise ImportError(
            "缺少 pywin32（win32com）依赖。请在命令行执行："
            "pip install pywin32，然后重试。") from e
    return win32com, pythoncom, win32gui, win32con


def _detect_cad_windows():
    """通过 COM GetActiveObject 检测所有正在运行的 CAD 实例的主窗口。

    每个 CAD 只查一次（GetActiveObject），取 app.HWND 精确获取其主窗口句柄，
    避免扫描全局窗口列表时 exe/标题匹配不可靠的问题。

    Returns:
        [(label, hwnd), ...], 如 [('autocad', 0x1234), ('gstar', 0x5678)]。
        label = 'autocad' / 'gstar'。检测不到返回空列表。
    """
    win32com, _, win32gui, _ = _win32()
    results = []
    for label, progids in [
        ("gstar", _list_cad_progids("GCAD.Application")
         + _list_cad_progids("GstarCAD.Application")),
        ("autocad", _list_cad_progids("AutoCAD.Application")),
    ]:
        for progid in progids:
            try:
                gc = win32com.client.GetActiveObject(progid)
                hwnd = int(getattr(gc, "HWND", 0))
                if hwnd and win32gui.IsWindow(hwnd):
                    results.append((label, hwnd))
                    break  # 该 CAD 已找到
            except Exception:
                continue
    return results


def foreground_cad_prefer(self_hwnd=None):
    """自动模式下，判断用户正在操作的 CAD。

    策略：通过 COM 获取所有正在运行的 CAD 主窗口句柄，再比较它们的 Z 序位置：
    - 仅一个 CAD 在跑 → 返回它。
    - 两个都在跑 → 从 Z 序顶层往下找，先出现的那个即用户刚在操作的那个
      （被 GUI 弹出前最后一次带到前台的 CAD）。
    - 全没检测到 → 返回 None。

    返回 'autocad' / 'gstar' / None。
    """
    try:
        _, _, win32gui, win32con = _win32()
        instances = _detect_cad_windows()
        if not instances:
            return None
        if len(instances) == 1:
            # 只跑了一个 CAD，不再需要猜 Z 序——肯定是用户正在用的
            return instances[0][0]
        # 两个都在跑：按 Z 序（从上到下）对比，先出现的即最近活跃的
        hwnd_set = {h for _, h in instances}
        hwnd = win32gui.GetTopWindow(None)
        while hwnd:
            if hwnd in hwnd_set:
                for label, h in instances:
                    if h == hwnd:
                        return label
            hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
    except Exception:
        return None
    return None


def get_gcad(prefer=None):
    """连接 CAD（AutoCAD 或 GstarCAD），返回 (gc, doc, ms, app_name)。

    优先连已运行的实例（GetActiveObject），否则尝试启动程序（Dispatch）。
    当 Dispatch 因 UAC 提权失败时，回退到 subprocess 直接启动 CAD 进程，
    再用 GetActiveObject 连接——这是非管理员环境下最常见的连接失败场景。
    连接成功但无打开图纸时自动新建一个。

    prefer:
      - None (自动)：先检测正在运行的 CAD 实例——
        * 只有一个在跑 → 精准连它（避免 CLI 场景下 AutoCAD 优先却把
          GstarCAD 用户的 AutoCAD 误启动出来）；
        * 两个都在跑 / 都没跑 → AutoCAD 优先，失败兜底 GstarCAD。
      - 'autocad' / 'gstar'（下拉手动指定）：只连该 CAD，连不上即报错，
        绝不悄悄连到另一种 CAD（否则会出现"选了 GstarCAD 却连上 AutoCAD"）。
    """
    win32com, _, _, _ = _win32()
    # 动态枚举 ProgID（含版本化项），覆盖多版本 AutoCAD 共存场景：
    # 无版本别名只指向最后注册的版本，用户开着旧版实例时必须用
    # AutoCAD.Application.23 等版本化 ProgID 才能 GetActiveObject 到。
    _autocad = _list_cad_progids("AutoCAD.Application") or list(_CAD_PROGIDS["autocad"])
    _gstar = (_list_cad_progids("GCAD.Application")
              + _list_cad_progids("GstarCAD.Application")) or list(_CAD_PROGIDS["gstar"])
    if prefer == "autocad":
        progids = _autocad
    elif prefer == "gstar":
        progids = _gstar
    else:  # 自动：先探测运行实例，单实例精准连
        try:
            running = _detect_cad_windows()
            if len(running) == 1:
                progids = _autocad if running[0][0] == "autocad" else _gstar
            else:
                progids = _autocad + _gstar
        except Exception:
            progids = _autocad + _gstar

    # 每次连接前清空进程检测缓存，避免进程状态过期
    _proc_running_cache.clear()

    last_err = None

    # 阶段 1：GetActiveObject 遍历全部 ProgID，优先连「已运行的实例」。
    # ⚠️ 绝不能逐 ProgID「先 GetActiveObject 再 Dispatch」：无版本别名
    # AutoCAD.Application 只指向最后注册版本(2022)，用户开着旧版(2020)时
    # 其 GetActiveObject 失败会立刻 Dispatch 冷启动 2022（等 20 秒且连错版本），
    # 根本轮不到 .23.1 去命中已运行的 2020。必须先找全运行实例，再考虑启动。
    gc = None
    for progid in progids:
        try:
            gc = win32com.client.GetActiveObject(progid)
            break
        except Exception as e:
            last_err = e
            gc = None

    # 阶段 2：没有任何运行实例 → 按版本新→旧 Dispatch 启动
    if gc is None:
        for progid in progids:
            try:
                gc = win32com.client.Dispatch(progid)
                break
            except Exception as e:
                last_err = e
                # Dispatch 因 UAC 提权失败 → 回退到 subprocess 直接启动
                if _is_elevation_error(e):
                    gc = _launch_cad_and_connect(progid)
                    if gc is not None:
                        break

    if gc is not None:
        try:
            gc.Visible = True
            doc = gc.ActiveDocument
            ms = doc.ModelSpace
            return gc, doc, ms, _cad_label(progid)
        except Exception as e:
            last_err = e
            # CAD 已连接但无打开图纸 → 尝试新建一个
            try:
                doc = gc.Documents.Add()
                ms = doc.ModelSpace
                return gc, doc, ms, _cad_label(progid)
            except Exception:
                # 新建失败 → 记录"需要打开图纸"的诊断
                label = _cad_label(progid)
                last_err = RuntimeError(
                    f'{label} 已连接但无打开图纸，请在 CAD 中新建或打开图纸。')
                gc = None

    # 所有 ProgID 都失败 → 收集场景诊断（仅失败路径，tasklist 有缓存）
    diag = _collect_connect_diag(progids)
    err_type, err_tip = _classify_com_error(last_err)
    detail = f"{last_err}" if last_err else "未知错误"
    parts = [f"{err_type}：{detail}", err_tip]
    if err_type != "CAD 未注册":
        # 权限/服务器等场景：附加 RUNASADMIN / 无图纸 / 未运行 的场景诊断
        parts.append(_build_connect_error(prefer, last_err, diag))
    else:
        # 类未注册场景：附加注册表扫描结果
        parts.append(diagnose_cad_connection(prefer))
    raise RuntimeError("\n\n".join(parts))


def _build_connect_error(prefer, last_err, diag):
    """根据诊断信息生成精确的连接失败错误消息。"""
    lines = [f'无法连接 CAD（{prefer or "自动"}）：{last_err}']

    for progid, exe_path, is_admin, is_running in diag:
        label = _cad_label(progid)
        if is_admin and is_running:
            lines.append(
                f'\n⚠ {label} 正以管理员身份运行，导致 COM 连接被隔离。\n'
                '  解决方法（任选其一）：\n'
                '  1. 关闭 CAD，右键 CAD 快捷方式 → 属性 → 兼容性 →\n'
                '     取消勾选「以管理员身份运行此程序」，然后重新启动 CAD；\n'
                '  2. 以管理员身份运行本程序（右键 → 以管理员身份运行）。'
            )
        elif is_admin and not is_running:
            lines.append(
                f'\n⚠ {label} 被设为「始终以管理员身份运行」（RUNASADMIN）。\n'
                '  非管理员进程无法通过 COM 启动它。\n'
                '  解决方法：右键 CAD 快捷方式 → 属性 → 兼容性 →\n'
                '     取消勾选「以管理员身份运行此程序」。'
            )
        elif is_running and not is_admin:
            lines.append(
                f'\n⚠ {label} 进程在运行但 COM 连接失败。\n'
                '  可能原因：CAD 尚未打开图纸（停在启动页），\n'
                '  请在 CAD 中新建或打开一个图纸后重试。'
            )
        else:
            lines.append(
                f'\n⚠ {label} 未运行且无法自动启动。\n'
                '  请手动启动 CAD 并打开一个图纸后重试。'
            )

    return '\n'.join(lines)


def get_gcad_or_elevate(prefer=None):
    """连接 CAD，检测到管理员权限不匹配时自动提权重启。

    流程：
      1. 先尝试 get_gcad(prefer) 连接
      2. 如果失败且错误与管理员权限有关（AutoCAD 以管理员运行但脚本不是）：
         a. 当前不是管理员 → 调用 _try_auto_elevate() 弹 UAC 确认
            - 用户同意 → 新管理员进程启动，当前进程退出
            - 用户拒绝 → 抛出原始错误
         b. 当前已是管理员 → 直接抛出错误（提权也解决不了）
      3. 其他错误直接抛出

    这样用户只需保持 AutoCAD 以管理员运行，脚本会自动弹 UAC 提权，
    不需要手动右键"以管理员身份运行"。
    """
    try:
        return get_gcad(prefer)
    except RuntimeError as e:
        err_msg = str(e)
        # 检测是否是管理员权限不匹配导致的连接失败（两种场景）：
        #   1. CAD 以管理员运行但脚本不是 → "COM 连接被隔离"
        #   2. CAD 设为 RUNASADMIN 但未运行 → "非管理员进程无法通过 COM 启动"
        is_admin_mismatch = (
            ('管理员身份运行' in err_msg and 'COM 连接被隔离' in err_msg)
            or ('RUNASADMIN' in err_msg and '非管理员进程无法通过 COM 启动' in err_msg)
        )
        if not is_admin_mismatch:
            raise  # 非权限问题，直接抛出

        if _is_running_as_admin():
            raise  # 已经是管理员了还连不上，问题在别处

        # 尝试自动提权
        print(
            '⚠ 检测到 CAD 需要管理员权限，正在请求提权\n'
            '  （请点击 UAC 确认框中的"是"）…',
            flush=True,
        )
        if _try_auto_elevate():
            # UAC 确认成功，新管理员进程已启动，当前进程退出
            import sys
            sys.exit(0)
        else:
            # 用户拒绝了 UAC 或提权失败
            raise RuntimeError(
                f'{err_msg}\n\n'
                '  自动提权被取消。请手动以管理员身份运行本程序：\n'
                '  右键终端/IDE → 以管理员身份运行，然后重新执行命令。'
            ) from e


def ensure_layer(doc, name, color=None):
    """确保图层存在（不存在则创建）；color 给定时设为该 ACI 颜色。"""
    try:
        layer = doc.Layers.Item(name)
    except Exception:
        layer = doc.Layers.Add(name)
    if color is not None:
        layer.Color = color


# ─── 交互拾取 ───────────────────────────────────────────────────────

def pick_point(doc, prompt):
    """在 CAD 中交互拾取一个点，返回 (x, y, z)。

    ⚠️ AutoCAD 的 GetPoint 不接受 Python None 作为基点（报
    "参数 Point 无效"），必须用 pythoncom.Empty 表示「无基点」。
    GstarCAD 两者皆可，统一用 Empty 最稳。
    """
    _, pythoncom, _, _ = _win32()
    p = doc.Utility.GetPoint(pythoncom.Empty, prompt)
    return (float(p[0]), float(p[1]), float(p[2]))


def pick_entity(doc, prompt, app_name=None):
    """在 CAD 中交互拾取一个图元，返回 COM 实体对象。

    GstarCAD 与 AutoCAD 2020 的 GetEntity 签名一致：
    GetEntity(Object, Point, [Prompt])，Object/Point 都是必填 ByRef 占位参数。
    必须传 (None, None, prompt)：
      - 只传 prompt → AutoCAD 把 prompt 字符串误当 Object 参数，报
        "The Python instance can not be converted to a COM object"；
      - 传 pythoncom.Empty 占位 → 报 "非选择性的参数"。
    两者统一用 None 占位最稳。app_name 仅保留兼容，不再分支。
    """
    entity, _ = doc.Utility.GetEntity(None, None, prompt)
    return entity


def pick_angle(doc, prompt):
    """在 CAD 命令行交互输入一个实数角度（度），返回 float。"""
    a = doc.Utility.GetReal(prompt)
    return float(a)


# ─── 图元提取 ───────────────────────────────────────────────────────

def extract_points(entity):
    """从 CAD 图元提取三维点列（用于距离计算）。"""
    name = entity.EntityName
    if name == "AcDbLine":
        return [tuple(entity.StartPoint), tuple(entity.EndPoint)]
    if name == "AcDb3dPolyline":
        c = list(entity.Coordinates)
        return [(c[i], c[i + 1], c[i + 2]) for i in range(0, len(c), 3)]
    if name == "AcDbPolyline":  # 轻量多段线 (LW)
        c = list(entity.Coordinates)
        z = getattr(entity, "Elevation", 0.0) or 0.0
        return [(c[i], c[i + 1], z) for i in range(0, len(c), 2)]
    if name == "AcDbArc":
        ctr = tuple(entity.Center)
        r = float(entity.Radius)
        a1, a2 = float(entity.StartAngle), float(entity.EndAngle)
        n = max(24, int(abs(a2 - a1) / (math.pi / 90)))
        return [(ctr[0] + r * math.cos(t), ctr[1] + r * math.sin(t), ctr[2])
                for t in [a1 + (a2 - a1) * k / n for k in range(n + 1)]]
    if name == "AcDbCircle":
        ctr = tuple(entity.Center)
        r = float(entity.Radius)
        n = 72
        return [(ctr[0] + r * math.cos(t), ctr[1] + r * math.sin(t), ctr[2])
                for t in [2 * math.pi * k / n for k in range(n + 1)]]
    if name == "AcDbSpline":
        for meth in ("GetFitPoints", "GetControlPoints"):
            try:
                c = list(getattr(entity, meth)())
                return [(c[i], c[i + 1], c[i + 2]) for i in range(0, len(c), 3)]
            except Exception:
                continue
        raise ValueError("样条线无法读取拟合点，请先转为多段线。")
    raise ValueError(f"暂不支持的图元类型: {name}")


def get_entity_by_handle(ms, handle):
    """按句柄在 ModelSpace 中查找图元 (GstarCAD 的 HandleToObject 不稳，改用遍历)。"""
    handle = str(handle).upper()
    for i in range(ms.Count):
        e = ms.Item(i)
        try:
            if str(e.Handle).upper() == handle:
                return e
        except Exception:
            continue
    raise ValueError(f"找不到句柄为 {handle} 的图元")


# ─── 绘制 ───────────────────────────────────────────────────────────

def _variant(coords):
    """构造 VT_ARRAY|VT_R8 VARIANT（COM 三维点/坐标数组通用）。"""
    win32com, pythoncom, _, _ = _win32()
    arr = pythoncom.VT_ARRAY | pythoncom.VT_R8
    return win32com.client.VARIANT(arr, coords)


def draw_catenary(doc, ms, points_3d, color=1):
    """用 Add3DPoly 绘制三维悬链线（红，图层'悬链线'），返回实体。"""
    ensure_layer(doc, LAYER_NAME)
    poly = ms.Add3DPoly(_variant([v for p in points_3d for v in p]))
    poly.Color = color
    poly.Layer = LAYER_NAME
    return poly


def draw_distance_line(doc, ms, p1, p2, color=3):
    """绘制两点间的空间连线（绿），置于图层'最小距离示意'，返回实体。"""
    ensure_layer(doc, LAYER_DISTANCE, color)
    line = ms.AddLine(_variant(list(p1)), _variant(list(p2)))
    line.Color = color
    line.Layer = LAYER_DISTANCE
    return line


def draw_distance_label(doc, ms, p1, p2, text, height=2.5, color=3):
    """在距离连线中点旁放置单个文字标注 (AcDbText)，置于图层'最小距离示意'。

    注：GstarCAD 的 AddMText 第二参数实为宽度而非字高，且设置 Attachment
    抛异常时会误触发回退、创建第二个文字。为"只标一个数字"，直接用
    AddText（height 即字高，明确可靠，单次只建一个）。
    """
    ensure_layer(doc, LAYER_DISTANCE, color)
    mx = (p1[0] + p2[0]) / 2.0
    my = (p1[1] + p2[1]) / 2.0
    mz = (p1[2] + p2[2]) / 2.0
    # 沿水平法向偏移，避免压在连线上
    dx, dy = p2[0] - p1[0], p2[1] - p1[1]
    lxy = math.hypot(dx, dy)
    nx, ny = (-dy / lxy, dx / lxy) if lxy > 1e-6 else (1.0, 0.0)
    off = height * 1.5
    px, py, pz = mx + nx * off, my + ny * off, mz
    t = ms.AddText(text, _variant([px, py, pz]), height)
    t.Color = color
    t.Layer = LAYER_DISTANCE
    return t


def draw_wind_curve(doc, ms, points_3d, color=COLOR_WIND, layer=LAYER_WIND):
    """用 Add3DPoly 绘制风偏弧垂曲线（真三维，Z 不丢），置于蓝色图层，返回实体。"""
    ensure_layer(doc, layer, color)
    poly = ms.Add3DPoly(_variant([v for p in points_3d for v in p]))
    poly.Color = color
    poly.Layer = layer
    return poly


def draw_wind_label(doc, ms, points_3d, angle_deg, color=COLOR_WIND, height=2.5):
    """在新曲线旁放置单个文字标注 '风偏X°'。

    标注**忽略输入角的正负号**，始终按绝对值标示（如输入 -30.6 也标 风偏30.6°）。
    用 AddText（height 即字高，单次单建，避免 MText 回退双创建的问题）。
    """
    ensure_layer(doc, LAYER_WIND, color)
    mid = points_3d[len(points_3d) // 2]
    mx, my, mz = mid
    # 沿水平法向偏移，避免压在曲线上
    p1, p2 = points_3d[0], points_3d[-1]
    dx, dy = p2[0] - p1[0], p2[1] - p1[1]
    lxy = math.hypot(dx, dy)
    nx, ny = (-dy / lxy, dx / lxy) if lxy > 1e-6 else (1.0, 0.0)
    off = height * 1.5
    px, py, pz = mx + nx * off, my + ny * off, mz
    text = f"风偏{abs(angle_deg)}°"  # 绝对值，忽略正负号
    t = ms.AddText(text, _variant([px, py, pz]), height)
    t.Color = color
    t.Layer = LAYER_WIND
    return t, text


def draw_bundle_line(doc, ms, points_3d, color=COLOR_BUNDLE, layer=LAYER_BUNDLE):
    """用 Add3DPoly 绘制一条分裂导线子曲线（统一黄色），返回实体。

    先 ensure_layer 创建/设置图层（图层不存在时直接赋 poly.Layer 会报
    「未找到主键」），与 draw_wind_curve 保持一致。
    """
    ensure_layer(doc, layer, color)
    poly = ms.Add3DPoly(_variant([v for p in points_3d for v in p]))
    poly.Color = color
    poly.Layer = layer
    return poly


def draw_suspension_line(doc, ms, top, bottom, color=COLOR_SUSPENSION,
                         layer=LAYER_SUSPENSION):
    """绘制悬垂串直线（品红），返回实体。"""
    ensure_layer(doc, layer, color)
    line = ms.AddLine(_variant(list(top)), _variant(list(bottom)))
    line.Color = color
    line.Layer = layer
    return line
