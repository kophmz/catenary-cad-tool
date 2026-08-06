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
    for label, progids in [("gstar", _CAD_PROGIDS["gstar"]),
                            ("autocad", _CAD_PROGIDS["autocad"])]:
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

    优先连已运行的实例（GetActiveObject），否则启动对应程序（Dispatch）。
    prefer:
      - None (自动)：先检测正在运行的 CAD 实例——
        * 只有一个在跑 → 精准连它（避免 CLI 场景下 AutoCAD 优先却把
          GstarCAD 用户的 AutoCAD 误启动出来）；
        * 两个都在跑 / 都没跑 → AutoCAD 优先，失败兜底 GstarCAD。
      - 'autocad' / 'gstar'（下拉手动指定）：只连该 CAD，连不上即报错，
        绝不悄悄连到另一种 CAD（否则会出现"选了 GstarCAD 却连上 AutoCAD"）。
    """
    win32com, _, _, _ = _win32()
    if prefer == "autocad":
        progids = list(_CAD_PROGIDS["autocad"])
    elif prefer == "gstar":
        progids = list(_CAD_PROGIDS["gstar"])
    else:  # 自动：先探测运行实例，单实例精准连
        try:
            running = _detect_cad_windows()
            if len(running) == 1:
                progids = _CAD_PROGIDS[running[0][0]]
            else:
                progids = _CAD_PROGIDS["autocad"] + _CAD_PROGIDS["gstar"]
        except Exception:
            progids = _CAD_PROGIDS["autocad"] + _CAD_PROGIDS["gstar"]
    last_err = None
    for progid in progids:
        gc = None
        try:
            gc = win32com.client.GetActiveObject(progid)  # 先连已运行的实例
        except Exception as e:
            last_err = e
        if gc is None:
            try:
                gc = win32com.client.Dispatch(progid)  # 否则启动程序
            except Exception as e:
                last_err = e
                continue
        try:
            gc.Visible = True
            doc = gc.ActiveDocument
            ms = doc.ModelSpace
            return gc, doc, ms, _CAD_LABELS.get(progid, progid)
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"无法连接 CAD（{prefer or '自动'}）：{last_err}")


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
