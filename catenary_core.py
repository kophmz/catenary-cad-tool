#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
catenary_core — 悬链线纯算法模块（零 CAD/UI 依赖）
============================================================

从主程序 catenary_app.py 抽取的纯算法部分，供主程序、风偏独立程序、
辅助脚本统一 import，消除三处代码重复（审查意见 Opt 2，漂移根源）。

包含：
  - 悬链线计算：catenary_point / solve_catenary_2d / catenary_3d
    （n_pts 全链路透传——修复 Bug 2；不等高弧垂=弦中点定义——修复 Bug 3；
     牛顿迭代收敛保护——Opt 4）
  - 二维断面悬链线：catenary_section_2d（图纸比例换算 + 反算实际尺寸，v0.24）
  - 弧垂最低点倒三角标记：low_point_marker（顶点在最低点、三角形在曲线上方，
    尺寸随图纸比例等比缩放，v0.24）
  - 曲线距离：closest_distance（AABB 快速排斥早退——Opt 3）
  - 刚体旋转：rotate_curve（轴向量/cos/sin 不变量只算一次——Opt 1，
    经 _rotate_precompute 预计算内联复用）；rotate_point（单点旋转，
    每点重算轴不变量，旋转整条曲线请用 rotate_curve）
  - 分裂导线：bundle_offsets / generate_bundle_lines
  - 悬垂串：suspension_geometry

角度约定（全部模块统一，与主程序一致）：
  右偏为正(+)、左偏为负(−)；"左右"以点选先后顺序为前进方向
  （面向 P1→P2 观察，右手边=右）。绕弦旋转用 theta = -radians(angle)
  （右手正旋把弧垂转向观察者左侧，故取负）。
"""

import math
import warnings

__all__ = [
    "catenary_point", "solve_catenary_2d", "catenary_3d", "catenary_section_2d",
    "low_point_marker", "MARKER_BASE_W", "MARKER_BASE_H",
    "MARKER_RATIO_REF_X", "MARKER_RATIO_REF_Y",
    "closest_distance", "rotate_point", "rotate_curve",
    "bundle_offsets", "generate_bundle_lines", "BUNDLE_HORIZ",
    "BUNDLE_VERT", "BUNDLE_QUAD", "suspension_geometry",
    "format_gstarcad_polyline", "format_xyz_csv", "print_summary",
]

# 弧垂最低点「倒三角」标记的基准尺寸：**默认图纸比例**（横 1:5000 / 纵 1:500）
# 下顶部边长 2 个绘图单位、高 2 个绘图单位。比例变化时按比例分母等比放大/缩小，
# 例：1:2500 / 1:250（分母都减半）→ 顶边 4、高 4（都翻倍）。
MARKER_RATIO_REF_X = 5000.0   # 基准横向比例分母
MARKER_RATIO_REF_Y = 500.0    # 基准纵向比例分母
MARKER_BASE_W = 2.0           # 基准比例下的顶部边长（绘图单位）
MARKER_BASE_H = 2.0           # 基准比例下的高度（绘图单位）


# ───────────────────────────────────────────────────────────────────
# 1) 悬链线数学 (基于《电力工程设计手册 架空输电线路设计》式14-1)
# ───────────────────────────────────────────────────────────────────

def catenary_point(x: float, a: float, x0: float, y0: float) -> float:
    """标准悬链线方程 y = a·cosh((x-x0)/a) + y0

    Args:
        x: 水平位置 (m)
        a: 悬链线参数 a = σ/γ = 1/(8K) (m)
        x0: 最低点水平位置 (m)
        y0: 最低点 y 偏移 (m)
    Returns:
        y 坐标 (m)
    """
    return a * math.cosh((x - x0) / a) + y0


def solve_catenary_2d(L: float, h: float, K: float, n_pts: int = 200):
    """求解二维悬链线（支持两端等高 / 不等高）。

    Args:
        L: 水平档距 (m) —— 两悬挂点水平投影距离
        h: 高差 (m) —— 右端悬挂点高度 - 左端悬挂点高度
        K: 弧垂 K 值 (1/m), K = γ/(8σ)
        n_pts: 采样点数（默认 200）

    Returns:
        (a, x0, y0, s_vals, z_vals) 元组
        - a: 悬链线参数 (m)
        - x0: 悬链线最低点 x 坐标 (相对左端点)
        - y0: 悬链线最低点 y 坐标
        - s_vals: 距左端点的水平距离数组 (m)
        - z_vals: 相对左端点的竖直偏移 (左端=0, 右端=h)
    """
    a = 1.0 / (8.0 * K)
    if abs(h) < 1e-6:
        # 等高悬链线（对称）：最低点在跨中
        x0 = L / 2.0
        y0 = -a * math.cosh(x0 / a)
    else:
        # 不等高：牛顿法求解 a·[cosh((L-x0)/a) - cosh(x0/a)] = h
        x0 = L / 2.0
        converged = False
        for _ in range(50):
            fx = a * (math.cosh((L - x0) / a) - math.cosh(x0 / a)) - h
            if abs(fx) < 1e-10:
                converged = True
                break
            dfx = -math.sinh((L - x0) / a) - math.sinh(x0 / a)
            if abs(dfx) < 1e-15:
                break  # 导数退化，无法继续
            x0 -= fx / dfx
        if not converged:
            warnings.warn(
                f"牛顿迭代未收敛 (L={L:.1f}m, h={h:.1f}m, K={K:.2e} 1/m)，"
                "结果可能不准确，请检查输入参数。", RuntimeWarning)
        y0 = -a * math.cosh(x0 / a)
    s_vals = [i * L / (n_pts - 1) for i in range(n_pts)]
    z_vals = [catenary_point(s, a, x0, y0) for s in s_vals]
    return a, x0, y0, s_vals, z_vals


def catenary_3d(P1, P2, K: float, n_pts: int = 200):
    """计算三维空间中的悬链线点集。

    Args:
        P1, P2: (x, y, z) 两悬挂点 (m)
        K: 弧垂 K 值 (1/m)
        n_pts: 采样点数（默认 200）

    Returns:
        (points_3d, sag_max, L, a, x0)
        - points_3d: [(x,y,z), ...] 三维点列表
        - sag_max: 最大弧垂 (m) —— 弦中点定义：
          两端连线中点沿 Z 轴竖直向下直线与曲线的交点 到中点 的距离
          （= midspan 处 |弦中点 z − 曲线点 z|；不等高时同样正确）
        - L: 水平档距 (m)
        - a: 悬链线参数 (m)
        - x0: 最低点 x 偏移 (m)

    Raises:
        ValueError: 两悬挂点水平投影重合
    """
    x1, y1, z1 = P1
    x2, y2, z2 = P2
    dx, dy = x2 - x1, y2 - y1
    L = math.hypot(dx, dy)   # 水平档距 (m)
    h = z2 - z1              # 高差 (m)
    if L < 1e-6:
        raise ValueError("两悬挂点水平投影重合，无法确定档距方向。")

    a, x0, y0, s_vals, z_vals = solve_catenary_2d(L, h, K, n_pts)

    # 映射回三维：水平沿 P1→P2 直线插值，竖直 = z1 + z_rel
    points_3d = []
    for s, z_rel in zip(s_vals, z_vals):
        t = s / L
        points_3d.append((x1 + t * dx, y1 + t * dy, z1 + z_rel))

    # 弧垂（用户定义）：跨中处「弦中点 z」与「曲线点 z」之差的绝对值。
    # 审查 F3：曲线点按**精确水平跨中** s=L/2 取值（与弦同站位），
    # 避免偶采样点数下「索引中点」≠「精确跨中」导致读数偏小
    # （默认 n=200 时旧实现读数点 s=150.75m，偏小约 0.13m / 1.7%）。
    # 注：此定义不适用于风偏旋转后的曲线（风偏功能不输出弧垂）。
    z_mid_chord = (z1 + z2) / 2.0
    z_mid_curve = z1 + catenary_point(L / 2.0, a, x0, y0)
    sag_max = abs(z_mid_chord - z_mid_curve)
    return points_3d, sag_max, L, a, x0


def catenary_section_2d(P1, P2, K: float, ratio_x: float = 5000.0,
                        ratio_y: float = 500.0, n_pts: int = 200):
    """二维断面悬链线：图纸坐标 ↔ 实际米数换算 → K 值悬链线 → 缩放回图。

    适用场景：输电线路纵断面图（二维平面 XY）。图纸绘图单位 mm、
    实际工程尺寸 m 时，出图比例 → 缩放系数为

        scale = 1000 / 比例分母
        横向 1:5000 → scale_x = 0.2  （图上 = 实际 × 1/5）
        纵向 1:500   → scale_y = 2.0  （图上 = 实际 × 2）

    即「图上坐标 = 实际坐标 × 1000 ÷ 比例分母」。

    为什么必须"先反算、再正算"：K 值 (1/m) 只有作用在**实际档距**上才有
    物理意义；若直接拿图上坐标当米用，档距会被缩小 scale_x 倍，画出来的
    线几乎成了直线（K 值输入等于白做）。故数据流为：
        ① 图上两点 P1、P2
        ② 反算实际：L_real = |Δx| / scale_x，h_real = Δy / scale_y
        ③ solve_catenary_2d(L_real, h_real, K, n_pts)  ← 算法完全复用
        ④ 缩放回图：t = s / L_real
                    x = x1 + t·(x2−x1)      （线性插值，自动兼容反向点选）
                    y = y1 + z_rel · scale_y（z_rel 通常为负 ⇒ 曲线向下垂）
    ② 与 ④ 互逆 ⇒ 曲线两端精确落在 P1、P2 上。

    二维约定：拾取点的 X = 档距方向（水平），Y = 高程方向（即原三维的 Z
    转为二维的 Y），拾取点的 Z 被忽略（始终画在 Z=0 平面）。

    弧垂最低点：曲线始终包含弧垂最低点。最低点位于水平位置 s = x0（由
    solve_catenary_2d 给出）。落在两挂点之间时曲线本就覆盖它；落在两挂点
    之外时（大高差下坡档，x0 < 0 或 x0 > L_real）自动把曲线延伸至最低点。
    延伸后两挂点仍为采样锚点 ⇒ 曲线仍精确经过用户点选的两点。

    最低点标记：同时在最低点处生成一个**倒三角**（顶点在最低点、三角形在
    曲线上方）的点列 info["marker_pts"]，尺寸随图纸比例缩放，见 low_point_marker。

    Args:
        P1, P2: 图上两点 (x, y[, z])，只取前两维
        K: 弧垂 K 值 (1/m)，调用方需已乘 1e-5
        ratio_x: 横向出图比例分母（默认 5000，即 1:5000）
        ratio_y: 纵向出图比例分母（默认 500，即 1:500）
        n_pts: 采样点数（默认 200）

    Returns:
        (points_2d, info)
        - points_2d: [(x, y), ...] 图纸坐标点列（Z 由绘图端置 0）
        - info: dict —— 含
            L_real   实际水平档距 (m)
            h_real   实际高差 (m)（右端点高程 − 左端点高程）
            sag_real 实际弧垂 (m)（弦中点定义，与 catenary_3d 同口径）
            sag_plot 图上弧垂（绘图单位）
            a        悬链线参数 a (m)
            scale_x / scale_y / ratio_x / ratio_y / n_pts
            x0       最低点距 P1 的实际水平距离 (m)（负值＝在 P1 外侧）
            z_low    最低点相对 P1 的实际高差 (m)（恒为负）
            low_plot 最低点图上坐标 (x, y)
            extended 最低点在两挂点之外、曲线已延伸至最低点
            low_in_span 最低点落在两挂点之间
            n_pts_used 实际采样点数（延伸时会加密）
            marker_pts 最低点倒三角标记的闭合点列（顶点＝最低点，见
                      low_point_marker）；marker_w / marker_h 为其顶边与高

    Raises:
        ValueError: 两点图上水平投影重合（|Δx| ≈ 0，档距方向无法确定）
        ValueError: ratio_x / ratio_y 非正数
    """
    if ratio_x <= 0 or ratio_y <= 0:
        raise ValueError("图纸比例分母必须为正数。")
    x1, y1 = float(P1[0]), float(P1[1])
    x2, y2 = float(P2[0]), float(P2[1])
    dx = x2 - x1
    if abs(dx) < 1e-9:
        raise ValueError("两点在图上水平投影重合（X 相同），无法确定档距方向。")

    scale_x = 1000.0 / float(ratio_x)   # 图纸单位 / 米
    scale_y = 1000.0 / float(ratio_y)

    # ② 反算实际尺寸（米）
    L_real = abs(dx) / scale_x          # 实际水平档距 (m)
    h_real = (y2 - y1) / scale_y        # 实际高差 (m)

    # ③ 现有算法（cosh + 牛顿迭代，含高差修正），一行不改
    a, x0, y0, s_vals, z_vals = solve_catenary_2d(L_real, h_real, K, n_pts)

    # ③-补 弧垂最低点：悬链线最低点位于水平位置 s = x0（求解器返回）。
    #       若最低点落在两挂点之外（x0 < 0 或 x0 > L_real，常见于大高差
    #       下坡档），把采样区间延伸到最低点，使画出的曲线包含弧垂最低处；
    #       延伸只扩展采样范围，曲线方程不变 ⇒ 原区间内形状与数值完全不变。
    #       采样时把「两挂点 s=0 / s=L_real」与「最低点 s=x0」作为锚点精确
    #       插入序列：挂点必须精确落在曲线上（否则均匀步长变化会让拾取点
    #       偏离），最低点也必须是曲线上的点（均匀采样时最低点通常落在两个
    #       采样点之间，见下）。
    extend_lo = x0 < 0.0
    extend_hi = x0 > L_real
    extended = extend_lo or extend_hi
    if extended:
        s_lo = x0 if extend_lo else 0.0
        s_hi = x0 if extend_hi else L_real
        span = s_hi - s_lo
        n_used = max(n_pts, int(math.ceil(n_pts * span / L_real)))  # 按跨度加密
        s_vals = [s_lo + i * span / (n_used - 1) for i in range(n_used)]
    else:
        s_vals = [i * L_real / (n_pts - 1) for i in range(n_pts)]
    # 锚点（挂点 + 最低点）精确并入，升序去重
    s_vals = sorted(set(s_vals + [0.0, L_real, x0]))
    z_vals = [catenary_point(s, a, x0, y0) for s in s_vals]

    # ④ 缩放回图纸坐标（s / z_rel 均以 P1 为原点）
    points_2d = []
    for s, z_rel in zip(s_vals, z_vals):
        t = s / L_real
        points_2d.append((x1 + t * dx, y1 + z_rel * scale_y))

    # 弧垂（与 catenary_3d 同口径）：精确水平跨中处「弦」与「曲线」之高差。
    # z_vals 以左端点为 0，故弦在跨中为 h_real/2。
    z_mid_chord = h_real / 2.0
    z_mid_curve = catenary_point(L_real / 2.0, a, x0, y0)
    sag_real = abs(z_mid_chord - z_mid_curve)

    info = {
        "L_real": L_real,
        "h_real": h_real,
        "sag_real": sag_real,
        "sag_plot": sag_real * scale_y,
        "a": a,
        "n_pts": n_pts,
        "n_pts_used": len(s_vals),
        "scale_x": scale_x,
        "scale_y": scale_y,
        "ratio_x": float(ratio_x),
        "ratio_y": float(ratio_y),
        # —— 弧垂最低点 ——
        # x0         最低点距 P1 的实际水平距离 (m)（负值＝在 P1 外侧）
        # z_low      最低点相对 P1 的实际高差 (m)（恒为负 → 比 P1 低）
        #            注意：曲线方程 z(s)=a·cosh((s-x0)/a)+y0 的最低值是
        #            z(x0) = a + y0（y0 是「顶点的 y 截距」，不是最低点高度）
        # low_plot   最低点的图上坐标 (x, y)
        # extended   最低点原本在两挂点之外、曲线已延伸至最低点
        # low_in_span 最低点落在两挂点水平区间内（曲线本就包含最低点）
        "x0": x0,
        "z_low": a + y0,
        "low_plot": (x1 + (x0 / L_real) * dx, y1 + (a + y0) * scale_y),
        "extended": extended,
        "low_in_span": not extended,
    }
    # —— 最低点倒三角标记：顶点＝最低点，三角形在曲线上方；尺寸随比例缩放 ——
    info["marker_pts"] = low_point_marker(info["low_plot"], ratio_x, ratio_y)
    info["marker_w"] = MARKER_BASE_W * (MARKER_RATIO_REF_X / float(ratio_x))
    info["marker_h"] = MARKER_BASE_H * (MARKER_RATIO_REF_Y / float(ratio_y))
    return points_2d, info


def low_point_marker(low_plot, ratio_x: float = MARKER_RATIO_REF_X,
                     ratio_y: float = MARKER_RATIO_REF_Y,
                     base_w: float = MARKER_BASE_W,
                     base_h: float = MARKER_BASE_H):
    """弧垂最低点「倒三角」标记的几何点列（纯几何，零 CAD 依赖）。

    形状：等腰三角形**倒置**，**顶点（尖端）落在弧垂曲线的最低点**，三角形
    整体位于曲线**上方**（底边在上、水平）。

        apex (顶点) = low_plot                ← 曲线最低点
        右上       = apex + ( +w/2, +h )
        左上       = apex + ( −w/2, +h )

    尺寸随图纸比例缩放（同一份图里标记的"实际大小"恒定，缩放后出图比例变了
    标记也跟着变）：基准比例（横 1:5000 / 纵 1:500）下顶边 2、高 2 个绘图单位

        w = base_w × 基准横向比例分母 / ratio_x    （1:5000→2，1:2500→4）
        h = base_h × 基准纵向比例分母 / ratio_y    （1:500 →2，1:250 →4）

    Args:
        low_plot: 最低点图上坐标 (x, y)（catenary_section_2d 的 info["low_plot"]）
        ratio_x / ratio_y: 当前图纸比例分母（与 catenary_section_2d 同参数）
        base_w / base_h: 基准比例下的顶边与高度（默认 2 / 2，一般不改）

    Returns:
        [(apex), (右上), (左上), (apex)] —— **闭合点列**（首尾同点），可直接
        交给 draw_catenary_2d / AddLightWeightPolyline 画成三角形线框。
    """
    x, y = float(low_plot[0]), float(low_plot[1])
    w = float(base_w) * (MARKER_RATIO_REF_X / float(ratio_x))
    h = float(base_h) * (MARKER_RATIO_REF_Y / float(ratio_y))
    half = w / 2.0
    return [(x, y), (x + half, y + h), (x - half, y + h), (x, y)]


# ───────────────────────────────────────────────────────────────────
# 2) 曲线采样 / 最小距离
# ───────────────────────────────────────────────────────────────────

def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _dist(a, b):
    return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2)


def _clamp(v, lo, hi):
    return lo if v < lo else (hi if v > hi else v)


def _seg_seg(p1, p2, p3, p4):
    """两线段 (p1-p2) 与 (p3-p4) 之间的最近点距离。"""
    d1 = (p2[0] - p1[0], p2[1] - p1[1], p2[2] - p1[2])
    d2 = (p4[0] - p3[0], p4[1] - p3[1], p4[2] - p3[2])
    r = (p1[0] - p3[0], p1[1] - p3[1], p1[2] - p3[2])
    a = _dot(d1, d1)
    e = _dot(d2, d2)
    f = _dot(d2, r)

    if a <= 1e-12 and e <= 1e-12:
        return _dist(p1, p3), p1, p3
    if a <= 1e-12:
        s = 0.0
        t = _clamp(f / e, 0.0, 1.0)
    else:
        c = _dot(d1, r)
        if e <= 1e-12:
            t = 0.0
            s = _clamp(-c / a, 0.0, 1.0)
        else:
            b = _dot(d1, d2)
            denom = a * e - b * b
            s = 0.0 if denom <= 1e-12 else _clamp((b * f - c * e) / denom, 0.0, 1.0)
            t = (b * s + f) / e
            if t < 0.0:
                t = 0.0
                s = _clamp(-c / a, 0.0, 1.0)
            elif t > 1.0:
                t = 1.0
                s = _clamp((b - c) / a, 0.0, 1.0)
    cp1 = (p1[0] + s * d1[0], p1[1] + s * d1[1], p1[2] + s * d1[2])
    cp2 = (p3[0] + t * d2[0], p3[1] + t * d2[1], p3[2] + t * d2[2])
    return _dist(cp1, cp2), cp1, cp2


def _aabb(points):
    """点列的轴对齐包围盒 (min_x, min_y, min_z, max_x, max_y, max_z)。"""
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    zs = [p[2] for p in points]
    return (min(xs), min(ys), min(zs), max(xs), max(ys), max(zs))


def _seg_aabb_dist(p1, p2, p3, p4):
    """两线段轴对齐包围盒之间的保守下界距离。

    真实段间距离 ≥ 该下界。用于段级快速排斥：下界已 ≥ 当前 best 时，
    该段对不可能产生更小距离，可安全跳过（不改变最终结果）。
    """
    # 每段各轴区间 [min, max]
    a0 = (min(p1[0], p2[0]), min(p1[1], p2[1]), min(p1[2], p2[2]))
    a1 = (max(p1[0], p2[0]), max(p1[1], p2[1]), max(p1[2], p2[2]))
    b0 = (min(p3[0], p4[0]), min(p3[1], p4[1]), min(p3[2], p4[2]))
    b1 = (max(p3[0], p4[0]), max(p3[1], p4[1]), max(p3[2], p4[2]))
    dx = max(0.0, b0[0] - a1[0], a0[0] - b1[0])
    dy = max(0.0, b0[1] - a1[1], a0[1] - b1[1])
    dz = max(0.0, b0[2] - a1[2], a0[2] - b1[2])
    return math.sqrt(dx * dx + dy * dy + dz * dz)


def closest_distance(ptsA, ptsB):
    """两条空间曲线 (点列) 之间的最小距离。

    对每条曲线按相邻点分段，做线段-线段最近距离搜索。
    段级快速排斥（Opt 3）：段对 AABB 下界距离已 ≥ 当前 best 时跳过，
    减少无效的 _seg_seg 计算——**只做剪枝，不改变结果**。

    Args:
        ptsA, ptsB: 三维点列 [(x,y,z), ...]（各 ≥2 个点）

    Returns:
        (min_dist, closest_on_A, closest_on_B)
    """
    best = float("inf")
    ba, bb = None, None
    for i in range(len(ptsA) - 1):
        a1, a2 = ptsA[i], ptsA[i + 1]
        for j in range(len(ptsB) - 1):
            b1, b2 = ptsB[j], ptsB[j + 1]
            # 快速排斥：两段 AABB 下界 ≥ best → 该段对不可能更近，跳过
            if _seg_aabb_dist(a1, a2, b1, b2) >= best:
                continue
            d, c1, c2 = _seg_seg(a1, a2, b1, b2)
            if d < best:
                best, ba, bb = d, c1, c2
    return best, ba, bb


# ───────────────────────────────────────────────────────────────────
# 3) 刚体旋转（绕曲线首尾端点连线，Rodrigues 公式）
# ───────────────────────────────────────────────────────────────────

def rotate_point(p, p1, p2, angle_rad):
    """将点 p 绕过点 p1、方向 (p2-p1) 的轴线刚体旋转 angle_rad（右手定则）。

    单点旋转用（含轴不变量每点重算——若旋转整条曲线请用 rotate_curve，
    那里轴/三角函数只算一次，性能更好）。

    Args:
        p: 待旋转点 (x, y, z)
        p1, p2: 轴上的两点
        angle_rad: 旋转角（弧度）
    Returns:
        旋转后的点 (x, y, z)。轴上的点（如 p1、p2）保持不变。
    """
    vx, vy, vz = p[0] - p1[0], p[1] - p1[1], p[2] - p1[2]
    ax, ay, az = p2[0] - p1[0], p2[1] - p1[1], p2[2] - p1[2]
    L = math.sqrt(ax * ax + ay * ay + az * az)
    if L < 1e-12:
        return list(p)  # 退化轴，无法旋转
    ux, uy, uz = ax / L, ay / L, az / L
    cos_t, sin_t = math.cos(angle_rad), math.sin(angle_rad)
    cx = uy * vz - uz * vy
    cy = uz * vx - ux * vz
    cz = ux * vy - uy * vx
    dot = ux * vx + uy * vy + uz * vz
    k = 1.0 - cos_t
    rx = vx * cos_t + cx * sin_t + ux * dot * k
    ry = vy * cos_t + cy * sin_t + uy * dot * k
    rz = vz * cos_t + cz * sin_t + uz * dot * k
    return (rx + p1[0], ry + p1[1], rz + p1[2])


def _rotate_precompute(p1, p2, angle_rad):
    """预计算一次旋转轴单位向量与三角函数（审查 F2）。

    返回 (p1, ux, uy, uz, cos_t, sin_t, k) 或 (p1, None)（退化轴）。
    供 rotate_curve 内联复用，避免每点重复 sqrt / cos / sin。
    """
    ax, ay, az = p2[0] - p1[0], p2[1] - p1[1], p2[2] - p1[2]
    L = math.sqrt(ax * ax + ay * ay + az * az)
    if L < 1e-12:
        return p1, None
    ux, uy, uz = ax / L, ay / L, az / L
    cos_t, sin_t = math.cos(angle_rad), math.sin(angle_rad)
    return (p1, ux, uy, uz, cos_t, sin_t, 1.0 - cos_t)


def rotate_curve(points, angle_deg):
    """将整条曲线绕其首尾端点连线旋转 angle_deg 度，返回新点列。

    不变量（轴单位向量、cos/sin）只算一次，内联旋转每个点（审查 F2）。

    Args:
        points: 原曲线三维点列 [(x,y,z), ...]，至少 2 个点。
        angle_deg: 旋转角（度）。**正 = 以起点→终点为前进方向向右偏**
                   （与模块⑤悬垂串的 A/B 角约定一致，右偏为正）。
    Returns:
        旋转后的新点列（与原曲线等长）。

    Raises:
        ValueError: 点列少于 2 个点
    """
    if len(points) < 2:
        raise ValueError("曲线点列至少需要 2 个点")
    p1, p2 = points[0], points[-1]
    # 绕弦右手正旋转把弧垂(向下)转向观察者的左侧 → 取负实现「正=右偏」
    theta = -math.radians(angle_deg)
    pre = _rotate_precompute(p1, p2, theta)
    if pre[1] is None:
        return [list(p) for p in points]  # 退化轴：整条曲线原样返回
    _, ux, uy, uz, cos_t, sin_t, k = pre
    out = []
    for p in points:
        vx = p[0] - p1[0]
        vy = p[1] - p1[1]
        vz = p[2] - p1[2]
        cx = uy * vz - uz * vy
        cy = uz * vx - ux * vz
        cz = ux * vy - uy * vx
        dot = ux * vx + uy * vy + uz * vz
        rx = vx * cos_t + cx * sin_t + ux * dot * k
        ry = vy * cos_t + cy * sin_t + uy * dot * k
        rz = vz * cos_t + cz * sin_t + uz * dot * k
        out.append((rx + p1[0], ry + p1[1], rz + p1[2]))
    return out


# ───────────────────────────────────────────────────────────────────
# 4) 分裂导线（双分裂 / 四分裂）
# ───────────────────────────────────────────────────────────────────

def _unit(v):
    """三维单位向量，退化时返回 None。"""
    L = math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])
    if L < 1e-12:
        return None
    return (v[0] / L, v[1] / L, v[2] / L)


def _cross(a, b):
    return (a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def _plane_normal(points):
    """由曲线点列拟合弦平面法线 n（逐点法线平均，无需 numpy）。

    对每个相邻三元组取局部法线（边叉积）并求和，最后归一化。
    对平面曲线（含风偏后的倾斜平面曲线）与带轻微折点噪声的
    3D 多段线均稳定。若曲线完全共线则返回 None。
    """
    acc = [0.0, 0.0, 0.0]
    cnt = 0
    for i in range(len(points) - 2):
        v1 = (points[i + 1][0] - points[i][0],
              points[i + 1][1] - points[i][1],
              points[i + 1][2] - points[i][2])
        v2 = (points[i + 2][0] - points[i + 1][0],
              points[i + 2][1] - points[i + 1][1],
              points[i + 2][2] - points[i + 1][2])
        n_ = _cross(v1, v2)
        L = math.sqrt(n_[0] * n_[0] + n_[1] * n_[1] + n_[2] * n_[2])
        if L < 1e-12:
            continue  # 局部共线，跳过
        acc[0] += n_[0] / L
        acc[1] += n_[1] / L
        acc[2] += n_[2] / L
        cnt += 1
    if cnt == 0:
        return None
    return _unit(tuple(acc))


def bundle_offsets(points, d):
    """计算分裂导线局部坐标系与各子导线偏移向量。

    局部系定义（弦平面坐标系）：
      u = 弦方向（首尾端点连线单位向量）
      n = 弦平面法线（水平分裂方向；未风偏时=水平横向）
      v = n × u（平面内垂直于弦的方向；未风偏时=竖直向下）

    Args:
        points: 中心线三维点列 [(x,y,z), ...]（≥3 个点）
        d:      分裂间距（米）

    Returns:
        (u, n, v, offsets)
        offsets: 子导线偏移向量列表（每条 = 相对中心线每点的平移量），
                 长度 = 6：
                 - 水平双分裂: offsets[0:2] (±n·d/2)
                 - 垂直双分裂: offsets[2:4] (±v·d/2)
                 - 四分裂: offsets[4:8] (±n±v)·d/2

    Raises:
        ValueError: 中心线点数不足 / 首尾重合 / 共线
    """
    if len(points) < 3:
        raise ValueError("中心线点列至少需要 3 个点")
    u = _unit((points[-1][0] - points[0][0],
               points[-1][1] - points[0][1],
               points[-1][2] - points[0][2]))
    if u is None:
        raise ValueError("中心线首尾端点重合，无法确定弦方向")
    n = _plane_normal(points)
    if n is None:
        raise ValueError("中心线为直线（共线），无法确定弦平面法线")
    v = _cross(n, u)  # 平面内 ⊥弦（未风偏时指向下方）
    half = d / 2.0
    nv = (n[0] * half, n[1] * half, n[2] * half)
    vv = (v[0] * half, v[1] * half, v[2] * half)
    return u, n, v, [
        nv, (-nv[0], -nv[1], -nv[2]),          # 水平双分裂：±n·d/2
        vv, (-vv[0], -vv[1], -vv[2]),          # 垂直双分裂：±v·d/2
        (nv[0] + vv[0], nv[1] + vv[1], nv[2] + vv[2]),   # 四分裂 1
        (nv[0] - vv[0], nv[1] - vv[1], nv[2] - vv[2]),   # 四分裂 2
        (-nv[0] + vv[0], -nv[1] + vv[1], -nv[2] + vv[2]),  # 四分裂 3
        (-nv[0] - vv[0], -nv[1] - vv[1], -nv[2] - vv[2]),  # 四分裂 4
    ]


BUNDLE_HORIZ = 0   # 水平双分裂（取 offsets[0:2]）
BUNDLE_VERT = 2    # 垂直双分裂（取 offsets[2:4]）
BUNDLE_QUAD = 4    # 四分裂（取 offsets[4:8]）


def generate_bundle_lines(points, d, mode):
    """生成分裂导线各子导线点列。

    Args:
        points: 中心线三维点列（米）
        d:      分裂间距（米）
        mode:   BUNDLE_HORIZ / BUNDLE_VERT / BUNDLE_QUAD

    Returns:
        [sub_line_1, sub_line_2, ...]，每个 sub_line 与中心线等长同构
        （即中心线整体平移，子导线处处平行、间距恒为 d）。
    """
    u, n, v, offsets = bundle_offsets(points, d)
    chosen = offsets[mode:mode + (4 if mode == BUNDLE_QUAD else 2)]
    lines = []
    for off in chosen:
        lines.append([(p[0] + off[0], p[1] + off[1], p[2] + off[2]) for p in points])
    return lines


# ───────────────────────────────────────────────────────────────────
# 5) 含悬垂绝缘子串的弧垂线
# ───────────────────────────────────────────────────────────────────

def _susp_lower_with_right(top, right, D, A_deg):
    """按给定的横向单位向量 right 计算悬垂串下端。

    下端 = 上端 + right·D·sin(A) − z·D·cos(A)；A=0 退化为竖直串。
    """
    A = math.radians(A_deg)
    h = D * math.sin(A)   # 横向偏移量
    v = -D * math.cos(A)  # 竖直下落量
    return (top[0] + right[0] * h,
            top[1] + right[1] * h,
            top[2] + v)


def _forward_right(p1, p2):
    """按前进方向 p1→p2 计算横向单位向量 right（面向 p1→p2，右手边=右）。

    Returns:
        right 单位向量；u 与 z 平行时退化为水平 y 方向 (0,1,0)。
    """
    u = _unit((p2[0] - p1[0], p2[1] - p1[1], p2[2] - p1[2]))
    if u is None:
        raise ValueError("两端点重合，无法确定前进方向")
    right = _unit((u[1], -u[0], 0.0))
    if right is None:
        right = (0.0, 1.0, 0.0)
    return right


def suspension_geometry(p1, p2, K, D1=None, A1=0.0, D2=None, A2=0.0,
                        B_deg=0.0, n_pts=200):
    """含悬垂绝缘子串的弧垂线完整计算（基准 + 可选风偏）。

    Args:
        p1, p2: 两个选点（顺序 = 前进方向）
        K:      弧垂 K 值（1/m，已含 ×10⁻⁵）
        D1, A1: 点1 悬垂串参数（D1=None 表示该点无悬垂串）
        D2, A2: 点2 悬垂串参数（D2=None 表示该点无悬垂串）
        B_deg:  弧垂线风偏角（度，右偏为正；0 = 不绘制风偏线）
        n_pts:  采样点数

    Returns:
        dict:
          ep1, ep2        — 风偏后的串下端（无串=选点；非风偏时=A=0下端）
          ref_ep1, ref_ep2— A=0（竖直串）下端，供参照基准线使用
          susp_lines      — 风偏后的悬垂串线段（浅蓝，仅风偏模式含）
          ref_susp_lines  — A=0 竖直串线段（品红参照，仅风偏模式含）
          base_pts        — A=0 参照基准弧垂线点列（红色，始终画）
          wind_pts        — 风偏线点列（B≠0 时为风偏后的线；否则 None）
          wind_base       — 风偏线对应的中间基准线点列（仅供内部，不绘制）
          sag, L          — 参照基准弧垂、档距

    模式说明:
      非风偏（A1=A2=0 且 B=0）:
        只画竖直串(品红) + 基准弧垂(红)，wind_pts=None。
      风偏模式（A1/A2≠0 或 B≠0）:
        画摆串(浅蓝) + 风偏线(蓝)，同时另画 A=0 竖直串(品红)
        + 同 K 无风偏基准弧垂(红) 作参照；中间基准线不绘制。
    """
    # 0) 整体前进方向 = P1→P2，两侧 right 相同
    right = _forward_right(p1, p2)

    # 1) 风偏串下端（按 A1/A2 摆）与 A=0 竖直串下端（参照）
    ep1 = _susp_lower_with_right(p1, right, D1, A1) if D1 is not None else p1
    ep2 = _susp_lower_with_right(p2, right, D2, A2) if D2 is not None else p2
    ref_ep1 = _susp_lower_with_right(p1, right, D1, 0.0) if D1 is not None else p1
    ref_ep2 = _susp_lower_with_right(p2, right, D2, 0.0) if D2 is not None else p2

    wind_mode = abs(B_deg) > 1e-9 or abs(A1) > 1e-9 or abs(A2) > 1e-9

    # 2) 风偏线（仅风偏模式）：以风偏串下端为端点画中间基准 → 绕弦转 B
    wind_pts = None
    wind_base = None
    if wind_mode:
        wind_base, _, _, _, _ = catenary_3d(ep1, ep2, K, n_pts)
        if abs(B_deg) > 1e-9:
            # ⚠️ rotate_curve 内部绕首尾连线转并取负实现「正=右偏」，
            # 与模块③风偏一致（同一约定）；顺带复用 F2 的不变量一次计算
            wind_pts = rotate_curve(wind_base, B_deg)
        else:
            wind_pts = wind_base  # B=0 时风偏线 = 摆串端点的基准线本身

    # 3) 参照基准弧垂线（A=0，红色，始终绘制）
    base_pts, sag, L, a, x0 = catenary_3d(ref_ep1, ref_ep2, K, n_pts)

    # 4) 悬垂串线段
    susp_lines, ref_susp_lines = [], []
    if wind_mode:
        # 摆串（浅蓝）+ A=0 竖直串（品红参照）
        if D1 is not None:
            susp_lines.append((p1, ep1))
            ref_susp_lines.append((p1, ref_ep1))
        if D2 is not None:
            susp_lines.append((p2, ep2))
            ref_susp_lines.append((p2, ref_ep2))
    else:
        # 非风偏：竖直串（品红）即为实际串线
        if D1 is not None:
            susp_lines.append((p1, ref_ep1))
        if D2 is not None:
            susp_lines.append((p2, ref_ep2))

    return {
        "ep1": ep1, "ep2": ep2,
        "ref_ep1": ref_ep1, "ref_ep2": ref_ep2,
        "susp_lines": susp_lines,
        "ref_susp_lines": ref_susp_lines,
        "base_pts": base_pts,
        "wind_pts": wind_pts,
        "wind_base": wind_base,
        "sag": sag, "L": L,
    }


# ───────────────────────────────────────────────────────────────────
# 6) 输出辅助（纯函数，零 CAD 依赖）
# ───────────────────────────────────────────────────────────────────

def format_gstarcad_polyline(points_3d):
    """
    将三维点列表格式化为 GstarCAD MCP draw_polyline 可用的 JSON 参数。
    """
    coords = []
    for p in points_3d:
        coords.append([round(p[0], 3), round(p[1], 3), round(p[2], 3)])
    return {"points": coords}


def format_xyz_csv(points_3d, filepath: str):
    """导出为 XYZ CSV 文件"""
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write("x,y,z\n")
        for p in points_3d:
            f.write(f"{p[0]:.3f},{p[1]:.3f},{p[2]:.3f}\n")
    print(f"已导出: {filepath}")


def print_summary(P1, P2, K, sag_max, L, a, x0):
    """打印悬链线参数摘要"""
    print("=" * 55)
    print("  三维悬链线参数汇总")
    print("=" * 55)
    print(f"  K 值:               {K:.2e}  1/m")
    print(f"  悬链线参数 a:       {a:.2f}  m")
    print(f"  悬挂点 P1:          ({P1[0]:.1f}, {P1[1]:.1f}, {P1[2]:.1f})")
    print(f"  悬挂点 P2:          ({P2[0]:.1f}, {P2[1]:.1f}, {P2[2]:.1f})")
    print(f"  水平档距 L:         {L:.2f}  m")
    print(f"  高差 h:             {P2[2]-P1[2]:.2f}  m")
    print(f"  最低点偏移 x₀:      {x0:.2f}  m (距左端点)")
    print(f"  最大弧垂(弦中点定义): {sag_max:.3f}  m")
    print(f"  K×L² (估算弧垂):    {K * L * L:.3f}  m")
    print("=" * 55)
