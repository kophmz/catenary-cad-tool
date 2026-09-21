#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
空间弧垂曲线工具 (CAD 通用版 · 支持 AutoCAD / GstarCAD / 中望CAD)
============================================================

功能:
  1. 在 CAD 中拾取两个悬挂点，输入 K 值，在两点间生成 1:1（米）的三维
     空间弧垂曲线 (AcDb3dPolyline)，弧垂方向为 Z 轴。
  2. 在 CAD 中依次拾取两条曲线（或直线/圆弧/圆/样条），计算并绘制两条
     空间曲线之间的最小距离连线，并在界面给出数值。
  3. 风偏弧垂曲线（绕曲线两端点连线旋转，右偏为正）。
  4. 分裂导线（双分裂/四分裂，间距 mm）。
  5. 含悬垂绝缘子串的弧垂线（串长/串角独立，可选风偏线）。

依赖:
  - pywin32 (win32com)      —— 直连 CAD COM (AutoCAD / GstarCAD / 中望CAD)
  - tkinter (标准库)        —— 界面
  - CAD 已启动并打开图纸 (AutoCAD / GstarCAD / 中望CAD)

代码结构（重构后，消除三处重复——审查 Opt 2）:
  - catenary_core.py  —— 纯算法（悬链线/距离/旋转/分裂/悬垂串），零 CAD 依赖
  - catenary_cad.py   —— win32com 封装（连接/拾取/提取/绘制，图层颜色常量）
  - catenary_app.py   —— GUI + CLI（本文件）

K 值约定 (与《电力工程设计手册 架空输电线路设计》一致):
  输入时省略 10⁻⁵ 后缀，例如 K=15.22 表示 15.22×10⁻⁵ 1/m。
  纵 1:500 / 横 1:5000 模板的 K 值即为物理 K 值，无需换算。

运行:
  # GUI 模式 (默认): 在桌面双击 / 终端运行
  python catenary_app.py

  # 命令行模式:
  #   生成弧垂 (交互点选两点)
  python catenary_app.py draw --K 15.22 --pick
  #   生成弧垂 (直接给坐标)
  python catenary_app.py draw --K 15.22 --p1 39954 38449 33 --p2 40039 38463 39
  #   两曲线最小距离 (交互选两条)
  python catenary_app.py mindist --pick
  #   两曲线最小距离 (给句柄)
  python catenary_app.py mindist --h1 728 --h2 8C2
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

from catenary_core import (
    catenary_3d,
    catenary_section_2d,
    closest_distance,
    rotate_curve,
    generate_bundle_lines,
    suspension_geometry,
    BUNDLE_HORIZ,
    BUNDLE_VERT,
    BUNDLE_QUAD,
)
from catenary_cad import (
    # 连接
    get_gcad,
    get_gcad_or_elevate,
    foreground_cad_target,
    diagnose_cad_connection,
    # CAD 家族配置表（GUI 下拉 / CLI choices / 文案的唯一出处）
    cad_family_labels,
    cad_family_keys,
    cad_label_of,
    _all_cad_names,
    # 图层
    ensure_layer,
    # 交互
    pick_point,
    pick_entity,
    pick_angle,
    # 提取
    extract_points,
    get_entity_by_handle,
    # 绘制
    draw_catenary,
    draw_catenary_2d,
    draw_low_point_marker,
    draw_distance_line,
    draw_distance_label,
    draw_wind_curve,
    draw_wind_label,
    draw_bundle_line,
    draw_suspension_line,
    # 图层/颜色常量（全项目唯一出处）
    LAYER_NAME,
    LAYER_WIND,
    LAYER_SUSPENSION,
    LAYER_SUSPENSION_WIND,
    COLOR_SUSPENSION_WIND,
)


def _help_path():
    """返回使用说明 HTML 文件的路径（兼容源码运行和 PyInstaller 打包）。"""
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, "使用说明.html")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "使用说明.html")


def _is_cancel_error(e):
    """判断 COM 异常是否由用户取消 (ESC / 右键) 引起（GUI 与 CLI 共用同一判定）。

    ⚠️ 各 CAD 的取消文案不统一（"用户取消" / "已取消" / "Cancel" / ESC…），
    故同时匹配多组中英文关键词。调用方另有兜底：连续点选过程中只要**取点失败**
    即结束链路并保留已画曲线，不依赖本判定是否命中。
    """
    s = str(e).lower()
    return any(k in s for k in ("cancel", "取消", "esc", "userinterrupt", "discard"))


def _low_point_note(info):
    """弧垂最低点的一句话说明（GUI 日志与 CLI 共用同一文案）。"""
    low = info["low_plot"]
    dy = info["z_low"]
    rel = f"低于起点 {abs(dy):.2f}m" if dy < 0 else f"高于起点 {abs(dy):.2f}m"
    where = "位于两挂点之间" if info["low_in_span"] else "在两挂点之外，曲线已延伸至最低点"
    mk = ""
    if "marker_w" in info:
        mk = f"  倒三角标记 顶边{info['marker_w']:g}×高{info['marker_h']:g}"
    return (f"最低点: 距起点 {info['x0']:.2f}m  {rel}  "
            f"图上({low[0]:.2f}, {low[1]:.2f})（{where}）{mk}")


# ───────────────────────────────────────────────────────────────────
# GUI
# ───────────────────────────────────────────────────────────────────

class CatenaryApp:
    def __init__(self, root):
        self.root = root
        root.title("空间弧垂曲线工具 · CAD · v0.24")
        # 高度固定、**宽度随内容自适应**（见 _render_status）：长图纸名会把窗口
        # 撑宽，而不是把「连接/刷新」按钮切掉一半。
        # ⚠️ 不能用 root.resizable(False, False) 来禁止用户缩放 —— 它会把窗口的
        # 最大尺寸钉死在「设置那一刻的宽度」，之后程序再调 geometry() 加宽会被
        # Windows 直接夹回去（实测设 860 仍停在 720）。改用 min=max=当前尺寸，
        # 既禁止用户拖拽，又允许程序改宽。
        self._win_h = 1000
        self._win_min_w = 720
        # 宽度上限：超过它就不再把窗口拉宽（否则长名字会把窗口拉满整个屏幕），
        # 改为把状态文字中间省略 —— 取 1100 可显示约 50 字的图纸名。
        self._win_max_w = 1100
        self._status_pending = False
        self._rendering = False        # _render_status 重入护栏
        self._set_window_width(self._win_min_w)

        self.gc = self.doc = self.ms = None
        self.app_name = None  # 连接后记录的 CAD 名称（AutoCAD / GstarCAD / 中望CAD）
        # 已连实例的主窗口句柄：多 CAD / 同 CAD 多版本共存时，
        # 用它判断「用户是否把另一个实例切到了前台」，以便自动跟随重连
        self._conn_hwnd = None

        self._build_ui()

    # ---- 界面 -------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # 连接状态
        f0 = ttk.LabelFrame(self.root, text="连接")
        f0.pack(fill="x", **pad)
        self._conn_frame = f0  # 供 _conn_bar_extra 计算宽度占用
        # ⚠️ pack 是按调用顺序分配空间的：必须**先 pack 右侧按钮、再 pack 状态文字**。
        # 顺序反过来时，长图纸名会先把空间吃满（实测「已连接 GstarCAD S17101S-D0301
        # 平断面图及杆塔明细表.dwg ✓」独占 363px，栏内仅有 692px，而三个按钮需
        # 302px + 目标选择器 104px），后 pack 的「连接/刷新 CAD窗口」只能分到 43px
        # 而被切掉一半。按钮优先分配 → 文字让位并自行省略（见 _render_status）。
        ttk.Button(f0, text="使用说明", command=self.open_help).pack(side="right", padx=(0, 4))
        ttk.Button(f0, text="诊断 CAD", command=self.diagnose).pack(side="right", padx=(0, 4))
        ttk.Button(f0, text="连接/刷新 CAD窗口", command=self.connect).pack(side="right")
        # 目标 CAD 选择器：自动检测 + 配置表里全部 CAD（见 catenary_cad._CAD_FAMILIES）
        # （自动模式按表序兜底；多开时用此项强制指定，避免连错 CAD）
        self.cad_prefer_var = tk.StringVar(value="自动")
        sel = ttk.Frame(f0)
        sel.pack(side="right", padx=(8, 8))
        ttk.Label(sel, text="目标:").pack(side="left")
        ttk.OptionMenu(
            sel, self.cad_prefer_var, "自动",
            *(["自动"] + [lbl for _, lbl in cad_family_labels()]),
        ).pack(side="left")
        # 状态文字最后 pack：吃掉剩余宽度，过长时中间省略（不硬裁、不挤按钮）
        self.status_var = tk.StringVar(value="未连接")
        self.status_show_var = tk.StringVar(value="未连接")
        self.status_var.trace_add("write", self._on_status_change)
        self.status_lbl = ttk.Label(f0, textvariable=self.status_show_var, anchor="w")
        self.status_lbl.pack(side="left", fill="x", expand=True)
        self.status_lbl.bind("<Configure>", self._on_status_resize)

        # 模块 1：生成弧垂曲线（支持 K 值 / 应力σ₀+比载γ 两种输入）
        f1 = ttk.LabelFrame(self.root, text="① 生成空间弧垂曲线")
        f1.pack(fill="x", **pad)

        # 输入方式切换
        mrow = ttk.Frame(f1)
        mrow.pack(fill="x")
        ttk.Label(mrow, text="输入方式:").pack(side="left")
        self.gen_mode_var = tk.StringVar(value="K值")
        ttk.OptionMenu(
            mrow, self.gen_mode_var, "K值", "K值", "应力+比载",
            command=self._on_gen_mode_change,
        ).pack(side="left", padx=4)

        # —— K 值输入子框（K 值 + 采样点数 + 按钮同一行）——
        self.gen_k_frame = ttk.Frame(f1)
        ttk.Label(self.gen_k_frame, text="K 值 (省略10⁻⁵):").pack(side="left")
        self.k_var = tk.StringVar(value="15.22")
        ttk.Entry(self.gen_k_frame, textvariable=self.k_var, width=10).pack(side="left", padx=4)
        ttk.Label(self.gen_k_frame, text="采样点数:").pack(side="left", padx=(12, 0))
        self.n_var = tk.StringVar(value="200")
        ttk.Entry(self.gen_k_frame, textvariable=self.n_var, width=6).pack(side="left", padx=4)
        ttk.Button(self.gen_k_frame, text="拾取两点并生成", command=self.do_generate).pack(side="right")

        # —— 应力σ₀ + 比载γ 输入子框（默认隐藏；σγ 输入 + 采样点数 + 按钮同一行）——
        self.gen_sg_frame = ttk.Frame(f1)
        ttk.Label(self.gen_sg_frame, text="应力σ₀(N/mm²):").pack(side="left")
        self.sigma_var = tk.StringVar(value="50")
        ttk.Entry(self.gen_sg_frame, textvariable=self.sigma_var, width=8).pack(side="left", padx=4)
        ttk.Label(self.gen_sg_frame, text="比载γ(N/(m·mm²)):").pack(side="left")
        self.gamma_var = tk.StringVar(value="0.06088")
        ttk.Entry(self.gen_sg_frame, textvariable=self.gamma_var, width=10).pack(side="left", padx=4)
        ttk.Label(self.gen_sg_frame, text="采样点数:").pack(side="left", padx=(12, 0))
        ttk.Entry(self.gen_sg_frame, textvariable=self.n_var, width=6).pack(side="left", padx=4)
        ttk.Button(self.gen_sg_frame, text="拾取两点并生成", command=self.do_generate).pack(side="right")

        # 初始按默认方式显示对应输入子框（紧跟输入方式行，中间不隔状态文字）
        self._on_gen_mode_change("K值")

        # 状态文字（放在输入行下方）
        self.gen_info = tk.StringVar(value="点击按钮后，在 CAD 中依次点选两个悬挂点。两种输入方式等价：K值 或 应力σ₀+比载γ（K=γ/(8σ₀)），均用悬链线方程 + 牛顿迭代求解。")
        ttk.Label(f1, textvariable=self.gen_info, foreground="#444", wraplength=620, justify="left").pack(anchor="w")

        # 模块 2：最小距离
        f2 = ttk.LabelFrame(self.root, text="② 两条曲线最小距离")
        f2.pack(fill="x", **pad)
        ttk.Button(f2, text="依次拾取两条曲线", command=self.do_distance).pack(anchor="w")
        self.dist_info = tk.StringVar(value="点击按钮后，在 CAD 中依次点选两条曲线（可为直线/圆弧/圆/多段线/样条）。")
        ttk.Label(f2, textvariable=self.dist_info, foreground="#444", wraplength=620, justify="left").pack(anchor="w")

        # 模块 3：风偏弧垂曲线
        fw = ttk.LabelFrame(self.root, text="③ 风偏弧垂曲线")
        fw.pack(fill="x", **pad)
        wrow = ttk.Frame(fw)
        wrow.pack(fill="x")
        ttk.Label(wrow, text="风偏角(度, 右偏为正):").pack(side="left")
        self.wind_angle_var = tk.StringVar(value="30.6")
        ttk.Entry(wrow, textvariable=self.wind_angle_var, width=10).pack(side="left", padx=4)
        ttk.Button(wrow, text="拾取曲线并生成", command=self.do_wind).pack(side="right")
        self.wind_info = tk.StringVar(value="点击按钮后，在 CAD 中点选一条弧垂曲线，将绕其两端点连线旋转生成风偏曲线（蓝色）。角度正负约定：右偏为正(+)、左偏为负(−)，以曲线起点→终点为前进方向观察。")
        ttk.Label(fw, textvariable=self.wind_info, foreground="#444", wraplength=620, justify="left").pack(anchor="w")

        # 模块 4：分裂导线
        fb = ttk.LabelFrame(self.root, text="④ 分裂导线")
        fb.pack(fill="x", **pad)
        brow = ttk.Frame(fb)
        brow.pack(fill="x")
        ttk.Label(brow, text="分裂间距(mm):").pack(side="left")
        self.bundle_d_var = tk.StringVar(value="600")
        ttk.Entry(brow, textvariable=self.bundle_d_var, width=8).pack(side="left", padx=4)
        ttk.Label(brow, text="模式:").pack(side="left")
        self.bundle_mode_var = tk.StringVar(value="四分裂")
        ttk.OptionMenu(
            brow, self.bundle_mode_var, "四分裂", "水平双分裂", "垂直双分裂", "四分裂"
        ).pack(side="left", padx=4)
        ttk.Button(brow, text="拾取中心线并生成", command=self.do_bundle).pack(side="right")
        self.bundle_info = tk.StringVar(value="拾取弧垂曲线为中心线，按分裂间距生成子导线（黄色），单位 mm（输入600=图上0.6m）。")
        ttk.Label(fb, textvariable=self.bundle_info, foreground="#444", wraplength=620, justify="left").pack(anchor="w")

        # 模块 5：含悬垂串弧垂线
        fs = ttk.LabelFrame(self.root, text="⑤ 含悬垂串弧垂线")
        fs.pack(fill="x", **pad)
        srow1 = ttk.Frame(fs)
        srow1.pack(fill="x")
        self.susp1_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(srow1, text="点1悬垂串", variable=self.susp1_var).pack(side="left")
        ttk.Label(srow1, text="串长D1(m):").pack(side="left")
        self.susp1_d_var = tk.StringVar(value="3.0")
        ttk.Entry(srow1, textvariable=self.susp1_d_var, width=6).pack(side="left", padx=2)
        srow1b = ttk.Frame(fs)
        srow1b.pack(fill="x")
        self.susp2_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(srow1b, text="点2悬垂串", variable=self.susp2_var).pack(side="left")
        ttk.Label(srow1b, text="串长D2(m):").pack(side="left")
        self.susp2_d_var = tk.StringVar(value="3.0")
        ttk.Entry(srow1b, textvariable=self.susp2_d_var, width=6).pack(side="left", padx=2)
        # 风偏选项
        srow2 = ttk.Frame(fs)
        srow2.pack(fill="x")
        self.susp_wind_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(srow2, text="绘制风偏线", variable=self.susp_wind_var).pack(side="left")
        ttk.Label(srow2, text="线偏角B(°,右偏正):").pack(side="left")
        self.susp_b_var = tk.StringVar(value="0")
        ttk.Entry(srow2, textvariable=self.susp_b_var, width=6).pack(side="left", padx=2)
        srow2b = ttk.Frame(fs)
        srow2b.pack(fill="x")
        ttk.Label(srow2b, text="串角A1(°,右偏正):").pack(side="left")
        self.susp1_a_var = tk.StringVar(value="0")
        ttk.Entry(srow2b, textvariable=self.susp1_a_var, width=6).pack(side="left", padx=2)
        ttk.Label(srow2b, text="串角A2(°,右偏正):").pack(side="left")
        self.susp2_a_var = tk.StringVar(value="0")
        ttk.Entry(srow2b, textvariable=self.susp2_a_var, width=6).pack(side="left", padx=2)
        # 输入方式（K值 / 应力+比载）放在风偏选项下方
        srowM = ttk.Frame(fs)
        srowM.pack(fill="x")
        ttk.Label(srowM, text="输入方式:").pack(side="left")
        self.sus_mode_var = tk.StringVar(value="K值")
        ttk.OptionMenu(
            srowM, self.sus_mode_var, "K值", "K值", "应力+比载",
            command=self._on_sus_mode_change,
        ).pack(side="left", padx=4)
        # —— K 值输入子框（K 值 + 采样点数 + 按钮同一行）——
        self.sus_k_frame = ttk.Frame(fs)
        ttk.Label(self.sus_k_frame, text="K 值 (省略10⁻⁵):").pack(side="left")
        self.sus_k_var = tk.StringVar(value="15.22")
        ttk.Entry(self.sus_k_frame, textvariable=self.sus_k_var, width=8).pack(side="left", padx=2)
        ttk.Label(self.sus_k_frame, text="采样点数:").pack(side="left", padx=(12, 0))
        ttk.Entry(self.sus_k_frame, textvariable=self.n_var, width=6).pack(side="left", padx=2)
        ttk.Button(self.sus_k_frame, text="拾取两点并生成", command=self.do_suspension).pack(side="right")
        # —— 应力σ₀ + 比载γ 输入子框（默认隐藏；σγ 输入 + 采样点数 + 按钮同一行）——
        self.sus_sg_frame = ttk.Frame(fs)
        ttk.Label(self.sus_sg_frame, text="应力σ₀(N/mm²):").pack(side="left")
        self.sus_sigma_var = tk.StringVar(value="50")
        ttk.Entry(self.sus_sg_frame, textvariable=self.sus_sigma_var, width=7).pack(side="left", padx=2)
        ttk.Label(self.sus_sg_frame, text="比载γ(N/(m·mm²)):").pack(side="left")
        self.sus_gamma_var = tk.StringVar(value="0.06088")
        ttk.Entry(self.sus_sg_frame, textvariable=self.sus_gamma_var, width=9).pack(side="left", padx=2)
        ttk.Label(self.sus_sg_frame, text="采样点数:").pack(side="left", padx=(12, 0))
        ttk.Entry(self.sus_sg_frame, textvariable=self.n_var, width=6).pack(side="left", padx=2)
        ttk.Button(self.sus_sg_frame, text="拾取两点并生成", command=self.do_suspension).pack(side="right")
        # 初始按默认方式显示对应输入子框
        self._on_sus_mode_change("K值")
        self.susp_info = tk.StringVar(value="点选两个悬挂点；勾选悬垂串输入串长D(m)。勾选绘制风偏线才需输入：线偏角B(°)、串角A1/A2(°)，未勾选时串角按0（竖直串）。角度正负约定：右偏为正(+)、左偏为负(−)，\"左右\"以点选先后顺序为前进方向——面向第1点→第2点方向观察，右手边即为\"右\"。")
        ttk.Label(fs, textvariable=self.susp_info, foreground="#444",
                  wraplength=620, justify="left").pack(anchor="w")

        # 模块 6：二维断面悬链线
        f2d = ttk.LabelFrame(self.root, text="⑥ 二维断面悬链线（断面图）")
        f2d.pack(fill="x", **pad)
        s2row = ttk.Frame(f2d)
        s2row.pack(fill="x")
        ttk.Label(s2row, text="K 值 (省略10⁻⁵):").pack(side="left")
        self.sec_k_var = tk.StringVar(value="15.22")
        ttk.Entry(s2row, textvariable=self.sec_k_var, width=8).pack(side="left", padx=2)
        ttk.Label(s2row, text="图纸比例").pack(side="left", padx=(12, 0))
        ttk.Label(s2row, text="横向").pack(side="left", padx=(6, 0))
        self.sec_rx_var = tk.StringVar(value="5000")
        ttk.Entry(s2row, textvariable=self.sec_rx_var, width=7).pack(side="left", padx=2)
        ttk.Label(s2row, text="纵向").pack(side="left", padx=(6, 0))
        self.sec_ry_var = tk.StringVar(value="500")
        ttk.Entry(s2row, textvariable=self.sec_ry_var, width=6).pack(side="left", padx=2)
        ttk.Label(s2row, text="采样点数:").pack(side="left", padx=(6, 0))
        ttk.Entry(s2row, textvariable=self.n_var, width=5).pack(side="left", padx=2)
        ttk.Button(s2row, text="拾取挂线点并生成", command=self.do_section).pack(side="right")
        # 说明栏：显式按句分行（每行一个完整句子），wraplength 取足够大，
        # 避免 Tk 只按空格断词而在句子中间折行（会出现一行孤字）。
        self.sec_info = tk.StringVar(
            value="可连续点选：点第 1、2 点生成第 1 段，之后每点一次接一段"
                  "（上段终点＝本段起点），ESC/右键结束。\n"
                  "X＝档距方向、Y＝高程方向。\n"
                  "比例：图纸 1 单位＝1mm → 横向 1:5000 时图上 1mm＝实际 5m（×1/5），"
                  "纵向 1:500 时图上 1mm＝实际 0.5m（×2）。")
        ttk.Label(f2d, textvariable=self.sec_info, foreground="#444",
                  wraplength=690, justify="left").pack(anchor="w")

        # 署名（先 pack 固定底部空间，日志框再占剩余空间）
        tk.Label(self.root, text="作者：Mz  ·  github.com/kophmz/catenary-cad-tool",
                 fg="#999", font=("Microsoft YaHei", 8)).pack(side="bottom", pady=(0, 4))

        # 日志
        f3 = ttk.LabelFrame(self.root, text="日志")
        f3.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(f3, height=15, wrap="word")
        self.log.pack(fill="both", expand=True, padx=4, pady=4)

    def log_msg(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def open_help(self):
        """用系统默认浏览器打开使用说明。"""
        path = _help_path()
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showinfo("说明文档", "未找到使用说明.html，请确认文件与程序在同一目录下。")

    # ---- 状态栏与窗口宽度自适应 --------------------------------------
    def _screen_work_width(self):
        """窗口可用的最大宽度：屏幕宽 - 左右边距（防止撑出屏幕）。"""
        try:
            return max(self._win_min_w, self.root.winfo_screenwidth() - 40)
        except Exception:
            return self._win_min_w

    def _conn_bar_extra(self):
        """「连接」栏里除状态文字以外的固定占用（控件需求宽 + 各种 padding）。

        按组件实测而非 estimate，避免改按钮/下拉文字后这里失配：
        按钮 padx=(0,4)×3 + 目标框 padx=(8,8) + LabelFrame 边框 + 外层 padx=12×2。
        ⚠️ 不能用 root.winfo_reqwidth() —— 它对子控件需求变化的更新是滞后的
        （实测设置长文本后它仍返回旧值，导致窗口宽度算不出来）。
        """
        total = 0
        for w in self._conn_frame.winfo_children():
            if w is not self.status_lbl:
                total += w.winfo_reqwidth()
        return total + 12 + 16 + 4 + 24

    def _on_status_change(self, *_a):
        """status_var 一变就排队重绘（合并同帧内多次 set，避免重入）。"""
        if self._status_pending:
            return
        self._status_pending = True
        try:
            self.root.after_idle(self._render_status)
        except Exception:
            self._status_pending = False

    def _render_status(self):
        """状态文字 + 窗口宽度一起自适应。

        先把**全文**放上去让 Tk 算出真实宽度，再把窗口撑到刚好放得下
        （下限 _win_min_w、上限 min(屏幕可用宽, _win_max_w)）。这样长图纸名
        （「已连接 GstarCAD S17101S-D0301 平断面图及杆塔明细表.dwg ✓」）既不
        挤掉右侧「连接/刷新」按钮，也不用牺牲信息；只有撑到上限仍放不下时，
        才退化为中间省略 + 保留尾部（见 _elide_to_fit）。

        ⚠️ 必须防重入：本函数内部的 `update_idletasks()` 会**触发已排队的
        after_idle 回调**（即又一次 _render_status）。没有护栏时第二次调用会
        量到「已被省略的短文本」，把刚算出的宽窗口又缩回 720。
        """
        if self._rendering:
            return
        self._rendering = True
        self._status_pending = False
        try:
            full = self.status_var.get()
            self.status_show_var.set(full)
            self.root.update_idletasks()
            full_w = self.status_lbl.winfo_reqwidth()   # 先量全文，再谈省略
            extra = self._conn_bar_extra()
            cap = min(self._screen_work_width(), self._win_max_w)
            w = max(self._win_min_w, min(full_w + extra, cap))
            self._set_window_width(w)
            # 用**算出来的**可用宽判断是否省略：实测设置完 geometry 后
            # update_idletasks() 拿到的 winfo_width() 仍是旧值（WM 尚未回灌），
            # 会造成「窗口已经变宽、文字却还是省略号」。
            self._apply_text_fit(max(0, w - extra))
        except Exception:
            pass
        finally:
            self._rendering = False

    def _apply_text_fit(self, avail):
        """在 avail 像素内显示状态文字：放得下就全文，放不下才省略。"""
        if avail <= 1:
            return
        prev, self._rendering = self._rendering, True
        try:
            full = self.status_var.get()
            self.status_show_var.set(full)
            self.root.update_idletasks()
            if self.status_lbl.winfo_reqwidth() > avail:
                self.status_show_var.set(self._elide_to_fit(full, avail))
        except Exception:
            pass
        finally:
            self._rendering = prev

    def _elide_to_fit(self, text, avail):
        """把文字压进 avail 像素，优先保留**尾部**（`.dwg ✓` 这类辨识度最高的信息）。

        度量一律走 Label 自己的 winfo_reqwidth()（Tk 排版用的同一套字体度量），
        不用 tkfont.measure —— 实测后者比 Tk 实际排版宽约 25%，会把文字压得过短。
        """
        ell = "…"
        prev, self._rendering = self._rendering, True   # 防重入（见 _render_status）

        def fits(s):
            self.status_show_var.set(s)
            self.root.update_idletasks()
            return self.status_lbl.winfo_reqwidth() <= avail

        try:
            if fits(text):
                return text
            if not fits(ell):
                return ell
            # 先保尾部，再二分求「头部最多留几个字」
            for tail_chars in (12, 8, 5, 0):
                if len(text) <= tail_chars:
                    continue
                tail = text[-tail_chars:] if tail_chars else ""
                lo, hi, ok = 1, len(text) - tail_chars, 0
                while lo <= hi:
                    mid = (lo + hi) // 2
                    if fits(text[:mid] + ell + tail):
                        ok, lo = mid, mid + 1
                    else:
                        hi = mid - 1
                if ok:
                    return text[:ok].rstrip() + ell + tail
            return ell
        finally:
            self._rendering = prev

    def _set_window_width(self, w):
        """改宽度、保持高度与纵向位置；若会超出屏幕右缘则整体左移。

        用 min=max=w 固定尺寸（等价于「不可缩放」），同时**必须同步更新 max
        约束**，否则加宽请求会被 Windows 夹回旧宽度（见 __init__ 注释）。
        """
        try:
            h = getattr(self, "_win_h", 1000)
            x, y = self.root.winfo_x(), self.root.winfo_y()
            sw = self.root.winfo_screenwidth()
            if x < 0:
                x = 0
            if x + w > sw:
                x = max(0, sw - w - 8)
            if y < 0:
                y = 0
            # 三步走：先放开上界（避免出现 maxsize < 当前 minsize 的矛盾区间，
            # 收窄时会走到这一步），再设 min=max=w 把尺寸钉死，最后设位置尺寸。
            self.root.maxsize(max(w, self.root.winfo_screenwidth()), h)
            self.root.minsize(w, h)
            self.root.maxsize(w, h)
            self.root.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _on_status_resize(self, _e=None):
        """状态栏宽度变化后重新适配文字（只写回真变了的文本，避免循环）。"""
        if self._status_pending or self._rendering:
            return
        avail = self.status_lbl.winfo_width()
        if avail <= 1:
            return
        self._apply_text_fit(avail)

    # ---- 连接 -------------------------------------------------------
    def connect(self, silent=False):
        """连接 CAD（下拉目标 + 前台实例判定）。

        silent=True：失败只写日志、不弹错误框 —— 供「前台切换后自动跟随
        重连」调用，避免用户每点一次功能按钮就被模态框打断。
        """
        # 将下拉选择映射到 get_gcad 的 prefer 参数（映射表由配置表生成）
        _prefer_map = {"自动": None}
        _prefer_map.update({lbl: key for key, lbl in cad_family_labels()})
        prefer = _prefer_map.get(self.cad_prefer_var.get(), None)

        # 前台实例判定，拿到 (key, hwnd, progid) 三元组。
        # 多 CAD / 同 CAD 多版本共存时，只有「产出该窗口的 ProgID」才能连到
        # 用户眼前的实例（如 AutoCAD 2026 要用 AutoCAD.Application.25，
        # 而裸别名 AutoCAD.Application 指向的是 2020）。
        fg = None
        try:
            fg = foreground_cad_target(self.root.winfo_id())
        except Exception:
            fg = None

        if prefer is None:
            # 自动：若某 CAD 确为前台则用它；否则 prefer 保持 None，
            # 交给 get_gcad 按配置表行序兜底
            prefer = fg[0] if fg else None
            if prefer:
                self.log_msg(f"→ 自动检测到：{cad_label_of(prefer)}（{fg[2]}）")
            else:
                self.log_msg(
                    f"→ 自动检测：未发现活跃CAD窗口，按默认顺序尝试（{_all_cad_names()}）")
        # 精确 ProgID：仅当前台实例与目标 CAD 同类时才锁定；
        # 用户手动指定了另一种 CAD 时不做精确锁定，尊重其选择
        exact = fg[2] if (fg and fg[0] == prefer) else None

        try:
            self.gc = self.doc = self.ms = None  # 先清空旧连接，确保真正刷新
            self._conn_hwnd = None
            self.gc, self.doc, self.ms, app = get_gcad(prefer, progid=exact)
            self.app_name = app  # 供后续 pick_entity 按 CAD 类型选择调用方式
            # 记录已连实例的窗口句柄（前台切换后自动跟随重连时比对用）
            self._conn_hwnd = int(getattr(self.gc, "HWND", 0)) or None
            # doc.Name 形如 "Drawing1"（未保存）或完整路径；直接展示连接的窗口
            self.status_var.set(f"已连接 {app} {self.doc.Name} ✓")
            self.log_msg(f"✔ 已连接 {app} {self.doc.Name}（目标={self.cad_prefer_var.get()}）")
            return True
        except Exception as e:
            self.status_var.set("连接失败")
            err_msg = str(e)
            self.log_msg(f"✘ 连接失败: {e}")
            if silent:
                return False
            # 检测管理员权限不匹配 → 询问是否提权重启
            # （get_gcad 已把 HRESULT 转成中文分类 + 场景化诊断）
            if (('管理员身份运行' in err_msg and 'COM 连接被隔离' in err_msg)
                    or ('RUNASADMIN' in err_msg
                        and '非管理员进程无法通过 COM 启动' in err_msg)):
                from catenary_cad import _is_running_as_admin, _try_auto_elevate
                if not _is_running_as_admin():
                    if messagebox.askyesno(
                        "需要管理员权限",
                        "检测到 CAD 正以管理员身份运行，导致 COM 连接被隔离。\n\n"
                        "是否以管理员身份重启本程序？\n"
                        '（点击"是"后会弹出 UAC 确认框）',
                    ):
                        if _try_auto_elevate():
                            self.root.destroy()
                            return
                messagebox.showerror(
                    "连接失败",
                    f"无法连接 CAD（{self.cad_prefer_var.get()}）：\n{e}\n\n"
                    "提示：请手动以管理员身份运行本程序，"
                    '或关闭 CAD 后取消其"以管理员身份运行"设置。',
                )
            else:
                # 其他错误：get_gcad 已含分类与诊断，直接展示
                messagebox.showerror(
                    "连接失败",
                    f"无法连接 CAD（{self.cad_prefer_var.get()}）：\n\n{e}",
                )
            return False

    def diagnose(self):
        """弹出本机 CAD COM 注册诊断结果，帮助用户自助排查。"""
        report = diagnose_cad_connection()
        messagebox.showinfo("CAD 连接诊断", report)
        self.log_msg("ℹ CAD 诊断信息已显示")

    # ---- 模块1 ------------------------------------------------------
    def _on_gen_mode_change(self, mode):
        """输入方式切换：K值 / 应力+比载，显示对应输入子框。"""
        if mode == "K值":
            self.gen_sg_frame.pack_forget()
            self.gen_k_frame.pack(fill="x")
        else:  # 应力+比载 需要 σ₀、γ 输入
            self.gen_k_frame.pack_forget()
            self.gen_sg_frame.pack(fill="x")

    def do_generate(self):
        if not self._ensure_conn():
            return
        mode = self.gen_mode_var.get()
        try:
            n = max(8, int(self.n_var.get()))
            if mode == "K值":
                K = float(self.k_var.get()) * 1e-5
                param_desc = f"K={float(self.k_var.get())}×10⁻⁵"
            else:
                # 应力σ₀(N/mm²) + 比载γ(N/(m·mm²)) → K = γ/(8σ₀)，单位 1/m
                sigma0 = float(self.sigma_var.get())
                gamma = float(self.gamma_var.get())
                if sigma0 <= 0:
                    raise ValueError("应力σ₀必须为正数")
                K = gamma / (8.0 * sigma0)
                param_desc = (f"σ₀={sigma0}N/mm², γ={gamma}N/(m·mm²) "
                              f"→ K={K * 1e5:.2f}×10⁻⁵")
        except ValueError as e:
            # 区分「数字格式错误」与「数值范围错误」（如 σ₀≤0），避免误导性文案
            if "必须为正数" in str(e):
                messagebox.showerror("输入错误", str(e))
            else:
                messagebox.showerror("输入错误", f"输入必须是数字：{e}")
            return

        self.log_msg("→ 请在 CAD 中点选两个悬挂点（窗口已最小化，选完恢复）…")
        self.root.update()
        try:
            self.root.iconify()  # 最小化 GUI，避免遮挡 CAD 选择
            p1 = pick_point(self.doc, "\n选择第1个悬挂点: ")
            p2 = pick_point(self.doc, "\n选择第2个悬挂点: ")
        except Exception as e:
            self.root.deiconify(); self.root.lift()
            self._pick_abort(e, "点选")
            return
        finally:
            self.root.deiconify(); self.root.lift()

        try:
            # K值 / 应力+比载 → 统一 cosh + 牛顿迭代求解（含高差修正）
            pts, sag, L, a, x0 = catenary_3d(p1, p2, K, n)
            draw_catenary(self.doc, self.ms, pts)
            self.gc.ZoomExtents()
            info = (f"{param_desc}\n"
                    f"档距={L:.2f}m  高差={p2[2]-p1[2]:.2f}m  弧垂={sag:.3f}m")
            self.gen_info.set(info)
            self.log_msg("✔ 已生成三维弧垂曲线: " + info.replace("\n", "  "))
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            self.log_msg(f"✘ 生成失败: {e}")

    # ---- 模块2 ------------------------------------------------------
    def do_distance(self):
        if not self._ensure_conn():
            return
        self.log_msg("→ 请在 CAD 中点选两条曲线（窗口已最小化，选完恢复）…")
        self.root.update()
        try:
            self.root.iconify()  # 最小化 GUI，避免遮挡 CAD 选择
            e1 = pick_entity(self.doc, "\n选择第1条曲线: ", self.app_name)
            e2 = pick_entity(self.doc, "\n选择第2条曲线: ", self.app_name)
        except Exception as e:
            self.root.deiconify(); self.root.lift()
            self._pick_abort(e, "选取")
            return
        finally:
            self.root.deiconify(); self.root.lift()

        try:
            pts1 = extract_points(e1)
            pts2 = extract_points(e2)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            self.log_msg(f"✘ {e}")
            return

        d, c1, c2 = closest_distance(pts1, pts2)
        try:
            draw_distance_line(self.doc, self.ms, c1, c2)
            draw_distance_label(self.doc, self.ms, c1, c2, f"{d:.2f}m")
            self.gc.ZoomExtents()
        except Exception as e:
            self.log_msg(f"✘ 绘制连线失败: {e}")

        info = f"最小距离 = {d:.3f} m"
        self.dist_info.set(info)
        self.log_msg("✔ " + info.replace("\n", "  "))

    # ---- 模块3 ------------------------------------------------------
    def do_wind(self):
        if not self._ensure_conn():
            return
        try:
            angle = float(self.wind_angle_var.get())
        except ValueError:
            messagebox.showerror("输入错误", "风偏角必须是数字（可带负号）。")
            return

        self.log_msg("→ 请在 CAD 中点选一条弧垂曲线（窗口已最小化，选完恢复）…")
        self.root.update()
        try:
            self.root.iconify()  # 最小化 GUI，避免遮挡 CAD 选择
            ent = pick_entity(self.doc, "\n选择一条弧垂曲线: ", self.app_name)
        except Exception as e:
            self.root.deiconify(); self.root.lift()
            self._pick_abort(e, "选取")
            return
        finally:
            self.root.deiconify(); self.root.lift()

        try:
            pts = extract_points(ent)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            self.log_msg(f"✘ {e}")
            return

        try:
            new_pts = rotate_curve(pts, angle)
            draw_wind_curve(self.doc, self.ms, new_pts)
            draw_wind_label(self.doc, self.ms, new_pts, angle)
            self.gc.ZoomExtents()
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            self.log_msg(f"✘ 生成失败: {e}")
            return

        info = f"风偏角={angle}°（标签按绝对值标示）"
        self.wind_info.set(info)
        self.log_msg("✔ " + info.replace("\n", "  "))

    # ---- 模块4 ------------------------------------------------------
    def do_bundle(self):
        if not self._ensure_conn():
            return
        try:
            d_mm = float(self.bundle_d_var.get())
            if d_mm <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("输入错误", "分裂间距必须是正数（单位 mm）。")
            return
        mode = self.bundle_mode_var.get()
        mode_map = {"水平双分裂": BUNDLE_HORIZ, "垂直双分裂": BUNDLE_VERT, "四分裂": BUNDLE_QUAD}
        mode_code = mode_map[mode]
        d_m = d_mm / 1000.0  # 用户单位 mm → 绘图单位 m

        self.log_msg("→ 请在 CAD 中点选中心线（窗口已最小化，选完恢复）…")
        self.root.update()
        try:
            self.root.iconify()  # 最小化 GUI，避免遮挡 CAD 选择
            ent = pick_entity(self.doc, "\n选择中心线(弧垂曲线): ", self.app_name)
        except Exception as e:
            self.root.deiconify(); self.root.lift()
            self._pick_abort(e, "选取")
            return
        finally:
            self.root.deiconify(); self.root.lift()

        try:
            pts = extract_points(ent)
            lines = generate_bundle_lines(pts, d_m, mode_code)
            for line in lines:
                draw_bundle_line(self.doc, self.ms, line)
            self.gc.ZoomExtents()
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            self.log_msg(f"✘ 生成失败: {e}")
            return

        info = f"模式={mode}  间距={d_mm:g}mm (= {d_m:g} m)"
        self.bundle_info.set(info)
        self.log_msg("✔ " + info.replace("\n", "  "))

    # ---- 模块5 ------------------------------------------------------
    def _on_sus_mode_change(self, mode):
        """悬垂串模块输入方式切换：K值 / 应力+比载，显示对应输入子框。"""
        if mode == "K值":
            self.sus_sg_frame.pack_forget()
            self.sus_k_frame.pack(fill="x")
        else:
            self.sus_k_frame.pack_forget()
            self.sus_sg_frame.pack(fill="x")

    def do_suspension(self):
        if not self._ensure_conn():
            return
        mode = self.sus_mode_var.get()
        try:
            n = max(8, int(self.n_var.get()))
            if mode == "K值":
                K = float(self.sus_k_var.get()) * 1e-5
                param_desc = f"K={float(self.sus_k_var.get())}×10⁻⁵"
            else:
                # 应力σ₀(N/mm²) + 比载γ(N/(m·mm²)) → K = γ/(8σ₀)，单位 1/m
                sigma0 = float(self.sus_sigma_var.get())
                gamma = float(self.sus_gamma_var.get())
                if sigma0 <= 0:
                    raise ValueError("应力σ₀必须为正数")
                K = gamma / (8.0 * sigma0)
                param_desc = (f"σ₀={sigma0}N/mm², γ={gamma}N/(m·mm²) "
                              f"→ K={K * 1e5:.2f}×10⁻⁵")
            D1 = float(self.susp1_d_var.get()) if self.susp1_var.get() else None
            D2 = float(self.susp2_d_var.get()) if self.susp2_var.get() else None
            wind_on = self.susp_wind_var.get()
            # 串风偏角仅在勾选「绘制风偏线」时生效；未勾选一律按 0（竖直串）
            A1 = float(self.susp1_a_var.get()) if (wind_on and D1 is not None) else 0.0
            A2 = float(self.susp2_a_var.get()) if (wind_on and D2 is not None) else 0.0
            B = float(self.susp_b_var.get()) if wind_on else 0.0
        except ValueError as e:
            if "必须为正数" in str(e):
                messagebox.showerror("输入错误", str(e))
            else:
                messagebox.showerror("输入错误", f"输入必须是数字：{e}")
            return

        self.log_msg("→ 请在 CAD 中点选两个悬挂点（窗口已最小化，选完恢复）…")
        self.root.update()
        try:
            self.root.iconify()  # 最小化 GUI，避免遮挡 CAD 选择
            p1 = pick_point(self.doc, "\n选择第1个悬挂点: ")
            p2 = pick_point(self.doc, "\n选择第2个悬挂点: ")
        except Exception as e:
            self.root.deiconify(); self.root.lift()
            self._pick_abort(e, "点选")
            return
        finally:
            self.root.deiconify(); self.root.lift()

        try:
            geo = suspension_geometry(p1, p2, K, D1, A1, D2, A2, B, n)
            wind_mode = geo["wind_pts"] is not None
            # 悬垂串：风偏模式 → 摆串浅蓝（新图层）；非风偏 → 竖直串品红
            if wind_mode:
                for top, bottom in geo["susp_lines"]:
                    draw_suspension_line(
                        self.doc, self.ms, top, bottom,
                        color=COLOR_SUSPENSION_WIND, layer=LAYER_SUSPENSION_WIND)
            else:
                for top, bottom in geo["susp_lines"]:
                    draw_suspension_line(self.doc, self.ms, top, bottom)
            # 风偏模式：只画风偏线（蓝），不再画未风偏的参照曲线（红基准）与
            # 未风偏的参照绝缘子（A=0 品红竖直串）；非风偏才画红基准弧垂
            if geo["wind_pts"] is not None:
                draw_wind_curve(self.doc, self.ms, geo["wind_pts"])
            else:
                draw_catenary(self.doc, self.ms, geo["base_pts"])
            self.gc.ZoomExtents()
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            self.log_msg(f"✘ 生成失败: {e}")
            return

        parts = [
            f"{param_desc}",
            f"档距={geo['L']:.2f}m  弧垂={geo['sag']:.3f}m",
        ]
        if wind_mode:
            parts.append("已生成：风偏线（蓝）+ 摆串（浅蓝）")
            parts.append(f"风偏角 B={B:g}°")
        else:
            parts.append("已生成：悬垂串（品红）+ 基准弧垂（红）")
        info = "\n".join(parts)
        self.susp_info.set(info)
        self.log_msg("✔ " + info.replace("\n", "  "))

    # ---- 模块6 ------------------------------------------------------
    def do_section(self):
        """二维断面悬链线（连续点选）：图上两点 ÷ 比例 → 实际米数 → 算 K 值悬链线 → × 比例画回图上。

        连续点选：点第 2 点生成「1-2」段后不退出，继续点第 3 点即生成「2-3」段
        （上一段终点自动作为本段起点），如此可连续生成多段，K 值与比例全程相同；
        按 ESC / 右键结束。端点不合法（如 X 相同）的段自动跳过并提示，不影响后续点选。
        """
        if not self._ensure_conn():
            return
        try:
            n = max(8, int(self.n_var.get()))
            K = float(self.sec_k_var.get()) * 1e-5
            rx = float(self.sec_rx_var.get())
            ry = float(self.sec_ry_var.get())
            if rx <= 0 or ry <= 0:
                raise ValueError("图纸比例分母必须为正数")
            param_desc = f"K={float(self.sec_k_var.get())}×10⁻⁵  横向1:{rx:g}  纵向1:{ry:g}"
        except ValueError as e:
            if "必须为正数" in str(e):
                messagebox.showerror("输入错误", str(e))
            else:
                messagebox.showerror("输入错误", f"输入必须是数字：{e}")
            return

        self.log_msg("→ 请在 CAD 中连续点选挂点（窗口已最小化）… "
                     "每点一次生成一段，ESC/右键结束")
        self.root.update()

        segs = []          # [(info, handle, mode, p_from, p_to), ...]
        err = None         # 致命错误（非用户取消）
        warn = None        # 段级警告（该段被跳过）
        p_prev = None
        first = True
        try:
            self.root.iconify()  # 最小化 GUI，避免遮挡 CAD 点选；整条链结束才恢复
            while True:
                if first:
                    prompt = "\n选择第1个挂点: "
                elif not segs:
                    prompt = "\n选择第2个挂点: "
                else:
                    prompt = "\n选择下一个挂点 (ESC/右键结束): "
                try:
                    p_next = pick_point(self.doc, prompt)
                except Exception as e:
                    if self._is_user_cancel(e):
                        break          # ESC / 右键 = 正常结束连续点选
                    # 兜底：链式模式下取点失败一律结束链路（已画的段全部保留），
                    # 不弹错误框、不丢结果 —— 避免各 CAD 取消文案不同导致误报
                    warn = f"⚠ 连续点选已结束（取点失败：{e}），已生成的曲线保留"
                    break
                if first:
                    p_prev, first = p_next, False
                    continue
                try:
                    # 只用 X/Y（丢弃拾取点的 Z）
                    pts, info = catenary_section_2d(p_prev, p_next, K, rx, ry, n)
                except ValueError as e:
                    warn = f"⚠ 本段已跳过（{e}）请另选一个挂点"
                    continue           # 不入链，仍以原挂点为起点
                poly, mode = draw_catenary_2d(self.doc, self.ms, pts)
                # 最低点倒三角标记（顶点＝最低点，画在曲线上方；随比例缩放）
                mark, _mk_mode = draw_low_point_marker(
                    self.doc, self.ms, info["marker_pts"])
                segs.append((info, poly.Handle, mode, p_prev, p_next, mark.Handle))
                p_prev = p_next
                self.log_msg(
                    f"  段{len(segs)}: 实际档距={info['L_real']:.2f}m  "
                    f"实际高差={info['h_real']:.2f}m  实际弧垂={info['sag_real']:.3f}m  "
                    f"图上弧垂={info['sag_plot']:.2f}  Handle={poly.Handle}")
                self.log_msg("     " + _low_point_note(info))
        except Exception as e:
            err = e
        finally:
            self.root.deiconify(); self.root.lift()

        if warn:
            self.log_msg(warn)
        if err is not None:
            messagebox.showerror("生成失败", str(err))
            self.log_msg(f"✘ 生成失败: {err}")
            return
        if not segs:
            if not warn:          # warn 已在上方记录过，避免重复
                self.log_msg("✘ 已取消点选（未生成曲线）")
            return

        # 断面图通常已有整幅图框，缩放会打断视图 —— 有意不调 ZoomExtents
        last_info, last_handle, last_mode = segs[-1][0], segs[-1][1], segs[-1][2]
        last_mark = segs[-1][5]
        tip = "" if last_mode == "lwpolyline" else "（本 CAD 无二维多段线接口，已退化为三维多段线 Z=0）"
        total_L = sum(s[0]["L_real"] for s in segs)
        msg = (f"{param_desc}{tip}  共 {len(segs)} 段\n"
               f"累计实际档距={total_L:.2f}m  末段: 实际高差={last_info['h_real']:.2f}m  "
               f"实际弧垂={last_info['sag_real']:.3f}m  图上弧垂={last_info['sag_plot']:.2f}  "
               f"a={last_info['a']:.1f}m")
        # ⑥ 块下方的说明栏是**固定用法说明**，不随生成结果变化 —— 结果只写日志区
        self.log_msg("✔ 已连续生成二维断面弧垂曲线: " + msg.replace("\n", "  "))
        self.log_msg(f"  末段 Handle={last_handle}  倒三角标记 Handle={last_mark}")

    # ---- 工具 -------------------------------------------------------
    def _is_user_cancel(self, e):
        """判断 COM 异常是否由用户取消 (ESC / 右键) 引起。"""
        return _is_cancel_error(e)

    def _pick_abort(self, e, what="点选"):
        """拾取阶段的异常统一收尾（①②③④⑤⑥ 共用）—— 按 ESC 取消不再"报错"。

        背景：各 CAD 的取消文案不统一（"用户取消" / "已取消" / Cancel / ESC…），
        只靠文案判定必然漏判，漏判就会把「用户取消」误报成「点选失败」（历史痛点）。
        故此处对**任何**拾取异常都按「用户主动中止」处理：
        **不弹错误框**，只在日志里留一句（非取消类附原始文案备查），调用方直接 return。
        """
        if _is_cancel_error(e):
            self.log_msg(f"✘ {what}已取消，未生成任何图元")
        else:
            self.log_msg(f"✘ {what}已中止（{e}），未生成任何图元")

    def _ensure_conn(self):
        if self.doc is None:
            self.connect()
            return self.doc is not None
        self._follow_foreground_cad()
        return self.doc is not None

    def _follow_foreground_cad(self):
        """自动模式下：用户把另一个 CAD 实例切到前台时，自动把连接切过去。

        多 CAD / 同 CAD 多版本共存时的必要动作 —— 否则连上 CAD A 之后，
        用户切到 CAD B 的窗口再点功能按钮，仍会画进 A。
        只在「自动」模式下跟随；用户手动指定了目标 CAD 就尊重其选择。
        跟随重连失败时保留原连接、不弹错误框（不该打断操作）。
        """
        if self.cad_prefer_var.get() != "自动" or not self._conn_hwnd:
            return
        try:
            tgt = foreground_cad_target(self.root.winfo_id())
        except Exception:
            return
        if not tgt:
            return  # 探测不到任何 CAD 实例 → 保持现状
        key, hwnd, progid = tgt
        if hwnd == self._conn_hwnd:
            return  # 前台仍是已连的那个实例，无需动作
        # 前台换成了另一个实例 → 先连新的，成功才替换
        old = (self.gc, self.doc, self.ms, self.app_name, self._conn_hwnd)
        self.log_msg(f"→ 前台已切到 {cad_label_of(key)}（{progid}），跟随重连…")
        if not self.connect(silent=True):
            self.gc, self.doc, self.ms, self.app_name, self._conn_hwnd = old
            self.status_var.set(f"已连接 {self.app_name} {self.doc.Name} ✓")
            self.log_msg("  （跟随重连失败，继续使用原连接）")


def main():
    root = tk.Tk()
    CatenaryApp(root)
    root.mainloop()


# ───────────────────────────────────────────────────────────────────
# 命令行模式 (CLI)
# ───────────────────────────────────────────────────────────────────

def _cli_pick(fn, *a, **kw):
    """CLI 交互拾取的统一包装（①~⑥ 共用）—— 按 ESC 取消不再抛 traceback。

    各 CAD 的取消文案不统一，无法只靠文案保证命中，故非取消类异常也一并
    转成一句可读信息（附原文），避免用户面对满屏栈帧。
    """
    try:
        return fn(*a, **kw)
    except Exception as e:
        if _is_cancel_error(e):
            sys.exit("已取消，未生成任何图元")
        sys.exit(f"✘ 拾取失败：{e}")


def _cli_draw(args):
    gc, doc, ms, _app = get_gcad_or_elevate(getattr(args, "cad", None))
    if args.pick:
        print("→ 在 CAD 中点选第1个悬挂点 …", flush=True)
        p1 = _cli_pick(pick_point, doc, "\n选择第1个悬挂点: ")
        print("→ 在 CAD 中点选第2个悬挂点 …", flush=True)
        p2 = _cli_pick(pick_point, doc, "\n选择第2个悬挂点: ")
    else:
        p1 = tuple(float(v) for v in args.p1)
        p2 = tuple(float(v) for v in args.p2)
    # K 值 或 σ₀+γ 二选一：K = γ/(8σ₀)
    if args.K is not None:
        K = float(args.K) * 1e-5
        param_desc = f"K={args.K}×10⁻⁵"
    elif args.sigma is not None and args.gamma is not None:
        if args.sigma <= 0:
            sys.exit("✘ 应力σ₀必须为正数")
        K = args.gamma / (8.0 * args.sigma)
        param_desc = f"σ₀={args.sigma}N/mm², γ={args.gamma}N/(m·mm²) → K={K * 1e5:.2f}×10⁻⁵"
    else:
        sys.exit("✘ 必须提供 --K 或 (--sigma 与 --gamma)")
    n = args.n
    # 统一 cosh + 牛顿迭代求解（含高差修正）
    pts, sag, L, a, x0 = catenary_3d(p1, p2, K, n)
    poly = draw_catenary(doc, ms, pts)
    gc.ZoomExtents()
    print("✔ 已生成三维弧垂曲线")
    print(f"  {param_desc}")
    print(f"  P1={tuple(round(v, 3) for v in p1)}")
    print(f"  P2={tuple(round(v, 3) for v in p2)}")
    print(f"  档距={L:.2f}m  高差={p2[2] - p1[2]:.2f}m  弧垂={sag:.3f}m")
    print(f"  Handle={poly.Handle}")


def _cli_draw2d(args):
    """二维断面悬链线：图上坐标 ÷ 比例 → 实际米 → 算 K 值悬链线 → × 比例画回图上。

    交互模式（--pick）支持**连续点选**：点第 2 点生成第 1 段后不退出，继续点
    第 3 点即以「上一段终点」为新起点生成下一段，K 值与比例全程相同；
    按 ESC / 右键结束。端点不合法（如 X 相同）的段自动跳过并提示。
    """
    gc, doc, ms, _app = get_gcad_or_elevate(getattr(args, "cad", None))
    K = float(args.K) * 1e-5          # 只做 K 值选项，不做 σ₀+γ
    rx, ry = float(args.rx), float(args.ry)
    if rx <= 0 or ry <= 0:
        sys.exit("✘ 图纸比例分母必须为正数")
    layer = args.layer or LAYER_NAME
    segs = []                          # [(info, handle, mode, p_from, p_to), ...]

    def _draw_seg(p_from, p_to):
        """生成一段；起止点不合法时抛 ValueError，由调用方决定跳过还是退出。"""
        pts, info = catenary_section_2d(p_from, p_to, K, rx, ry, args.n)
        poly, mode = draw_catenary_2d(doc, ms, pts, layer=layer)
        # 最低点倒三角标记（顶点＝最低点，画在曲线上方；随比例缩放）
        mark, _mk_mode = draw_low_point_marker(doc, ms, info["marker_pts"], layer=layer)
        segs.append((info, poly.Handle, mode, p_from, p_to, mark.Handle))
        print(f"✔ 段{len(segs)}: 实际档距={info['L_real']:.2f}m  "
              f"实际高差={info['h_real']:.2f}m  实际弧垂={info['sag_real']:.3f}m  "
              f"图上弧垂={info['sag_plot']:.2f}   Handle={poly.Handle}", flush=True)
        print("    " + _low_point_note(info) + f"  标记Handle={mark.Handle}", flush=True)

    if args.pick:
        print("→ 在 CAD 中连续点选挂点（每点一次生成一段，ESC/右键结束）…", flush=True)
        try:
            p_prev = pick_point(doc, "\n选择第1个挂点: ")
        except Exception as e:
            if _is_cancel_error(e):
                print("已取消，未生成曲线")
                return
            raise
        while True:
            prompt = ("\n选择第2个挂点: " if not segs
                      else "\n选择下一个挂点 (ESC/右键结束): ")
            try:
                p_next = pick_point(doc, prompt)
            except Exception as e:
                if _is_cancel_error(e):
                    break              # ESC / 右键 = 正常结束连续点选
                # 兜底：取点失败一律结束链路（已画的段保留），不依赖取消文案
                print(f"⚠ 连续点选已结束（取点失败：{e}），已生成的曲线保留", flush=True)
                break
            try:
                _draw_seg(p_prev, p_next)
            except ValueError as e:
                print(f"⚠ 本段已跳过（{e}）请另选一个挂点", flush=True)
                continue               # 起止点不变，等待下一个挂点
            p_prev = p_next
        if not segs:
            print("未生成任何曲线（已取消）")
            return
    else:
        p1 = tuple(float(v) for v in args.p1)
        p2 = tuple(float(v) for v in args.p2)
        try:
            _draw_seg(p1, p2)
        except ValueError as e:
            sys.exit(f"✘ {e}")

    # 断面图通常已有整幅图框，缩放会打断视图 —— 有意不调 ZoomExtents
    last_info, last_handle, last_mode = segs[-1][0], segs[-1][1], segs[-1][2]
    print("✔ 已生成二维断面弧垂曲线")
    print(f"  K={args.K}×10⁻⁵  横向1:{rx:g}  纵向1:{ry:g}  "
          f"(scale_x={last_info['scale_x']:g}  scale_y={last_info['scale_y']:g})")
    print(f"  P1=({segs[0][3][0]:.3f}, {segs[0][3][1]:.3f})   "
          f"P末=({segs[-1][4][0]:.3f}, {segs[-1][4][1]:.3f})")
    print(f"  共 {len(segs)} 段  累计实际档距={sum(s[0]['L_real'] for s in segs):.2f}m")
    print(f"  末段: 实际高差={last_info['h_real']:.2f}m  实际弧垂={last_info['sag_real']:.3f}m  "
          f"图上弧垂={last_info['sag_plot']:.2f}")
    n_note = (f"{last_info['n_pts']}→{last_info['n_pts_used']}（延伸后加密）"
              if last_info["extended"] else f"{last_info['n_pts']}")
    print(f"  悬链线参数 a={last_info['a']:.2f}m  采样点数={n_note}")
    print(f"  {_low_point_note(last_info)}")
    print(f"  图层={layer}  类型="
          + ("二维多段线(AcDbPolyline)" if last_mode == "lwpolyline"
             else "三维多段线(AcDb3dPolyline, Z=0 兜底)"))
    print(f"  末段 Handle={last_handle}  倒三角标记 Handle={segs[-1][5]}")
    print(f"  倒三角标记: 顶点=最低点 顶边={last_info['marker_w']:g} 高={last_info['marker_h']:g} "
          f"（比例 1:{rx:g}/1:{ry:g}；随比例等比缩放）")


def _cli_mindist(args):
    gc, doc, ms, _app = get_gcad_or_elevate(getattr(args, "cad", None))
    if args.pick:
        print("→ 在 CAD 中点选第1条曲线 …", flush=True)
        e1 = _cli_pick(pick_entity, doc, "\n选择第1条曲线: ", _app)
        print("→ 在 CAD 中点选第2条曲线 …", flush=True)
        e2 = _cli_pick(pick_entity, doc, "\n选择第2条曲线: ", _app)
    else:
        e1 = get_entity_by_handle(ms, args.h1)
        e2 = get_entity_by_handle(ms, args.h2)
    pts1 = extract_points(e1)
    pts2 = extract_points(e2)
    d, c1, c2 = closest_distance(pts1, pts2)
    draw_distance_line(doc, ms, c1, c2)
    draw_distance_label(doc, ms, c1, c2, f"{d:.2f}m")
    gc.ZoomExtents()
    print(f"✔ 最小距离 = {d:.3f} m")
    print(f"  曲线1最近点 ({c1[0]:.3f}, {c1[1]:.3f}, {c1[2]:.3f})")
    print(f"  曲线2最近点 ({c2[0]:.3f}, {c2[1]:.3f}, {c2[2]:.3f})")


def _cli_wind(args):
    gc, doc, ms, _app = get_gcad_or_elevate(getattr(args, "cad", None))
    if args.handle:
        ent = get_entity_by_handle(ms, args.handle)
        print(f"→ 已按句柄 {args.handle} 定位曲线")
    else:
        print("→ 在 CAD 中点选一条弧垂曲线 …", flush=True)
        ent = _cli_pick(pick_entity, doc, "\n请选择一条弧垂曲线: ", _app)
        print(f"→ 已拾取曲线 (句柄 {ent.Handle})")
    if args.angle is None:
        angle = _cli_pick(pick_angle, doc, "\n请输入风偏角(度, 右偏为正): ")
    else:
        angle = args.angle
    print(f"→ 风偏角 = {angle}°")
    pts = extract_points(ent)
    new_pts = rotate_curve(pts, angle)
    poly = draw_wind_curve(doc, ms, new_pts)
    label, label_text = draw_wind_label(doc, ms, new_pts, angle)
    gc.ZoomExtents()
    print(f"✔ 已生成风偏弧垂曲线: Handle={poly.Handle}, 蓝色, 图层'{LAYER_WIND}'")
    print(f"✔ 文字标注: '{label_text}' (Handle={label.Handle}, 按绝对值标示, 忽略正负号)")


def _cli_bundle(args):
    gc, doc, ms, _app = get_gcad_or_elevate(getattr(args, "cad", None))
    if args.handle:
        ent = get_entity_by_handle(ms, args.handle)
        print(f"→ 已按句柄 {args.handle} 定位中心线")
    else:
        print("→ 在 CAD 中点选中心线 …", flush=True)
        ent = _cli_pick(pick_entity, doc, "\n请选择中心线(弧垂曲线): ", _app)
        print(f"→ 已拾取中心线 (句柄 {ent.Handle})")
    d_m = args.d / 1000.0  # 用户单位 mm → 绘图单位 m
    mode_map = {"horiz": BUNDLE_HORIZ, "vert": BUNDLE_VERT, "quad": BUNDLE_QUAD}
    mode_code = mode_map[args.mode]
    pts = extract_points(ent)
    lines = generate_bundle_lines(pts, d_m, mode_code)
    handles = []
    for line in lines:
        poly = draw_bundle_line(doc, ms, line)
        handles.append(poly.Handle)
    gc.ZoomExtents()
    mode_name = {"horiz": "水平双分裂", "vert": "垂直双分裂", "quad": "四分裂"}[args.mode]
    print(f"✔ 已生成{mode_name}: 间距 {args.d}mm (= {d_m:g} m), 共 {len(handles)} 条子导线")
    print(f"  中心线 Handle={ent.Handle}")
    print(f"  子导线 Handles={handles}")


def _cli_suspension(args):
    gc, doc, ms, _app = get_gcad_or_elevate(getattr(args, "cad", None))
    if args.pick:
        print("→ 在 CAD 中点选第1个悬挂点 …", flush=True)
        p1 = _cli_pick(pick_point, doc, "\n选择第1个悬挂点: ")
        print("→ 在 CAD 中点选第2个悬挂点 …", flush=True)
        p2 = _cli_pick(pick_point, doc, "\n选择第2个悬挂点: ")
    else:
        p1 = tuple(float(v) for v in args.p1)
        p2 = tuple(float(v) for v in args.p2)
    # K 值 或 σ₀+γ 二选一：K = γ/(8σ₀)
    if args.K is not None:
        K = float(args.K) * 1e-5
        param_desc = f"K={args.K}×10⁻⁵"
    elif args.sigma is not None and args.gamma is not None:
        if args.sigma <= 0:
            sys.exit("✘ 应力σ₀必须为正数")
        K = args.gamma / (8.0 * args.sigma)
        param_desc = f"σ₀={args.sigma}N/mm², γ={args.gamma}N/(m·mm²) → K={K * 1e5:.2f}×10⁻⁵"
    else:
        sys.exit("✘ 必须提供 --K 或 (--sigma 与 --gamma)")
    n = args.n
    D1 = args.d1 if args.s1 else None
    D2 = args.d2 if args.s2 else None
    # 串风偏角仅在 --wind 时生效；未勾选一律按 0（竖直串）
    A1 = args.a1 if (args.wind and D1 is not None) else 0.0
    A2 = args.a2 if (args.wind and D2 is not None) else 0.0
    B = args.b if args.wind else 0.0

    geo = suspension_geometry(p1, p2, K, D1, A1, D2, A2, B, n)
    wind_mode = geo["wind_pts"] is not None
    susp_handles = []
    if wind_mode:
        # 摆串 → 浅蓝图层
        for top, bottom in geo["susp_lines"]:
            line = draw_suspension_line(
                doc, ms, top, bottom,
                color=COLOR_SUSPENSION_WIND, layer=LAYER_SUSPENSION_WIND)
            susp_handles.append(line.Handle)
    else:
        for top, bottom in geo["susp_lines"]:
            line = draw_suspension_line(doc, ms, top, bottom)
            susp_handles.append(line.Handle)
    # 风偏模式：只画风偏线（蓝），不再画未风偏的参照曲线（红基准）与
    # 未风偏的参照绝缘子（A=0 品红竖直串）；非风偏才画红基准弧垂
    wind_poly = None
    if geo["wind_pts"] is not None:
        wind_poly = draw_wind_curve(doc, ms, geo["wind_pts"])
    else:
        base_poly = draw_catenary(doc, ms, geo["base_pts"])
    gc.ZoomExtents()

    print(f"✔ 已生成含悬垂串弧垂线: {param_desc}")
    print(f"  档距={geo['L']:.2f}m 弧垂={geo['sag']:.3f}m")
    print(f"  端点1={tuple(round(v,3) for v in geo['ep1'])}  端点2={tuple(round(v,3) for v in geo['ep2'])}")
    if not wind_mode:
        print(f"  基准弧垂 Handle={base_poly.Handle}")
    if susp_handles:
        tag = "风偏串" if wind_mode else "悬垂串"
        layer = LAYER_SUSPENSION_WIND if wind_mode else LAYER_SUSPENSION
        color = "浅蓝" if wind_mode else "品红"
        print(f"  {tag} Handles={susp_handles} (图层'{layer}', {color})")
    if wind_poly is not None:
        print(f"  风偏线 Handle={wind_poly.Handle} (B={B:g}°, 右偏为正)")


def cli_main():
    # ⚠️ Windows 控制台默认 GBK：CLI 输出含 ✔ ✘ → 等 Unicode 符号会
    # UnicodeEncodeError 崩溃。保持原编码（中文正常显示），仅加容错：
    # 无法编码的符号（✔✘→）替换为 '?'，不崩溃、不乱码。
    for _s in (sys.stdout, sys.stderr):
        if _s is not None and hasattr(_s, "reconfigure"):
            try:
                _s.reconfigure(errors="replace")
            except Exception:
                pass
    import argparse
    parent = argparse.ArgumentParser(add_help=False)
    parent.add_argument("--cad", choices=cad_family_keys(), default=None,
                        help="强制指定 CAD (默认自动检测: 优先已运行的实例)")
    parser = argparse.ArgumentParser(
        description=f"空间弧垂曲线工具 (支持 {_all_cad_names()})")
    sub = parser.add_subparsers(dest="cmd")

    d = sub.add_parser("draw", parents=[parent], help="生成空间弧垂曲线")
    d.add_argument("--K", default=None, help="K值(省略10⁻⁵)，如 15.22（与--sigma/--gamma二选一）")
    d.add_argument("--sigma", default=None, type=float, help="导线应力σ₀(N/mm²)")
    d.add_argument("--gamma", default=None, type=float, help="导线比载γ(N/(m·mm²))")
    d.add_argument("--n", type=int, default=200, help="采样点数 (默认200)")
    d.add_argument("--pick", action="store_true", help="在CAD中交互点选两点")
    d.add_argument("--p1", nargs=3, metavar=("X", "Y", "Z"), help="第1点坐标")
    d.add_argument("--p2", nargs=3, metavar=("X", "Y", "Z"), help="第2点坐标")

    d2 = sub.add_parser("draw2d", parents=[parent],
                        help="二维断面悬链线（平面图，带图纸比例）")
    d2.add_argument("--K", required=True, help="K值(省略10⁻⁵)，如 15.22")
    d2.add_argument("--rx", type=float, default=5000.0, help="横向出图比例分母，默认 5000（1:5000）")
    d2.add_argument("--ry", type=float, default=500.0, help="纵向出图比例分母，默认 500（1:500）")
    d2.add_argument("--n", type=int, default=200, help="采样点数 (默认200)")
    d2.add_argument("--pick", action="store_true",
                    help="在CAD中交互点选挂点（连续点选：每点一次生成一段，ESC/右键结束）")
    d2.add_argument("--p1", nargs=2, metavar=("X", "Y"), help="第1点图上坐标（二维）")
    d2.add_argument("--p2", nargs=2, metavar=("X", "Y"), help="第2点图上坐标（二维）")
    d2.add_argument("--layer", default=None, help=f"图层名 (默认 {LAYER_NAME})")

    m = sub.add_parser("mindist", parents=[parent], help="两条曲线最小距离")
    m.add_argument("--pick", action="store_true", help="在CAD中交互选两条曲线")
    m.add_argument("--h1", help="第1条曲线句柄")
    m.add_argument("--h2", help="第2条曲线句柄")

    w = sub.add_parser("wind", parents=[parent], help="风偏弧垂曲线")
    w.add_argument("--pick", action="store_true", help="在CAD中交互点选曲线")
    w.add_argument("--handle", help="曲线句柄")
    w.add_argument("--angle", type=float, default=None, help="风偏角(度, 右偏为正)，省略则在CAD命令行输入")

    b = sub.add_parser("bundle", parents=[parent], help="分裂导线（双分裂/四分裂）")
    b.add_argument("--pick", action="store_true", help="在CAD中交互点选中心线")
    b.add_argument("--handle", help="中心线句柄")
    b.add_argument("--d", type=float, required=True, help="分裂间距(mm)，如 600")
    b.add_argument("--mode", choices=["horiz", "vert", "quad"], default="quad",
                   help="分裂模式: horiz=水平双分裂 vert=垂直双分裂 quad=四分裂 (默认quad)")

    s = sub.add_parser("suspension", parents=[parent], help="含悬垂绝缘子串的弧垂线")
    s.add_argument("--K", default=None, help="K值(省略10⁻⁵)，如 15.22（与--sigma/--gamma二选一）")
    s.add_argument("--sigma", default=None, type=float, help="导线应力σ₀(N/mm²)")
    s.add_argument("--gamma", default=None, type=float, help="导线比载γ(N/(m·mm²))")
    s.add_argument("--n", type=int, default=200, help="采样点数 (默认200)")
    s.add_argument("--pick", action="store_true", help="在CAD中交互点选两点")
    s.add_argument("--p1", nargs=3, metavar=("X", "Y", "Z"), help="第1点坐标")
    s.add_argument("--p2", nargs=3, metavar=("X", "Y", "Z"), help="第2点坐标")
    s.add_argument("--s1", action="store_true", help="第1点有悬垂串")
    s.add_argument("--d1", type=float, default=3.0, help="第1点串长D1(m)")
    s.add_argument("--a1", type=float, default=0.0, help="第1点串风偏角A1(°)，右偏为正")
    s.add_argument("--s2", action="store_true", help="第2点有悬垂串")
    s.add_argument("--d2", type=float, default=3.0, help="第2点串长D2(m)")
    s.add_argument("--a2", type=float, default=0.0, help="第2点串风偏角A2(°)，右偏为正")
    s.add_argument("--wind", action="store_true", help="绘制风偏线")
    s.add_argument("--b", type=float, default=0.0, help="线风偏角B(°)，右偏为正")

    sub.add_parser("check", parents=[parent], help="检查本机 CAD COM 注册状态")

    args = parser.parse_args()
    try:
        if args.cmd == "check":
            print(diagnose_cad_connection(args.cad))
            return
        if args.cmd == "draw":
            if not args.pick and (not args.p1 or not args.p2):
                parser.error("draw 需 --pick 或同时提供 --p1 --p2")
            _cli_draw(args)
        elif args.cmd == "mindist":
            if not args.pick and (not args.h1 or not args.h2):
                parser.error("mindist 需 --pick 或同时提供 --h1 --h2")
            _cli_mindist(args)
        elif args.cmd == "wind":
            if not args.pick and not args.handle:
                parser.error("wind 需 --pick 或 --handle")
            _cli_wind(args)
        elif args.cmd == "bundle":
            if not args.pick and not args.handle:
                parser.error("bundle 需 --pick 或 --handle")
            _cli_bundle(args)
        elif args.cmd == "suspension":
            if not args.pick and (not args.p1 or not args.p2):
                parser.error("suspension 需 --pick 或同时提供 --p1 --p2")
            _cli_suspension(args)
        elif args.cmd == "draw2d":
            if not args.pick and (not args.p1 or not args.p2):
                parser.error("draw2d 需 --pick 或同时提供 --p1 X Y --p2 X Y")
            _cli_draw2d(args)
        else:
            parser.print_help()
    except RuntimeError as e:
        # get_gcad 已经把 HRESULT 转成中文说明并附带诊断
        print(f"错误：{e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"未预期的错误：{e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("draw", "draw2d", "mindist", "wind",
                                             "bundle", "suspension", "check"):
        cli_main()
    else:
        main()
