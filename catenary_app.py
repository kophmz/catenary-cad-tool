#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
空间弧垂曲线工具 (CAD 通用版 · 支持 GstarCAD / AutoCAD)
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
  - pywin32 (win32com)      —— 直连 CAD COM (GstarCAD / AutoCAD)
  - tkinter (标准库)        —— 界面
  - CAD 已启动并打开图纸 (GstarCAD 或 AutoCAD)

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
    foreground_cad_prefer,
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
    draw_distance_line,
    draw_distance_label,
    draw_wind_curve,
    draw_wind_label,
    draw_bundle_line,
    draw_suspension_line,
    # 图层/颜色常量（全项目唯一出处）
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


# ───────────────────────────────────────────────────────────────────
# GUI
# ───────────────────────────────────────────────────────────────────

class CatenaryApp:
    def __init__(self, root):
        self.root = root
        root.title("空间弧垂曲线工具 · CAD · v0.21")
        root.geometry("700x900")
        root.resizable(False, False)

        self.gc = self.doc = self.ms = None
        self.app_name = None  # 连接后记录的 CAD 名称（AutoCAD / GstarCAD）

        self._build_ui()

    # ---- 界面 -------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # 连接状态
        f0 = ttk.LabelFrame(self.root, text="连接")
        f0.pack(fill="x", **pad)
        self.status_var = tk.StringVar(value="未连接")
        ttk.Label(f0, textvariable=self.status_var).pack(side="left")
        # 目标 CAD 选择器：自动检测 / 强制 GstarCAD / 强制 AutoCAD
        # （自动模式下先尝试 GstarCAD，再 AutoCAD；两者都开着时用此项强制 AutoCAD）
        self.cad_prefer_var = tk.StringVar(value="自动")
        sel = ttk.Frame(f0)
        sel.pack(side="left", padx=8)
        ttk.Label(sel, text="目标:").pack(side="left")
        ttk.OptionMenu(
            sel, self.cad_prefer_var, "自动", "自动", "GstarCAD", "AutoCAD"
        ).pack(side="left")
        ttk.Button(f0, text="使用说明", command=self.open_help).pack(side="right", padx=(0,4))
        ttk.Button(f0, text="连接/刷新 CAD窗口", command=self.connect).pack(side="right")

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

    # ---- 连接 -------------------------------------------------------
    def connect(self):
        # 将下拉选择映射到 get_gcad 的 prefer 参数
        _prefer_map = {"自动": None, "GstarCAD": "gstar", "AutoCAD": "autocad"}
        prefer = _prefer_map.get(self.cad_prefer_var.get(), None)
        if prefer is None:
            # 自动：若某 CAD 确为前台则用它；否则 prefer 保持 None，
            # 交给 get_gcad 走「AutoCAD 优先 + GstarCAD 兜底」
            fg = foreground_cad_prefer(self.root.winfo_id())
            prefer = fg
            if prefer:
                self.log_msg(f"→ 自动检测到：{prefer}")
            else:
                self.log_msg("→ 自动检测：未发现活跃CAD窗口，默认AutoCAD优先")
        try:
            self.gc = self.doc = self.ms = None  # 先清空旧连接，确保真正刷新
            self.gc, self.doc, self.ms, app = get_gcad(prefer)
            self.app_name = app  # 供后续 pick_entity 按 CAD 类型选择调用方式
            # doc.Name 形如 "Drawing1"（未保存）或完整路径；直接展示连接的窗口
            self.status_var.set(f"已连接 {app} {self.doc.Name} ✓")
            self.log_msg(f"✔ 已连接 {app} {self.doc.Name}（目标={self.cad_prefer_var.get()}）")
        except Exception as e:
            self.status_var.set("连接失败")
            messagebox.showerror("连接失败", f"无法连接 CAD（{self.cad_prefer_var.get()}）：\n{e}")
            self.log_msg(f"✘ 连接失败: {e}")

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
            if self._is_user_cancel(e):
                messagebox.showinfo("已取消", "点选已取消。")
                self.log_msg("✘ 已取消点选")
                return
            messagebox.showerror("点选失败", str(e))
            self.log_msg(f"✘ 点选失败: {e}")
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
            if self._is_user_cancel(e):
                messagebox.showinfo("已取消", "选取已取消。")
                return
            messagebox.showerror("选取失败", str(e))
            self.log_msg(f"✘ 选取失败: {e}")
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
            if self._is_user_cancel(e):
                messagebox.showinfo("已取消", "选取已取消。")
                self.log_msg("✘ 已取消点选")
                return
            messagebox.showerror("选取失败", str(e))
            self.log_msg(f"✘ 选取失败: {e}")
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
            if self._is_user_cancel(e):
                messagebox.showinfo("已取消", "选取已取消。")
                self.log_msg("✘ 已取消点选")
                return
            messagebox.showerror("选取失败", str(e))
            self.log_msg(f"✘ 选取失败: {e}")
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
            if self._is_user_cancel(e):
                messagebox.showinfo("已取消", "点选已取消。")
                self.log_msg("✘ 已取消点选")
                return
            messagebox.showerror("点选失败", str(e))
            self.log_msg(f"✘ 点选失败: {e}")
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

    # ---- 工具 -------------------------------------------------------
    def _is_user_cancel(self, e):
        """判断 COM 异常是否由用户取消 (ESC / 右键) 引起。"""
        s = str(e).lower()
        return any(k in s for k in ("cancel", "已取消", "userinterrupt",
                                    "user cancel", "discard"))

    def _ensure_conn(self):
        if self.doc is None:
            self.connect()
        if self.doc is None:
            return False
        return True


def main():
    root = tk.Tk()
    CatenaryApp(root)
    root.mainloop()


# ───────────────────────────────────────────────────────────────────
# 命令行模式 (CLI)
# ───────────────────────────────────────────────────────────────────

def _cli_draw(args):
    gc, doc, ms, _app = get_gcad(getattr(args, "cad", None))
    if args.pick:
        print("→ 在 CAD 中点选第1个悬挂点 …", flush=True)
        p1 = pick_point(doc, "\n选择第1个悬挂点: ")
        print("→ 在 CAD 中点选第2个悬挂点 …", flush=True)
        p2 = pick_point(doc, "\n选择第2个悬挂点: ")
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


def _cli_mindist(args):
    gc, doc, ms, _app = get_gcad(getattr(args, "cad", None))
    if args.pick:
        print("→ 在 CAD 中点选第1条曲线 …", flush=True)
        e1 = pick_entity(doc, "\n选择第1条曲线: ", _app)
        print("→ 在 CAD 中点选第2条曲线 …", flush=True)
        e2 = pick_entity(doc, "\n选择第2条曲线: ", _app)
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
    gc, doc, ms, _app = get_gcad(getattr(args, "cad", None))
    if args.handle:
        ent = get_entity_by_handle(ms, args.handle)
        print(f"→ 已按句柄 {args.handle} 定位曲线")
    else:
        print("→ 在 CAD 中点选一条弧垂曲线 …", flush=True)
        ent = pick_entity(doc, "\n请选择一条弧垂曲线: ", _app)
        print(f"→ 已拾取曲线 (句柄 {ent.Handle})")
    if args.angle is None:
        angle = pick_angle(doc, "\n请输入风偏角(度, 右偏为正): ")
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
    gc, doc, ms, _app = get_gcad(getattr(args, "cad", None))
    if args.handle:
        ent = get_entity_by_handle(ms, args.handle)
        print(f"→ 已按句柄 {args.handle} 定位中心线")
    else:
        print("→ 在 CAD 中点选中心线 …", flush=True)
        ent = pick_entity(doc, "\n请选择中心线(弧垂曲线): ", _app)
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
    gc, doc, ms, _app = get_gcad(getattr(args, "cad", None))
    if args.pick:
        print("→ 在 CAD 中点选第1个悬挂点 …", flush=True)
        p1 = pick_point(doc, "\n选择第1个悬挂点: ")
        print("→ 在 CAD 中点选第2个悬挂点 …", flush=True)
        p2 = pick_point(doc, "\n选择第2个悬挂点: ")
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
    parent.add_argument("--cad", choices=["gstar", "autocad"], default=None,
                        help="强制指定 CAD (默认自动检测: 先 GstarCAD 后 AutoCAD)")
    parser = argparse.ArgumentParser(description="空间弧垂曲线工具 (支持 GstarCAD / AutoCAD)")
    sub = parser.add_subparsers(dest="cmd")

    d = sub.add_parser("draw", parents=[parent], help="生成空间弧垂曲线")
    d.add_argument("--K", default=None, help="K值(省略10⁻⁵)，如 15.22（与--sigma/--gamma二选一）")
    d.add_argument("--sigma", default=None, type=float, help="导线应力σ₀(N/mm²)")
    d.add_argument("--gamma", default=None, type=float, help="导线比载γ(N/(m·mm²))")
    d.add_argument("--n", type=int, default=200, help="采样点数 (默认200)")
    d.add_argument("--pick", action="store_true", help="在CAD中交互点选两点")
    d.add_argument("--p1", nargs=3, metavar=("X", "Y", "Z"), help="第1点坐标")
    d.add_argument("--p2", nargs=3, metavar=("X", "Y", "Z"), help="第2点坐标")

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

    args = parser.parse_args()
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
    else:
        parser.print_help()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("draw", "mindist", "wind", "bundle", "suspension"):
        cli_main()
    else:
        main()
