#!/usr/bin/env python3
"""
GstarCAD 三维悬链线绘制工具（win32com 直连版）

连接 GstarCAD / AutoCAD（COM 直连），根据 K 值、两悬挂点三维坐标绘制
1:1（米）悬链线多段线。

⚠️ 历史修复（审查 Bug 1）：早期版本 import `mcp__gstarcad`（MCP 工具
命名空间，不可作为 Python 模块导入），绘制路径永远 ImportError。现改为
复用 catenary_cad.py 的 win32com 直连封装（与主程序 catenary_app.py
同一套代码）。

用法:
    # 默认：K=8e-5, 档距 300m, 等高
    python catenary_cad_draw.py 8e-5

    # 指定悬挂点
    python catenary_cad_draw.py 8e-5 --p1 0 0 100 --p2 300 0 100

    # 不等高 + 3D 方向 + 采样点数
    python catenary_cad_draw.py 6e-5 --p1 0 0 50 --p2 400 100 65 --n-pts 500

    # 使用 JSON 参数文件
    python catenary_cad_draw.py --param-file example.json

    # 仅计算不绘制（预览 / 无 CAD 环境）
    python catenary_cad_draw.py 8e-5 --no-cad

参考:
    《电力工程设计手册 架空输电线路设计》式 14-1 / 14-2
    K = γ/(8σ), 典型值 4e-5 ~ 15e-5 (1/m)
"""

import sys
import os
import json
import argparse

from catenary_core import catenary_3d, print_summary, format_gstarcad_polyline
from catenary_cad import get_gcad, ensure_layer, draw_catenary

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def draw_catenary_in_gstarcad(points_3d, layer: str = "悬链线",
                               color: int = 1, lineweight: int = 30):
    """调用 win32com 直连 CAD 绘制三维悬链线多段线（修复 Bug 1）。

    需要: 本机已安装 CAD（GstarCAD / AutoCAD）并至少打开一个图纸。

    参数:
        points_3d: [(x,y,z), ...] 三维点序列
        layer: 图层名
        color: ACI 颜色号
        lineweight: 线宽 (整数, 0=默认；win32com 绘制多段线时仅用于图层创建)

    Returns:
        bool 是否绘制成功
    """
    try:
        gc, doc, ms, app_name = get_gcad(None)
    except Exception as e:
        print(f"❌ 无法连接 CAD：{e}")
        print("   请确认已启动 GstarCAD / AutoCAD 并打开图纸。")
        print("   替代方案: 使用 --no-cad 或 --cad-json 导出点集后手动导入。")
        return False

    print(f"✅ 已连接: {app_name}, 文档: {doc.Name}")

    # 1. 创建/切换图层（ensure_layer 不存在时自动创建并设色）
    try:
        ensure_layer(doc, layer, color)
        print(f"📐 图层: {layer} (ACI {color})")
    except Exception as e:
        print(f"⚠️  图层操作警告: {e}")

    # 2. 绘制三维多段线（复用主程序 draw_catenary，保证真 3D）
    try:
        poly = draw_catenary(doc, ms, points_3d, color=color)
        print(f"✏️  正在绘制三维悬链线 ({len(points_3d)} 个顶点)...")
        print(f"✅ 悬链线绘制完成, Handle={poly.Handle}")
    except Exception as e:
        print(f"❌ 绘制失败: {e}")
        return False

    # 3. 缩放全图
    try:
        gc.ZoomExtents()
        print("🔍 已缩放至全图范围")
    except Exception:
        pass

    return True


# ─── 参数文件格式 ──────────────────────────────────────────────────

SAMPLE_PARAM_JSON = {
    "_说明": "悬链线绘制参数 — 参考《电力工程设计手册 架空输电线路设计》",
    "K": 8e-5,
    "P1": [0.0, 0.0, 100.0],
    "P2": [300.0, 0.0, 100.0],
    "layer": "悬链线",
    "color": 1,
    "lineweight": 30,
    "n_pts": 200
}


def write_sample_param_file(path: str):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(SAMPLE_PARAM_JSON, f, ensure_ascii=False, indent=2)
    print(f"已生成示例参数文件: {path}")


# ─── CLI ────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="GstarCAD 三维悬链线绘制工具"
    )
    parser.add_argument("K", nargs="?", type=float, default=None,
                        help="弧垂 K 值 (1/m), 典型 4e-5 ~ 15e-5")
    parser.add_argument("--p1", nargs=3, type=float,
                        default=[0, 0, 100],
                        help="左悬挂点 x y z (m)")
    parser.add_argument("--p2", nargs=3, type=float,
                        default=None,
                        help="右悬挂点 x y z (m)")
    parser.add_argument("--span", type=float, default=300.0,
                        help="水平档距 (m), --p2 未指定时自动推算")
    parser.add_argument("--dh", type=float, default=0.0,
                        help="高差 (m)")
    parser.add_argument("--layer", type=str, default="悬链线",
                        help="图层名")
    parser.add_argument("--color", type=int, default=1,
                        help="颜色 ACI 号")
    parser.add_argument("--lw", type=int, default=30,
                        help="线宽")
    parser.add_argument("--n-pts", type=int, default=200,
                        help="采样点数")
    parser.add_argument("--param-file", type=str, default=None,
                        help="从 JSON 文件读取参数")
    parser.add_argument("--gen-sample", type=str, default=None,
                        help="生成示例参数 JSON 文件路径")
    parser.add_argument("--no-cad", action="store_true",
                        help="仅计算不绘制（预览模式）")
    parser.add_argument("--json-out", type=str, default=None,
                        help="导出点集到 JSON 文件")

    args = parser.parse_args()

    # 生成示例
    if args.gen_sample:
        write_sample_param_file(args.gen_sample)
        return

    # 从文件读取参数
    if args.param_file:
        with open(args.param_file, 'r', encoding='utf-8') as f:
            params = json.load(f)
        K = params.get("K", args.K)
        P1 = tuple(params.get("P1", args.p1))
        P2 = tuple(params.get("P2", args.p2)) if "P2" in params else None
        layer = params.get("layer", args.layer)
        color = params.get("color", args.color)
        lw = params.get("lineweight", args.lw)
        n_pts = params.get("n_pts", args.n_pts)
        span = params.get("span", args.span)
        dh = params.get("dh", args.dh)
    else:
        K = args.K
        P1 = tuple(args.p1)
        P2 = tuple(args.p2) if args.p2 else None
        layer = args.layer
        color = args.color
        lw = args.lw
        n_pts = args.n_pts
        span = args.span
        dh = args.dh

    if K is None:
        parser.error("必须提供 K 值，或使用 --param-file / --gen-sample")

    # 确定端点
    if P2 is None:
        P1 = (float(P1[0]), float(P1[1]), float(P1[2]))
        P2 = (float(P1[0]) + span, float(P1[1]), float(P1[2]) + dh)

    # 计算悬链线（n_pts 透传——修复 Bug 2）
    print(f"📐 K = {K:.2e}, P1={P1}, P2={P2}, 采样点数={n_pts}")
    points_3d, sag_max, L, a, x0 = catenary_3d(P1, P2, K, n_pts)
    print_summary(P1, P2, K, sag_max, L, a, x0)

    # 导出 JSON
    if args.json_out:
        data = format_gstarcad_polyline(points_3d)
        with open(args.json_out, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"✅ 已导出点集: {args.json_out}")

    # 绘制到 CAD
    if not args.no_cad:
        print("\n🔄 准备发送到 CAD ...")
        success = draw_catenary_in_gstarcad(
            points_3d, layer=layer, color=color, lineweight=lw
        )
        if success:
            print("✅ 三维悬链线已成功绘制到 CAD！")
        else:
            print("⚠️  CAD 绘制未完成。已导出计算数据，可手动处理。")
    else:
        print("🔍 预览模式（未发送到 CAD）。移除 --no-cad 可启动绘图。")


if __name__ == "__main__":
    main()
