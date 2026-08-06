#!/usr/bin/env python3
"""
三维悬链线生成脚本（纯算法 CLI）
基于《电力工程设计手册 架空输电线路设计》中弧垂模板 K 值公式，生成 1:1（米）悬链线点集。

核心公式（手册式14-1）：
    f_c = σ_c/γ_c × (ch(γ_c·l/(2σ_c)) - 1)
    K = γ_c/(8σ_c) = P_c/(8T_c)

  代入得：
    y(x) = 1/(8K) × (cosh(4K·x) - 1)

  其中：
    - x 为距最低点的水平距离（m）
    - y 为弧垂（m，向下为正）
    - K 为弧垂模板 K 值（1/m），钢芯铝绞线典型范围 4×10⁻⁵ ~ 15×10⁻⁵

比例尺说明（手册式14-2）：
    纵 1:500、横 1:5000 的模板 K 值即为物理 K 值 K = γ/(8σ)。
    模板只是比例绘图工具，K 值本身已是物理量，无需额外换算。

三维支持：
    支持两端不等高悬挂，悬链线在竖直平面内弯曲。
    通过两悬挂点的 3D 坐标生成沿空间直线的悬链线点集。

代码说明：
    算法实现统一在 catenary_core.py（纯算法模块），本文件仅保留 CLI 与
    输出辅助，避免与主程序/风偏程序重复漂移（审查 Opt 2）。
"""

import argparse
import json

from catenary_core import (
    catenary_3d,
    solve_catenary_2d,
    format_gstarcad_polyline,
    format_xyz_csv,
    print_summary,
)


def generate_sag_curve_points(K: float, span: float, delta_h: float = 0.0,
                              n_pts: int = 200):
    """
    纯二维悬链线点集（用于导出/CAD绘制/可视化），
    以左悬挂点为原点 (0, 0)，右悬挂点为 (span, delta_h)。

    返回:
        (x_vals, y_vals) 两个列表
    """
    a, x0, y0, x_vals, y_vals = solve_catenary_2d(span, delta_h, K, n_pts)
    return x_vals, y_vals


# ─── CLI ────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="三维悬链线生成器 — 基于线路手册 K 值公式"
    )
    parser.add_argument("K", type=float,
                        help="弧垂模板 K 值 (1/m)，典型 4e-5 ~ 15e-5")
    parser.add_argument("-s", "--span", type=float, default=200.0,
                        help="水平档距 (m)，默认 200")
    parser.add_argument("--dh", type=float, default=0.0,
                        help="悬挂点高差 (m)，默认 0（等高）")
    parser.add_argument("--n-pts", type=int, default=200,
                        help="采样点数 (默认 200)")
    parser.add_argument("--p1", nargs=3, type=float,
                        default=[0, 0, 100],
                        help="左悬挂点三维坐标 x y z (m)，默认 0 0 100")
    parser.add_argument("--p2", nargs=3, type=float,
                        default=None,
                        help="右悬挂点三维坐标 x y z (m)，"
                             "默认自动根据 --span 推算")
    parser.add_argument("--csv", type=str, default=None,
                        help="导出 XYZ 到 CSV 文件路径")
    parser.add_argument("--cad-json", type=str, default=None,
                        help="导出 GstarCAD draw_polyline JSON 文件路径")

    args = parser.parse_args()

    # 处理端点
    P1 = tuple(args.p1)
    if args.p2:
        P2 = tuple(args.p2)
    else:
        # 默认右端点：沿 X 轴方向，档距为 --span
        P1 = (0.0, 0.0, P1[2])
        P2 = (args.span, 0.0, P1[2] + args.dh)

    # 计算悬链线（n_pts 全链路透传——修复 Bug 2）
    points_3d, sag_max, L, a, x0 = catenary_3d(P1, P2, args.K, args.n_pts)

    # 打印摘要
    print_summary(P1, P2, args.K, sag_max, L, a, x0)

    # 导出
    if args.csv:
        format_xyz_csv(points_3d, args.csv)

    if args.cad_json:
        data = format_gstarcad_polyline(points_3d)
        with open(args.cad_json, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"已导出 GstarCAD JSON: {args.cad_json}")

    # CLI 模式：默认打印前10个和最后10个点
    print("\n  悬链线点集 (前5 + 后5):")
    for p in points_3d[:5]:
        print(f"    ({p[0]:.3f}, {p[1]:.3f}, {p[2]:.3f})")
    print(f"    ... 共 {len(points_3d)} 点 ...")
    for p in points_3d[-5:]:
        print(f"    ({p[0]:.3f}, {p[1]:.3f}, {p[2]:.3f})")


if __name__ == "__main__":
    main()
