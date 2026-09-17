"""风偏弧垂曲线生成器（独立程序）。

功能
----
在 AutoCAD / GstarCAD / 中望CAD 中选取一条弧垂曲线（AcDb3dPolyline / Line /
Polyline / Spline 等），输入一个角度（度），基于原曲线生成一条
「风偏弧垂曲线」：
  * 原曲线保留不动；
  * 新曲线 = 原曲线绕「两端端点连线」为轴旋转给定角度；
  * 新曲线置于图层 "风偏弧垂曲线"，蓝色显示（ACI 5）。

旋转方向约定（全项目统一，与主程序 catenary_app.py 一致）
--------------------------------------------------------
**右偏为正(+)、左偏为负(−)**；"左右"以曲线起点→终点（点选先后顺序）
为前进方向——面向 P1→P2 观察，右手边即为"右"。
（历史：早期独立版用 ANGLE_SIGN 常量实现"正=顺时针"，与主程序
"正=右偏"等价——顺时针 = 面向 P1→P2 的右手方向。重构后统一用
catenary_core.rotate_curve 的右偏为正约定，ANGLE_SIGN 已废弃。）

用法
----
  python catenary_winddev.py                 # 默认：纯数学自检（不需要 CAD）
  python catenary_winddev.py selftest        # 同上
  python catenary_winddev.py cad --pick                 # 交互：点选曲线 + 命令行输入角度
  python catenary_winddev.py cad --handle ABCD --angle 30.6   # 指定句柄 + 角度

代码结构（重构后，消除重复——审查 Opt 2）:
  - catenary_core.py —— 旋转/悬链线纯算法（rotate_curve / rotate_point）
  - catenary_cad.py  —— win32com 封装（get_gcad / pick_entity /
                        extract_points / draw_wind_curve / draw_wind_label）
  - catenary_winddev.py —— 本文件：独立 CLI + selftest

依赖：pywin32（win32com / pythoncom）。仅 cad 子命令需要 CAD 已启动。
selftest 纯数学自检**不依赖 pywin32**（catenary_cad 惰性导入，审查 F1）。
"""

from __future__ import annotations

import argparse
import math

from catenary_core import rotate_curve


# ---------------------------------------------------------------------------
# 纯数学自检（不需要 CAD）
# ---------------------------------------------------------------------------
def _selftest():
    """验证旋转数学：端点不动、+90° 右偏、-90° 左偏。"""
    # 水平档距沿 X 轴，弧垂向下（-Z）
    P1 = (0.0, 0.0, 0.0)
    P2 = (100.0, 0.0, 0.0)
    M = (50.0, 0.0, -10.0)  # 跨中最低点
    curve = [P1, M, P2]

    def approx(a, b, tol=1e-9):
        return all(abs(x - y) <= tol for x, y in zip(a, b))

    # +90°（右偏：面向 P1→P2，右手边 = -Y；最低点由 -Z 摆到 -Y）
    rot_p = rotate_curve(curve, 90.0)
    assert approx(rot_p[0], P1), f"起点应不变, got {rot_p[0]}"
    assert approx(rot_p[2], P2), f"终点应不变, got {rot_p[2]}"
    assert approx(rot_p[1], (50.0, -10.0, 0.0)), f"+90° 中点应为(50,-10,0), got {rot_p[1]}"
    print("  [+] +90°: 端点不变, 中点 (50,0,-10)->(50,-10,0) OK（右偏）")

    # -90°（左偏：最低点由 -Z 摆到 +Y）
    rot_n = rotate_curve(curve, -90.0)
    assert approx(rot_n[1], (50.0, 10.0, 0.0)), f"-90° 中点应为(50,10,0), got {rot_n[1]}"
    print("  [+] -90°: 中点 (50,0,-10)->(50,10,0) OK（左偏）")

    # 30.6° 合理性：跨中点到轴的距离应等于原垂距 10，仅方向转过 30.6°
    rot_a = rotate_curve(curve, 30.6)
    assert approx(rot_a[0], P1) and approx(rot_a[2], P2), "30.6° 端点应不变"
    # 跨中点应仍在垂直于轴的平面内，且离轴距离 = 10
    mx, my, mz = rot_a[1]
    # 轴沿 X，故离轴距离 = sqrt(my^2+mz^2)
    r = math.hypot(my, mz)
    assert abs(r - 10.0) < 1e-6, f"离轴距离应=10, got {r}"
    # 与 -Z 方向夹角应≈30.6°（右偏，落在 -Y 侧）
    ang = math.degrees(math.atan2(my, -mz))  # my<0, mz≈0 -> 负角
    assert abs(ang - (-30.6)) < 1e-6, f"偏转角应≈-30.6°, got {ang}"
    print(f"  [+] 30.6°: 离轴距离={r:.4f}, 偏转角={ang:.4f}° OK")

    print("SELFTEST PASSED ✅")


# ---------------------------------------------------------------------------
# CAD 交互流程
# ---------------------------------------------------------------------------
def cad_flow(angle: float | None, handle: str | None, cad: str | None = None):
    """连接 CAD，拾取/定位曲线，旋转并绘制风偏弧垂曲线。

    ⚠️ catenary_cad 惰性导入（审查 F1）：仅在此处需要，避免 selftest
    纯数学自检被迫依赖 pywin32。
    """
    from catenary_cad import (
        get_gcad,
        pick_entity,
        pick_angle,
        extract_points,
        get_entity_by_handle,
        draw_wind_curve,
        draw_wind_label,
        LAYER_WIND,
    )
    gc, doc, ms, app = get_gcad(cad)
    print(f"→ 已连接 {app}，当前图: {doc.Name}")
    if handle:
        ent = get_entity_by_handle(ms, handle)
        print(f"→ 已按句柄 {handle} 定位曲线")
    else:
        print("→ 请在 CAD 绘图区点选一条弧垂曲线（CAD 命令行也会显示提示）...")
        ent = pick_entity(doc, "\n请选择一条弧垂曲线: ")
        print(f"→ 已拾取曲线 (句柄 {ent.Handle})")

    if angle is None:
        print("→ 请在 CAD 命令行输入风偏角(度)，回车确认（右偏为正）...")
        angle = pick_angle(doc, f"\n请输入风偏角(度, 右偏为正): ")
    print(f"→ 风偏角 = {angle}°")

    pts = extract_points(ent)
    if len(pts) < 2:
        raise ValueError("所选曲线点列不足 2 点，无法旋转")
    p1, p2 = pts[0], pts[-1]
    print(f"→ 旋转轴端点 P1={p1}  P2={p2}")

    new_pts = rotate_curve(pts, angle)
    poly = draw_wind_curve(doc, ms, new_pts)
    label, label_text = draw_wind_label(doc, ms, new_pts, angle)
    print(f"→ 已生成「风偏弧垂曲线」: 句柄 {poly.Handle}, 蓝色, 图层'{LAYER_WIND}'")
    print(f"→ 文字标注: '{label_text}' (句柄 {label.Handle}, 按绝对值标示, 忽略正负号)")
    print("→ 原曲线保持不变。")
    return poly


def main(argv=None):
    # cad_family_keys() 只读配置表（catenary_cad 顶层无 pywin32 依赖），
    # 在此处导入可让 selftest 分支完全不触碰 CAD 相关代码。
    from catenary_cad import cad_family_keys

    parser = argparse.ArgumentParser(description="风偏弧垂曲线生成器（独立程序）")
    sub = parser.add_subparsers(dest="cmd")
    sub.add_parser("selftest", help="纯数学自检（不需要 CAD）")
    p_cad = sub.add_parser("cad", help="在 CAD 中交互生成")
    p_cad.add_argument("--pick", action="store_true", help="交互点选曲线（默认）")
    p_cad.add_argument("--handle", type=str, default=None, help="按句柄指定曲线")
    p_cad.add_argument("--angle", type=float, default=None, help="风偏角(度)，省略则在 CAD 命令行输入")
    p_cad.add_argument("--cad", choices=cad_family_keys(), default=None,
                       help="强制指定 CAD (默认自动检测: 优先已运行的实例)")
    args = parser.parse_args(argv)

    if args.cmd == "cad":
        cad_flow(args.angle, args.handle, args.cad)
    else:
        # 默认与 selftest 等价
        print("运行纯数学自检（无需 CAD）...")
        _selftest()


if __name__ == "__main__":
    main()
