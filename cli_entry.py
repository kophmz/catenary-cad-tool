#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
空间弧垂曲线工具 CLI 入口（console 打包专用）
============================================================

强制走命令行模式（cli_main），**不启动 GUI**：
  - 带子命令（draw / mindist / wind / bundle / suspension）→ 正常执行
  - 无参数 → 打印用法帮助后退出

用途：打包成「空间弧垂曲线工具CLI.exe」（PyInstaller --console），
供命令行/脚本调用，有控制台输出、跑完正常退出。

示例:
  空间弧垂曲线工具CLI.exe draw --K 15.22 --p1 0 0 0 --p2 100 30 12
  空间弧垂曲线工具CLI.exe wind --handle E90 --angle 30.6
  空间弧垂曲线工具CLI.exe suspension --K 15.22 --p1 0 0 0 --p2 100 30 12 --s1 --d1 3
"""

import catenary_app

if __name__ == "__main__":
    catenary_app.cli_main()
