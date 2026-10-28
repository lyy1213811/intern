import sys
import argparse
import importlib

# 真正的实现代码位于: 2_fetch_official_list.py
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="官网专业列表抓取入口")
    parser.add_argument("--school", default="cuhk", help="院校代码，如 cuhk, hku")
    args = parser.parse_args()

    mod = importlib.import_module("2_fetch_official_list")
    if args.school.lower() == "cuhk":
        mod.fetch_cuhk_official_list("cuhk")
    else:
        mod.fetch_hku_official_list(args.school)
