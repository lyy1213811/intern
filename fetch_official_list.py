import importlib

if __name__ == "__main__":
    mod = importlib.import_module("2_fetch_official_list")
    mod.fetch_hku_official_list()
