import importlib

if __name__ == "__main__":
    mod = importlib.import_module("1_export_db_slice")
    mod.export_db_slice()
