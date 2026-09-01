from pathlib import Path

folder = Path(".")

for pattern in ["*.xlsx", "*.pdf"]:
    for file in folder.glob(pattern):
        try:
            file.unlink()
            print(f"已删除: {file}")
        except Exception as e:
            print(f"失败: {file} -> {e}")

print("完成")