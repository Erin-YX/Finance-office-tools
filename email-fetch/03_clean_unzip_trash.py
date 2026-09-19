```python
# 获取配置文件
# Load the configuration file

def load_config() -> dict:
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "clean_config.json")
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


# 清理功能：
# 将解压产生的子文件夹中的文件移动到目标文件夹，
# 并删除目标文件夹内所有空的子文件夹和 ZIP 压缩包
#
# Cleaning function:
# Move files from extracted subfolders to the target folder,
# then delete all empty subfolders and ZIP files inside the target folder.

def clean_target_folder(target: str) -> None:
    tmp_idx = 0

    # ==================================================
    # 1. 为防止子文件夹与文件同名造成移动冲突，
    #    先将每个子文件夹重命名为临时名称，
    #    再将其中的内容移动到 target 文件夹
    #
    # 1. To avoid conflicts between files and subfolders with the same name,
    #    temporarily rename each subfolder first,
    #    then move its contents into the target folder.
    # ==================================================

    subfolders = [
        name for name in os.listdir(target)
        if os.path.isdir(os.path.join(target, name))
    ]

    for name in subfolders:
        folder_path = os.path.join(target, name)
        if not os.path.isdir(folder_path):
            continue

        tmp_idx += 1
        tmp_name = f"__tmpmove_{tmp_idx}__"
        tmp_path = os.path.join(target, tmp_name)

        # 关键步骤：先重命名子文件夹，释放原文件夹名称，
        # 避免移动内容时与 target 中已有的同名文件发生冲突
        #
        # Key step: rename the subfolder first to free up its original name
        # and avoid conflicts with files in the target folder.
        os.rename(folder_path, tmp_path)

        print(f"Moving contents from: {folder_path}")

        for item in os.listdir(tmp_path):
            src_item = os.path.join(tmp_path, item)
            dst_item = os.path.join(target, item)
            shutil.move(src_item, dst_item)


    # ==================================================
    # 2. 删除空的子文件夹
    # 2. Delete empty subfolders
    # ==================================================

    for name in os.listdir(target):
        folder_path = os.path.join(target, name)
        if os.path.isdir(folder_path):
            try:
                os.rmdir(folder_path)
            except OSError:
                pass


    # ==================================================
    # 3. 删除 ZIP 压缩包
    # 3. Delete ZIP files
    # ==================================================

    for name in os.listdir(target):
        if name.lower().endswith(".zip"):
            file_path = os.path.join(target, name)
            try:
                os.remove(file_path)
            except OSError:
                pass


# 单任务模式：目标文件夹中包含 ZIP 压缩包及其解压后产生的子文件夹
# Single mode: the target folder contains ZIP files and their extracted subfolders.

def run_single_mode(target: str) -> None:
    # target = input("Please enter the target folder path: ").strip()

    if target.endswith("\\"):
        target = target[:-1]

    if not os.path.isdir(target):
        print("\nERROR: Folder does not exist.\n")
        # input("Press Enter to exit...")
        sys.exit(1)

    print("\nTarget folder:")
    print(target)
    print()

    clean_target_folder(target)

    print("\n========================================")
    print("Cleaning completed.")
    print("========================================")


# 批量模式：将所有需要清理的目标文件夹放在同一个父文件夹下
# Batch mode: place all target folders under the same parent folder.

def run_batch_mode(parent: str) -> None:
    # parent = input("Please enter the parent folder path: ").strip()

    if parent.endswith("\\"):
        parent = parent[:-1]

    if not os.path.isdir(parent):
        print("\nERROR: Parent folder does not exist.\n")
        # input("Press Enter to exit...")
        sys.exit(1)

    print("\nParent folder:")
    print(parent)
    print()

    subfolders = [
        os.path.join(parent, name)
        for name in os.listdir(parent)
        if os.path.isdir(os.path.join(parent, name))
    ]

    for target in subfolders:
        print("\n========================================")
        print(f"Processing: {target}")
        print("========================================")

        clean_target_folder(target)

        print(f"Cleaning completed for: {target}")

    print("\n========================================")
    print("Batch cleaning completed.")
    print("========================================")


def main() -> None:
    config = load_config()
    mode = config.get("mode")

    if mode == "single":
        run_single_mode(config["target"])

    elif mode == "batch":
        run_batch_mode(config["parent"])

    else:
        print(f"\nERROR: Invalid mode in config: {mode!r} (expected 'single' or 'batch')\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
```
