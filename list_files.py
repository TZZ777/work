import json
import os
import sys

def list_files(folder_path):
    abs_root = os.path.abspath(folder_path)
    file_paths = []
    for root, _, files in os.walk(abs_root):
        for name in files:
            file_paths.append(os.path.join(root, name))
    return file_paths

def main():
    if len(sys.argv) < 2:
        print("需要传入文件夹路径", file=sys.stderr)
        sys.exit(1)
    folder_path = sys.argv[1]
    if not os.path.isdir(folder_path):
        print("路径不是有效文件夹", file=sys.stderr)
        sys.exit(1)
    result = list_files(folder_path)
    print(json.dumps(result, ensure_ascii=False))

if __name__ == "__main__":
    main()
