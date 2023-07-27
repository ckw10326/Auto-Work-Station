from .copy_file import copy_the_excel_file2

def test_path():
    file_source_folder = "/workspaces/Auto-Work-Station/00source"
    file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
    print(file_source_folder)
    print(file_dest_folder)
    if os.path.exists(file_source_folder):
        print("路徑file_source_folder確實存在")
    else:
        print("不存在")
    if os.path.exists(file_dest_folder):
        print("路徑file_dest_folder確實存在")
    else:
        print("不存在")