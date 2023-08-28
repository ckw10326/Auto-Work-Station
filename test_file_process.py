'''
pytest -s test_file_process.py
'''
import os
from file_process import files_list, check_plan

def test_files_list():
    # 測試情境1：沒有指定關鍵字
    xpath = "/path/to/directory"
    expected_output = [
        "/path/to/directory/file1.txt",
        "/path/to/directory/subdirectory/file2.txt"
    ]
    result = files_list(xpath)
    assert result.sort() == expected_output.sort()

    # 測試情境2：指定關鍵字
    xpath = "/path/to/directory"
    search_str = "file2"
    expected_output = [
        "/path/to/directory/subdirectory/file2.txt"
    ]
    result = files_list(xpath, search_str)
    assert result.sort() == expected_output.sort()

def test_check_plan_with_matching_file():
    # 建立包含符合條件的檔案的測試資料夾結構
    test_folder = "test_folder"
    os.makedirs(test_folder, exist_ok=True)
    matching_file = os.path.join(test_folder, "file-CTC-.txt")
    with open(matching_file, "w") as file:
        file.write("Test file with -CTC-")

    # 執行測試
    assert check_plan(test_folder) == True

    # 清理測試資料夾
    os.remove(matching_file)
    os.rmdir(test_folder)

def test_check_plan_without_matching_file():
    # 建立不包含符合條件的檔案的測試資料夾結構
    test_folder = "test_folder"
    os.makedirs(test_folder, exist_ok=True)
    non_matching_file = os.path.join(test_folder, "file.txt")
    with open(non_matching_file, "w") as file:
        file.write("Test file without -CTC-")

    # 執行測試
    assert check_plan(test_folder) == False

    # 清理測試資料夾
    os.remove(non_matching_file)
    os.rmdir(test_folder)