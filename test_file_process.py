'''
pytest -s test_file_process.py
'''
import os
from file_process import files_list

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
