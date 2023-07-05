import os
from constants import *
"""測試路徑是否存在"""
#確認相關路徑是否存在
def test_path():
    if os.path.exists(AZURE_SOURCE):
        print("1. The folder exists!", AZURE_SOURCE)
    else:
        print("1. The folder does not exist!", AZURE_SOURCE)
    if os.path.exists(AZURE_DESTINY):
        print("2. The folder exists!", AZURE_DESTINY)
    else:
        print("2. The folder does not exist!", AZURE_DESTINY)
    #確認套印樣本是否存在
    if os.path.exists(HT_PRINT_STD_FILE10):
        print("3. The file exists!", HT_PRINT_STD_FILE10)
    else:
        print("3. The file does not exist!", HT_PRINT_STD_FILE10)

    if os.path.exists(TC_PRINT_STD_FILE20):
        print("4. The file exists!", TC_PRINT_STD_FILE20)
    else:
        print("4. The file does not exist!", TC_PRINT_STD_FILE20)
    #確認傳真樣本是否存在
    if os.path.exists(HT_FAX_FILE10):
        print("5. The file exists!", HT_FAX_FILE10)
    else:
        print("5. The file does not exist!", HT_FAX_FILE10)
    if os.path.exists(TC_FAX_FILE20):
        print("6. The file exists!", TC_FAX_FILE20)
    else:
        print("6. The file does not exist!", TC_FAX_FILE20)
    #確認審查意見是否存在
    if os.path.exists(HT_COMMENT_FILE10):
        print("3. The file exists!", HT_COMMENT_FILE10)
    else:
        print("3. The file does not exist!", HT_COMMENT_FILE10)
    if os.path.exists(TC_COMMENT_FILE20):
        print("3. The file exists!", TC_COMMENT_FILE20)
    else:
        print("3. The file does not exist!", TC_COMMENT_FILE20)

def test_():
    print(type(DOC_NO_STRUCTURE))

if __name__ == '__main__':
    test_path()
