import os
from copy_file import copy_the_excel_file2

source = "/workspaces/Auto-Work-Station/00source"
dest = "/workspaces/Auto-Work-Station/00dest"
if not os.path.exists(dest):
    os.makedirs(dest)
copy_the_excel_file2(source, dest)