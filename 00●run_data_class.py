import os
import shutil
import sys
import file_process

def main():
    dest_folder = "/workspaces/Auto-Work-Station/00dest"
    if os.path.exists(dest_folder):
        shutil.rmtree(dest_folder)
    file_process.work_flow()

if __name__ == '__main__':
    main()