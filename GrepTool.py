"""Grep ツールの起動用スクリプト (python GrepTool.py [フォルダ])"""
import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support()
    from greptool.app_main import main
    main()
