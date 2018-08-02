if __name__ == '__main__':
    from PyInstaller.__main__ import run
    opts = ['sscom.py', '-w', '--icon=my.ico']
    run(opts)
