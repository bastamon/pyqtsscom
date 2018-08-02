from PyInstaller.__main__ import run
if __name__ == '__main__':
    opts = ['sscom.py', '-w', '--icon=my.ico']
    run(opts)
