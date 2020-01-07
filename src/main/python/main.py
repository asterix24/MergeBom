from fbs_runtime.application_context.PyQt5 import ApplicationContext
from lib.qtgui import MergeBomGUI
import sys

if __name__ == '__main__':
    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext
    window = MergeBomGUI()
    window.show()
    exit_code = appctxt.app.exec_()      # 2. Invoke appctxt.app.exec_()
    sys.exit(exit_code)
