import sys
from PyQt5 import QtWidgets
from layout import Ui_Form



class main(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(main, self).__init__(parent)
        self.ui = Ui_Form()
        self.ui.setupUi(self)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = main()
    window.show()
    sys.exit(app.exec_())