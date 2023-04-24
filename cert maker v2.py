import importlib
import subprocess

# check if psd-tools is installed
try:
    importlib.import_module('psd_tools')
except ImportError:
    # install psd-tools
    subprocess.check_call(['pip', 'install', 'psd-tools'])

# check if PyQt5 is installed
try:
    importlib.import_module('PyQt5')
except ImportError:
    # install PyQt5
    subprocess.check_call(['pip', 'install', 'PyQt5'])

from psd_tools import PSDImage
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QComboBox, QLabel, QPushButton


class PSDWriter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # create UI elements
        self.setGeometry(300, 300, 350, 250)
        self.setWindowTitle('PSD Writer')

        self.file_label = QLabel(self)
        self.file_label.move(10, 10)
        self.file_label.resize(330, 20)

        self.layer_label = QLabel('Select Layer:', self)
        self.layer_label.move(10, 40)

        self.layer_box = QComboBox(self)
        self.layer_box.move(100, 40)
        self.layer_box.resize(240, 20)

        self.write_button = QPushButton('Write Text', self)
        self.write_button.move(10, 80)
        self.write_button.resize(330, 30)
        self.write_button.clicked.connect(self.write_text)

        self.show()

    def browse_file(self):
        # open file dialog to choose PSD file
        filename, _ = QFileDialog.getOpenFileName(self, "Open PSD file", "", "PSD Files (*.psd)")
        if filename:
            self.file_label.setText(filename)
            # get layer names from PSD file
            img = PSDImage.open(filename)
            self.layer_names = [layer.name for layer in img.layers if not layer.is_group()]
            self.layer_box.addItems(self.layer_names)

    def write_text(self):
        # get selected layer name and user input text
        layer_name = self.layer_box.currentText()
        text, ok = QInputDialog.getText(self, 'Write Text', 'Enter text:')

        if ok:
            # open PSD file and write text on selected layer
            img = PSDImage.open(self.file_label.text())
            for layer in img.layers:
                if layer.name == layer_name:
                    layer.text = text
                    break
            # save file as JPG
            img.save(self.file_label.text().replace('.psd', '.jpg'))
            # show success message
            QMessageBox.information(self, 'PSD Writer', 'Text written successfully!')

if __name__ == '__main__':
    # create application
    app = QApplication([])
    # create main window
    window = PSDWriter()
    # run application
    app.exec_()
