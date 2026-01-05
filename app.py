from PySide6.QtWidgets import QApplication
import sys
from dashboard import Dashboard

app = QApplication(sys.argv)
janela = Dashboard()
janela.show()
sys.exit(app.exec())

