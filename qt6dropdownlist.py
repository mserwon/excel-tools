import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QFileDialog, QPushButton, QLabel, QListWidget, QAbstractItemView

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PyQt6 File Selector and Column Reader")

        # Create a button to select the file
        self.select_file_button = QPushButton("Select Excel/CSV File")
        self.select_file_button.clicked.connect(self.select_file)

        # Create a label for the multi-select combo box
        self.combo_box_label = QLabel("Select Common Column Headers")

        # Create a multi-select list widget
        self.multi_select_combo_box = QListWidget()
        self.multi_select_combo_box.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)

        # Create a button to print selected items
        self.print_button = QPushButton("Print Selected Columns")
        self.print_button.clicked.connect(self.print_selected_columns)

        # Set layout
        layout = QVBoxLayout()
        layout.addWidget(self.select_file_button)
        layout.addWidget(self.combo_box_label)
        layout.addWidget(self.multi_select_combo_box)
        layout.addWidget(self.print_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file(self):
        # Open a file dialog to select a file
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Open File", "", "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)")

        if file_path:
            self.load_columns(file_path)

    def load_columns(self, file_path):
        try:
            # Read the file to extract column headers
            if file_path.endswith('.csv'):
                data = pd.read_csv(file_path)
            else:
                data = pd.read_excel(file_path)

            # Get the column headers
            column_headers = list(data.columns)

            # Populate the multi-select list widget
            self.multi_select_combo_box.clear()
            self.multi_select_combo_box.addItems(column_headers)

        except Exception as e:
            print(f"Error reading file: {e}")

    def print_selected_columns(self):
        # Get selected items
        selected_items = [item.text() for item in self.multi_select_combo_box.selectedItems()]
        print("Selected Columns:", selected_items)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
