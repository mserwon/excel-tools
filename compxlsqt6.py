import pandas as pd
from PyQt6.QtWidgets import QApplication, QFileDialog, QVBoxLayout, QPushButton, QLabel, QWidget

class ExcelComparer(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Select two Excel files to compare and highlight differences")
        layout.addWidget(self.label)

        self.btn1 = QPushButton("Select First Excel File", self)
        self.btn1.clicked.connect(self.selectFile1)
        layout.addWidget(self.btn1)

        self.btn2 = QPushButton("Select Second Excel File", self)
        self.btn2.clicked.connect(self.selectFile2)
        layout.addWidget(self.btn2)

        self.compareBtn = QPushButton("Compare and Highlight Differences", self)
        self.compareBtn.clicked.connect(self.compareFiles)
        layout.addWidget(self.compareBtn)

        self.setLayout(layout)
        
        self.file1 = None
        self.file2 = None

    def selectFile1(self):
        options = QFileDialog.Option()
        fileName, _ = QFileDialog.getOpenFileName(self, "Select First Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.file1 = fileName
            self.label.setText(f"First file selected: {fileName}")

    def selectFile2(self):
        options = QFileDialog.Option()
        fileName, _ = QFileDialog.getOpenFileName(self, "Select Second Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.file2 = fileName
            self.label.setText(f"Second file selected: {fileName}")

    def compareFiles(self):
        if not self.file1 or not self.file2:
            self.label.setText("Please select both Excel files to compare.")
            return

        df1 = pd.read_excel(self.file1)
        df2 = pd.read_excel(self.file2)

        comparison_values = df1.values == df2.values
        df_diff = pd.DataFrame(df1.values, columns=df1.columns)

        for row in range(len(comparison_values)):
            for col in range(len(comparison_values[row])):
                if not comparison_values[row][col]:
                    df_diff.iloc[row, col] = f'{df1.iloc[row, col]} --> {df2.iloc[row, col]}'

        output_file = 'differences.xlsx'
        df_diff.to_excel(output_file, index=False)
        
        self.label.setText(f"The differences have been highlighted and saved to '{output_file}'.")

if __name__ == '__main__':
    app = QApplication([])
    ex = ExcelComparer()
    ex.show()
    app.exec()