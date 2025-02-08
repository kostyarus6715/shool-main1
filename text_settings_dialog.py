from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QLineEdit, QHBoxLayout, QPushButton, QFileDialog, 
    QColorDialog, QMessageBox, QTableWidget, QTableWidgetItem, QScrollArea, QWidget
)
from PyQt6.QtCore import Qt

class TextSettingsDialog(QDialog):
    def __init__(self, template_file_path, columns, font_file_path="TimesnewRomanBold.ttf", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройка текста")
        self.setGeometry(200, 200, 700, 500)

        self.template_file_path = template_file_path
        self.font_file_path = font_file_path if font_file_path else "TimesNewRomanBold.ttf"
        self.columns = columns  

        self.common_settings = {"x_position": 220, "y_position": 380}

        self.column_font_sizes = {column: 12 for column in columns}
        self.column_spacing = {column: 0 for column in columns}
        self.column_colors = {column: '#003380' for column in columns}  

        layout = QVBoxLayout()

        layout.addWidget(QLabel("Общие настройки текста:"))

        x_layout = QHBoxLayout()
        x_layout.addWidget(QLabel("Начальная позиция X:"))
        self.x_input = QLineEdit(str(self.common_settings["x_position"]))
        x_layout.addWidget(self.x_input)
        layout.addLayout(x_layout)

        y_layout = QHBoxLayout()
        y_layout.addWidget(QLabel("Начальная позиция Y:"))
        self.y_input = QLineEdit(str(self.common_settings["y_position"]))
        y_layout.addWidget(self.y_input)
        layout.addLayout(y_layout)

        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("Файл шрифта:"))
        self.font_input = QLineEdit(self.font_file_path)
        self.font_button = QPushButton("Выбрать файл")
        self.font_button.clicked.connect(self.select_font_file)
        font_layout.addWidget(self.font_input)
        font_layout.addWidget(self.font_button)
        layout.addLayout(font_layout)

        layout.addWidget(QLabel("Настройки для каждой колонки:"))

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)  

        self.column_widget = QWidget()
        self.column_layout = QVBoxLayout()
        self.column_widget.setLayout(self.column_layout)

        self.scroll_area.setWidget(self.column_widget)
        layout.addWidget(self.scroll_area)

        self.table = QTableWidget(len(self.columns), 4)  
        self.table.setHorizontalHeaderLabels(["Колонка", "Размер шрифта", "Отступ", "Цвет"])
        self.table.setColumnWidth(0, 200)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(3, 100)

        for row, column in enumerate(self.columns):
            self.table.setItem(row, 0, QTableWidgetItem(column))

            font_size_item = QTableWidgetItem(str(self.column_font_sizes[column]))
            font_size_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, 1, font_size_item)

            spacing_item = QTableWidgetItem(str(self.column_spacing[column]))
            spacing_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, 2, spacing_item)

            color_button = QPushButton("Выбрать цвет")
            color_button.clicked.connect(lambda checked, r=row: self.select_color(r))
            self.table.setCellWidget(row, 3, color_button)

        self.column_layout.addWidget(self.table)

        self.save_button = QPushButton("Сохранить настройки")
        self.save_button.clicked.connect(self.save_settings)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

    def save_settings(self):
        """Сохраняет текущие настройки."""
        try:

            self.common_settings["x_position"] = int(self.x_input.text())
            self.common_settings["y_position"] = int(self.y_input.text())

            if not self.font_input.text():
                QMessageBox.warning(self, "Ошибка", "Выберите файл шрифта.")
                return
            self.font_file_path = self.font_input.text()

            for row in range(self.table.rowCount()):
                column_name = self.table.item(row, 0).text()

                font_size = int(self.table.item(row, 1).text())
                self.column_font_sizes[column_name] = font_size

                spacing_text = self.table.item(row, 2).text()
                spacing = int(spacing_text) if spacing_text.strip().isdigit() else 0
                self.column_spacing[column_name] = spacing

                color_str = self.table.cellWidget(row, 3).styleSheet().split(":")[-1].strip()
                if color_str.startswith('#') and len(color_str) == 7:
                    self.column_colors[column_name] = color_str
                else:

                    self.column_colors[column_name] = '#003380'  

            QMessageBox.information(self, "Успех", "Настройки успешно сохранены.")
            self.accept()  
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите корректные числовые значения.")

    def select_font_file(self):
        """Открывает диалог выбора файла шрифта."""
        font_file, _ = QFileDialog.getOpenFileName(self, "Выберите файл шрифта", "", "Шрифты (*.ttf);;Все файлы (*)")
        if font_file:
            self.font_file_path = font_file
            self.font_input.setText(font_file)

    def select_color(self, row):
        """Открывает диалог выбора цвета шрифта для колонки."""
        color = QColorDialog.getColor(self.column_colors[self.columns[row]], self)
        if color.isValid():
            button = self.table.cellWidget(row, 3)
            button.setStyleSheet(f"background-color: {color.name()}")
            self.column_colors[self.columns[row]] = color

    def get_settings(self):
        """Возвращает все настройки."""
        return {
            "common_settings": self.common_settings,
            "column_font_sizes": self.column_font_sizes,
            "column_spacing": self.column_spacing,
            "column_colors": self.column_colors,  
            "font_file_path": self.font_file_path,
        }