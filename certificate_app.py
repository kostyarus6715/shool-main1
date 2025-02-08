from PyQt6.QtWidgets import QMainWindow, QVBoxLayout, QLabel, QPushButton, QLineEdit, QFileDialog, QListWidget, QWidget, QMessageBox, QDialog, QHBoxLayout, QCheckBox, QScrollArea
from PyQt6.QtCore import Qt
import openpyxl
from text_settings_dialog import TextSettingsDialog
from utils import load_excel_data, create_certificate
import os

class CertificateApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Настройка сертификатов")
        self.setGeometry(100, 100, 600, 700)

        layout = QVBoxLayout()

        layout.addWidget(QLabel("Настройка файлов:"))

        self.excel_input = QLineEdit()
        self.excel_button = QPushButton("Обзор")
        self.excel_button.clicked.connect(self.select_excel_file)

        layout.addWidget(QLabel("Excel файл:"))
        layout.addWidget(self.excel_input)
        layout.addWidget(self.excel_button)

        self.template_input = QLineEdit()
        self.template_button = QPushButton("Обзор")
        self.template_button.clicked.connect(self.select_template_file)

        layout.addWidget(QLabel("PDF шаблон:"))
        layout.addWidget(self.template_input)
        layout.addWidget(self.template_button)

        self.output_input = QLineEdit()
        self.output_button = QPushButton("Обзор")
        self.output_button.clicked.connect(self.select_output_folder)

        layout.addWidget(QLabel("Папка для сохранения сертификатов:"))
        layout.addWidget(self.output_input)
        layout.addWidget(self.output_button)

        layout.addWidget(QLabel("Выберите колонки и заголовки:"))

        self.columns_header_layout = QHBoxLayout()
        self.columns_header_layout.addWidget(QLabel("Колонка"))
        self.columns_header_layout.addWidget(QLabel("Отобразить данные"))
        self.columns_header_layout.addWidget(QLabel("Отобразить заголовок"))
        layout.addLayout(self.columns_header_layout)

        self.scroll_area = QScrollArea()
        self.columns_list = QWidget()
        self.columns_layout = QVBoxLayout()
        self.columns_list.setLayout(self.columns_layout)
        self.scroll_area.setWidget(self.columns_list)  
        self.scroll_area.setWidgetResizable(True)  
        layout.addWidget(self.scroll_area)

        self.text_settings_button = QPushButton("Настройка текста")
        self.text_settings_button.clicked.connect(self.open_text_settings)
        layout.addWidget(self.text_settings_button)

        self.generate_button = QPushButton("Сгенерировать сертификаты")
        self.generate_button.clicked.connect(self.generate_certificates)
        layout.addWidget(self.generate_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.excel_file_path = ""
        self.template_file_path = ""
        self.output_folder_path = ""
        self.font_file_path = ""
        self.common_settings = {}
        self.column_font_sizes = {}
        self.column_spacing = {}
        self.column_colors = {}

        self.checkboxes = {}  

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel файл", "", "Excel Files (*.xlsx)")
        if file_path:
            self.excel_file_path = file_path
            self.excel_input.setText(file_path)
            self.load_excel_columns()

    def select_template_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите PDF шаблон", "", "PDF Files (*.pdf)")
        if file_path:
            self.template_file_path = file_path
            self.template_input.setText(file_path)

    def select_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        if folder_path:
            self.output_folder_path = folder_path
            self.output_input.setText(folder_path)

    def load_excel_columns(self):
        if not self.excel_file_path:
            return

        try:
            workbook = openpyxl.load_workbook(self.excel_file_path)
            sheet = workbook.active
            headers = [cell.value for cell in sheet[1]]  
            print("Headers:", headers)  

            for i in reversed(range(self.columns_layout.count())):
                self.columns_layout.itemAt(i).widget().deleteLater()

            self.checkboxes.clear()

            for header in headers:
                row_widget = QWidget()
                row_layout = QHBoxLayout()
                row_widget.setLayout(row_layout)

                column_label = QLabel(header)  
                data_checkbox = QCheckBox()
                header_checkbox = QCheckBox()

                row_layout.addWidget(column_label)
                row_layout.addWidget(data_checkbox)
                row_layout.addWidget(header_checkbox)

                self.columns_layout.addWidget(row_widget)
                self.checkboxes[header] = {"data": data_checkbox, "header": header_checkbox}

            print("Checkboxes:", self.checkboxes)  

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить Excel файл: {str(e)}")

    def generate_certificates(self):
        if not all([self.excel_file_path, self.template_file_path, self.output_folder_path, self.font_file_path]):
            QMessageBox.warning(self, "Ошибка", "Заполните все поля перед началом генерации!")
            return

        selected_columns = {}
        show_headers = {}

        for column, checkboxes in self.checkboxes.items():
            selected_columns[column] = checkboxes["data"].isChecked()
            show_headers[column] = checkboxes["header"].isChecked()

        column_spacing = {}
        for column in selected_columns:
            if selected_columns[column]:  
                column_spacing[column] = self.column_spacing.get(column, 10)  

        try:
            x_position = self.common_settings.get("x_position", 220)
            y_position = self.common_settings.get("y_position", 380)
            line_spacing = self.common_settings.get("line_spacing", 20)

            data_list = load_excel_data(self.excel_file_path)

            for i, data in enumerate(data_list):
                output_path = os.path.join(self.output_folder_path, f"certificate_{i+1}.pdf")
                print("Selected Columns:", selected_columns)
                print("Show Headers:", show_headers)
                print("Column Font Sizes:", self.column_font_sizes)
                print("Column Spacing:", self.column_spacing)
                print("Column Colors:", self.column_colors)

                create_certificate(
                    data,
                    output_path,
                    self.template_file_path,
                    self.font_file_path,
                    line_spacing,
                    x_position,
                    y_position,
                    selected_columns,
                    show_headers,
                    self.column_font_sizes,
                    column_spacing,  
                    self.column_colors  
                )

            QMessageBox.information(self, "Успех", "Сертификаты успешно созданы!")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

    def open_text_settings(self):
        if not self.checkboxes:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите Excel файл с колонками.")
            return

        selected_columns = [header for header, checkboxes in self.checkboxes.items() if checkboxes["data"].isChecked()]
        print("Selected Columns:", selected_columns)  

        if not selected_columns:
            QMessageBox.warning(self, "Ошибка", "Не выбраны колонки для отображения данных.")
            return

        dialog = TextSettingsDialog(self.template_file_path, selected_columns, self.font_file_path, self)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            settings = dialog.get_settings()
            self.common_settings = settings["common_settings"]
            self.column_font_sizes = settings["column_font_sizes"]
            self.font_file_path = settings["font_file_path"]
            self.column_spacing = settings["column_spacing"]  
            self.column_colors = settings["column_colors"]  
            print("Settings from Dialog:", settings)