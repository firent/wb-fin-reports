import sys
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog,
    QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem,
    QMessageBox, QHBoxLayout, QLabel
)
from PySide6.QtCore import Qt

class ReportGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Генератор отчетов по финансовым отчетам Wildberries")
        self.setGeometry(100, 100, 1000, 600)

        self.df = None  # Для хранения загруженных данных
        self.summary_df = None # Для хранения итоговых данных

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # --- Панель управления ---
        controls_layout = QHBoxLayout()
        
        self.label_file = QLabel("Файл не выбран")
        controls_layout.addWidget(self.label_file)

        self.btn_load = QPushButton("Выбрать файл")
        self.btn_load.clicked.connect(self.load_file)
        controls_layout.addWidget(self.btn_load)

        self.btn_save = QPushButton("Сохранить отчет")
        self.btn_save.clicked.connect(self.save_report)
        self.btn_save.setEnabled(False) # Кнопка неактивна, пока нет данных
        controls_layout.addWidget(self.btn_save)
        
        layout.addLayout(controls_layout)

        # --- Таблица для отображения итогов ---
        self.table_summary = QTableWidget()
        # Устанавливаем политику размера, чтобы таблица растягивалась
        self.table_summary.horizontalHeader().setStretchLastSection(True) 
        layout.addWidget(self.table_summary)

    def load_file(self):
        options = QFileDialog.Option.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx *.xls)", options=options
        )
        if file_name:
            self.label_file.setText(f"Выбран файл: {file_name}")
            try:
                self.process_data(file_name)
                self.display_summary()
                self.btn_save.setEnabled(True)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось обработать файл: {e}")

    def process_data(self, file_path):
        # Загружаем данные
        self.df = pd.read_excel(file_path, sheet_name=0) # Читаем первый лист

        # --- Расчет итогов ---
        sales_df = self.df.copy()

        if sales_df.empty:
             QMessageBox.warning(self, "Предупреждение", "В файле не найдено строк с типом документа 'Продажа'.")
             self.summary_df = pd.DataFrame() # Создаем пустой DataFrame
             return

        # Группируем по артикулу поставщика и названию
        summary = sales_df.groupby(['Артикул поставщика', 'Название'], as_index=False).agg({
            'Кол-во': 'sum',
            'Вайлдберриз реализовал Товар (Пр)': 'sum', # Суммируем выручку
            'К перечислению Продавцу за реализованный Товар': 'sum', # Суммируем доход
            'Услуги по доставке товара покупателю': 'sum' # Суммируем расходы на логистику
        })

        # Переименовываем колонки для ясности
        summary.rename(columns={
            'Артикул поставщика': 'Артикул',
            'Название': 'Наименование',
            'Кол-во': 'Количество',
            'Вайлдберриз реализовал Товар (Пр)': 'Выручка',
            'К перечислению Продавцу за реализованный Товар': 'Доход',
            'Услуги по доставке товара покупателю': 'Логистика'
        }, inplace=True)
        
        # Добавляем колонку "Прибыль" (Доход - Логистика)
        summary['Прибыль'] = summary['Доход'] - summary['Логистика']

        self.summary_df = summary

    def display_summary(self):
        self.table_summary.clear()
        if self.summary_df is None or self.summary_df.empty:
            self.table_summary.setRowCount(0)
            self.table_summary.setColumnCount(0)
            return

        rows, cols = self.summary_df.shape
        self.table_summary.setRowCount(rows)
        self.table_summary.setColumnCount(cols)
        self.table_summary.setHorizontalHeaderLabels(self.summary_df.columns.tolist())

        for r in range(rows):
            for c in range(cols):
                item = QTableWidgetItem(str(self.summary_df.iat[r, c]))
                item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable) # Делаем ячейки не редактируемыми
                self.table_summary.setItem(r, c, item)

        self.table_summary.resizeColumnsToContents()

    def save_report(self):
        if self.summary_df is None or self.summary_df.empty:
            QMessageBox.warning(self, "Предупреждение", "Нет данных для сохранения.")
            return

        options = QFileDialog.Option.ReadOnly
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчет", "итоги_отчета.xlsx", "Excel Files (*.xlsx)", options=options
        )
        if file_path:
            # Убедимся, что файл имеет расширение .xlsx
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            try:
                self.summary_df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Успех", f"Отчет успешно сохранен в {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ReportGeneratorApp()
    window.show()
    sys.exit(app.exec())
