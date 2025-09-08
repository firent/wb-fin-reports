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
        self.setGeometry(50, 50, 1200, 700)

        self.df = None  # Для хранения загруженных данных
        self.summary_df = None # Для хранения итоговых данных
        self.total_row = None # Для хранения итоговой строки

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
             QMessageBox.warning(self, "Предупреждение", "В файле не найдено данных.")
             self.summary_df = pd.DataFrame() # Создаем пустой DataFrame
             self.total_row = None
             return

        # Группируем по артикулу поставщика и названию
        summary = sales_df.groupby(['Артикул поставщика', 'Название'], as_index=False, dropna=False).agg({
            'Кол-во': 'sum',
            'Вайлдберриз реализовал Товар (Пр)': 'sum', # Суммируем выручку
            'К перечислению Продавцу за реализованный Товар': 'sum', # Суммируем доход
            'Услуги по доставке товара покупателю': 'sum', # Суммируем расходы на логистику
            'Хранение': 'sum', # Суммируем расходы на хранение
            'Удержания': 'sum', # Суммируем удержания
            'Платная приемка': 'sum', # Расходы на платную приемку
            'Компенсация скидки по программе лояльности': 'sum', # Суммируем компенсацию скидки
            'Общая сумма штрафов': 'sum' # Суммируем общую сумму штрафов
        })

        # Переименовываем колонки для ясности
        summary.rename(columns={
            'Артикул поставщика': 'Артикул',
            'Название': 'Наименование',
            'Кол-во': 'Количество',
            'Вайлдберриз реализовал Товар (Пр)': 'Выручка',
            'К перечислению Продавцу за реализованный Товар': 'Доход',
            'Услуги по доставке товара покупателю': 'Логистика',
            'Компенсация скидки по программе лояльности': 'Компенсация скидки',
            'Общая сумма штрафов': 'Штрафы'
        }, inplace=True)
        
        # Добавляем колонку "Прибыль" (Доход - все расходы)
        summary['Прибыль'] = (
            summary['Доход'] - 
            summary['Логистика'] - 
            summary['Хранение'] - 
            summary['Удержания'] - 
            summary['Платная приемка'] - 
            summary['Штрафы']
        )

        # Округляем все числовые колонки до 2 знаков после запятой
        numeric_columns = summary.select_dtypes(include=['number']).columns
        summary[numeric_columns] = summary[numeric_columns].round(2)

        # Создаем итоговую строку
        total_data = {
            'Артикул': 'ИТОГО',
            'Наименование': '',
            'Количество': summary['Количество'].sum(),
            'Выручка': summary['Выручка'].sum(),
            'Доход': summary['Доход'].sum(),
            'Логистика': summary['Логистика'].sum(),
            'Хранение': summary['Хранение'].sum(),
            'Удержания': summary['Удержания'].sum(),
            'Платная приемка': summary['Платная приемка'].sum(),
            'Компенсация скидки': self.df['Компенсация скидки по программе лояльности'].sum(),
            'Штрафы': summary['Штрафы'].sum(),
            'Прибыль': summary['Прибыль'].sum()
        }
        
        # Округляем итоговые значения
        for key, value in total_data.items():
            if isinstance(value, (int, float)):
                total_data[key] = round(value, 2)
        
        self.total_row = total_data
        self.summary_df = summary

    def display_summary(self):
        self.table_summary.clear()
        if self.summary_df is None or self.summary_df.empty:
            self.table_summary.setRowCount(0)
            self.table_summary.setColumnCount(0)
            return

        rows, cols = self.summary_df.shape
        # Добавляем дополнительную строку для итогов
        self.table_summary.setRowCount(rows + 1)
        self.table_summary.setColumnCount(cols)
        self.table_summary.setHorizontalHeaderLabels(self.summary_df.columns.tolist())

        # Заполняем данные товаров
        for r in range(rows):
            for c in range(cols):
                value = self.summary_df.iat[r, c]
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                self.table_summary.setItem(r, c, item)

        # Добавляем итоговую строку
        if self.total_row:
            for c, col_name in enumerate(self.summary_df.columns):
                value = self.total_row.get(col_name, '')
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                
                # Выделяем итоговую строку жирным шрифтом
                font = item.font()
                font.setBold(True)
                item.setFont(font)
                item.setBackground(Qt.GlobalColor.lightGray)
                
                self.table_summary.setItem(rows, c, item)

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
                # Создаем копию DataFrame для сохранения
                save_df = self.summary_df.copy()
                
                # Добавляем итоговую строку как новую строку в DataFrame
                if self.total_row:
                    total_series = pd.Series(self.total_row)
                    save_df = pd.concat([save_df, total_series.to_frame().T], ignore_index=True)
                
                save_df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Успех", f"Отчет успешно сохранен в {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ReportGeneratorApp()
    window.show()
    sys.exit(app.exec())
