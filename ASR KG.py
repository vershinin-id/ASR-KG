import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox
import pandas as pd
import os

class AdmissionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.file_path_places = ""
        self.file_path_applicants = ""
        self.file_path_save = ""
        self.groups_places = {}
        self.data_applicants = None
        self.log = []
        

        footer = QLabel("<p align=\"center\" style=\"color: #808080; font-weight: normal;\">author graduate student USURT I.Vershinin "
                        "<a href=\"https://github.com/vershinin-id\">github.com/IVershinin</a></p>")
        footer.setOpenExternalLinks(True)
        self.layout().addWidget(footer)

    def initUI(self):
        self.setWindowTitle("Распределение Абитуриентов")

        layout = QVBoxLayout()

        self.label_places = QLabel("Выберите файл с количеством мест для конкурсных групп:")
        layout.addWidget(self.label_places)

        self.choose_file_places_button = QPushButton("Выбрать файл с местами")
        self.choose_file_places_button.clicked.connect(self.choose_file_places)
        layout.addWidget(self.choose_file_places_button)

        self.label_applicants = QLabel("Выберите файл с данными абитуриентов (Excel):")
        layout.addWidget(self.label_applicants)

        self.choose_file_applicants_button = QPushButton("Выбрать файл с абитуриентами")
        self.choose_file_applicants_button.clicked.connect(self.choose_file_applicants)
        layout.addWidget(self.choose_file_applicants_button)

        self.label_save = QLabel("Выберите место для сохранения итогового файла:")
        layout.addWidget(self.label_save)

        self.choose_file_save_button = QPushButton("Выбрать место для сохранения")
        self.choose_file_save_button.clicked.connect(self.choose_file_save)
        layout.addWidget(self.choose_file_save_button)

        self.process_button = QPushButton("Распределить абитуриентов")
        self.process_button.clicked.connect(self.process_data)
        self.process_button.setEnabled(False)
        layout.addWidget(self.process_button)

        self.setLayout(layout)



    def choose_file_places(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        self.file_path_places, _ = QFileDialog.getOpenFileName(self, "Выберите файл с местами", "", "Excel файлы (*.xlsx *.xls);;All Files (*)", options=options)

        if self.file_path_places:
            self.label_places.setText(f"Файл с местами выбран: {os.path.basename(self.file_path_places)}")
            self.load_places_data()
            self.enable_process_button()
        else:
            self.label_places.setText("Файл с местами не выбран")

    def load_places_data(self):
        try:
            data_places = pd.read_excel(self.file_path_places)
            required_columns_places = ['Конкурсная группа', 'Места']
            if not all(col in data_places.columns for col in required_columns_places):
                QMessageBox.critical(self, "Ошибка", "Файл должен содержать столбцы: 'Конкурсная группа' и 'Места'.")
                self.file_path_places = ""
                return
            self.groups_places = {}
            for _, row in data_places.iterrows():
                group_name = row['Конкурсная группа']
                if group_name in self.groups_places:
                    self.groups_places[group_name] += row['Места']
                else:
                    self.groups_places[group_name] = row['Места']

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при загрузке данных о местах: {str(e)}")
            self.file_path_places = ""

    def choose_file_applicants(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        self.file_path_applicants, _ = QFileDialog.getOpenFileName(self, "Выберите файл с данными абитуриентов", "", "Excel файлы (*.xlsx *.xls);;All Files (*)", options=options)

        if self.file_path_applicants:
            self.label_applicants.setText(f"Файл с абитуриентами выбран: {os.path.basename(self.file_path_applicants)}")
            self.enable_process_button()
        else:
            self.label_applicants.setText("Файл с абитуриентами не выбран")

    def choose_file_save(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.file_path_save, _ = QFileDialog.getSaveFileName(self, "Выберите место для сохранения", "", "Excel файлы (*.xlsx);;All Files (*)", options=options)

        if self.file_path_save:
            self.label_save.setText(f"Файл для сохранения выбран: {os.path.basename(self.file_path_save)}")
            self.enable_process_button()
        else:
            self.label_save.setText("Место для сохранения не выбрано")

    def enable_process_button(self):
        if self.file_path_places and self.file_path_applicants and self.file_path_save:
            self.process_button.setEnabled(True)
        else:
            self.process_button.setEnabled(False)
    def validate_data_for_df(self, data):
            """Проверка, что все словари в списке имеют одинаковые ключи."""
            if not data:
                return True
            keys = set(data[0].keys())
            for item in data:
                if set(item.keys()) != keys:
                    return False
            return True
    def process_data(self):
        if not self.file_path_places or not self.file_path_applicants or not self.file_path_save:
            QMessageBox.critical(self, "Ошибка", "Сначала выберите все файлы.")
            return

        try:
            self.data_applicants = pd.read_excel(self.file_path_applicants)
            required_columns_applicants = ['Уникальный код', 'ФИО', 'Баллы', 'Предмет 1', 'Приоритет', 'Конкурсная группа', 'Телефон', 'Почта']
            if not all(col in self.data_applicants.columns for col in required_columns_applicants):
                QMessageBox.critical(self, "Ошибка", "Файл должен содержать столбцы: 'Уникальный код', 'ФИО', 'Баллы', 'Предмет 1', 'Приоритет', 'Конкурсная группа', 'Телефон', 'Почта'.")
                return

            # Сортируем данные по уникальному коду и приоритету, а затем по баллам в порядке убывания
            self.data_applicants.sort_values(by=['Уникальный код', 'Приоритет', 'Баллы'], ascending=[True, True, False], inplace=True)

            distribution = {group: [] for group in self.groups_places.keys()}
            unsuccessful_attempts = []
            enrolled_students = set()

            # Словарь для отслеживания приоритетов абитуриентов
            student_priorities = {}

            def allocate_student(unique_code):
                # Проверяем, есть ли уже этот абитуриент в приоритетах
                if unique_code in student_priorities:
                    current_priority = student_priorities[unique_code]
                else:
                    current_priority = 1
                    student_priorities[unique_code] = current_priority

                while current_priority <= 10:  # Предположим, что максимальное количество приоритетов = 10
                    # Получаем абитуриента с текущим приоритетом
                    applicable_rows = self.data_applicants[
                        (self.data_applicants['Уникальный код'] == unique_code) &
                        (self.data_applicants['Приоритет'] == current_priority)
                    ]

                    for _, row in applicable_rows.iterrows():
                        group = row['Конкурсная группа']
                        if group in self.groups_places and self.groups_places[group] > 0:
                            if unique_code not in enrolled_students:
                                distribution[group].append(row.to_dict())
                                self.groups_places[group] -= 1
                                enrolled_students.add(unique_code)
                                return
                        else:
                            # Проверяем, может ли текущий абитуриент вытеснить кого-то
                            lowest_score_student = min(distribution[group], key=lambda x: x['Баллы'], default=None)
                            if lowest_score_student and row['Баллы'] > lowest_score_student['Баллы']:
                                distribution[group].remove(lowest_score_student)
                                unsuccessful_attempts.append({
                                    'Уникальный код': lowest_score_student['Уникальный код'],
                                    'ФИО': lowest_score_student['ФИО'],
                                    'Баллы': lowest_score_student['Баллы'],
                                    'Предмет 1': lowest_score_student['Предмет 1'],
                                    'Приоритет': lowest_score_student['Приоритет'],
                                    'Конкурсная группа': group,
                                    'Проблема': f'Вытеснен абитуриентом с более высокими баллами (код {unique_code}), '
                                    })
                                enrolled_students.remove(lowest_score_student['Уникальный код'])
                                allocate_student(lowest_score_student['Уникальный код'])
                                distribution[group].append(row.to_dict())
                                enrolled_students.add(unique_code)
                                return

                    # Если не удалось зачислить, проверяем следующий приоритет
                    student_priorities[unique_code] = current_priority + 1
                    current_priority += 1

                # Если абитуриент не был зачислен в течение всех приоритетов
                unsuccessful_attempts.append({
                    'Уникальный код': unique_code,
                    'ФИО': self.data_applicants[self.data_applicants['Уникальный код'] == unique_code].iloc[0]['ФИО'],
                    'Баллы': self.data_applicants[self.data_applicants['Уникальный код'] == unique_code].iloc[0]['Баллы'],
                    'Предмет 1': self.data_applicants[self.data_applicants['Уникальный код'] == unique_code].iloc[0]['Предмет 1'],
                    'Приоритет': self.data_applicants[self.data_applicants['Уникальный код'] == unique_code].iloc[0]['Приоритет'],
                    'Конкурсная группа': self.data_applicants[self.data_applicants['Уникальный код'] == unique_code].iloc[0]['Конкурсная группа'],
                    'Проблема': 'Не поступил ни в одну группу'
                })

            # Запускаем распределение для всех уникальных абитуриентов
            for unique_code in self.data_applicants['Уникальный код'].unique():
                allocate_student(unique_code)

            # Создание отчетов
            with pd.ExcelWriter(self.file_path_save, engine='xlsxwriter') as writer:
                df_vacant_places = pd.DataFrame({
                    'Конкурсная группа': list(self.groups_places.keys()),
                    'Вакантные места': list(self.groups_places.values())
                })
                df_vacant_places.to_excel(writer, sheet_name="Вакантные места", index=False)

                df_pass_scores = pd.DataFrame({
                    'Конкурсная группа': list(distribution.keys()),
                    'Минимальный балл': [min([student['Баллы'] for student in students]) for students in distribution.values() if students]
                })
                df_pass_scores.to_excel(writer, sheet_name="Минимальные баллы", index=False)

                for group, students in distribution.items():
                    df_group = pd.DataFrame(students)
                    sheet_name = self._get_unique_sheet_name(writer, group[:31])
                    df_group.to_excel(writer, sheet_name=sheet_name, index=False)

                if unsuccessful_attempts:
                    df_unsuccessful = pd.DataFrame(unsuccessful_attempts)
                    df_unsuccessful.to_excel(writer, sheet_name="Не прошедшие", index=False)

            QMessageBox.information(self, "Успех", "Распределение успешно завершено и сохранено.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при обработке данных: {str(e)}")


    def _get_unique_sheet_name(self, writer, base_name):
        sheet_names = writer.book.sheetnames
        sheet_name = base_name
        counter = 1
        while sheet_name in sheet_names:
            sheet_name = f"{base_name[:28]}_{counter}"
            counter += 1
        return sheet_name

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AdmissionApp()
    ex.show()
    sys.exit(app.exec_())


