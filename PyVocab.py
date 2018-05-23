from PyQt5.QtWidgets import *
from PyQt5.QtGui import QFont
from pyexcel_xls import *
import sys
import os
import pickle
import random


class Subject:

    def __init__(self, name):
        self.name = name
        self.lessons = {}
        self.is_mistake_book = False

    def new_lesson(self, lesson_name):
        self.lessons[lesson_name] = Lesson(self, lesson_name)
        return self.lessons[lesson_name]


class Lesson:

    def __init__(self, subject, name):
        self.subject = subject
        self.name = name
        self.dict = []

    def save_dict(self):
        with open('obj/' + self.subject.name + '/' + self.name + '.pkl', 'wb') as f:
            pickle.dump(self.dict, f, pickle.DEFAULT_PROTOCOL)
        self.dict = None

    def clear_dict_ref(self):
        self.dict = None

    def load_dict(self):
        with open('obj/' + self.subject.name + '/' + self.name + '.pkl', 'rb') as f:
            self.dict = pickle.load(f)

    def add_entry(self, a, b):
        self.dict.append(Entry(self, a, b))


class Entry:

    def __init__(self, lesson, a, b):
        self.lesson = lesson
        self.ent1 = a
        self.ent2 = b


class PyVocab(QWidget):

    main_data = {}

    def __init__(self, data):
        super().__init__()
        self.main_data = data
        self.setFixedSize(600, 400)
        self.setWindowTitle('PyVocab')

        self.subject_select_label = QLabel('Subjects', self)
        self.subject_select_label.move(20, 20)
        self.subject_select = QScrollArea(self)
        self.subject_select.setGeometry(20, 40, 110, 315)
        self.subject_list = QListWidget(self)
        self.subject_list.setFixedWidth(108)
        self.subject_list.setMinimumHeight(313)
        self.update_subjects()
        self.subject_select.setWidget(self.subject_list)
        self.subject_list.clicked.connect(self.update_lessons)

        self.lesson_select_label = QLabel('Lessons', self)
        self.lesson_select_label.move(150, 20)
        self.lesson_select = QScrollArea(self)
        self.lesson_select.setGeometry(150, 40, 110, 315)
        self.lesson_list = QListWidget(self)
        self.lesson_list.setFixedWidth(108)
        self.lesson_list.setMinimumHeight(313)
        self.lesson_select.setWidget(self.lesson_list)
        self.lesson_list.clicked.connect(self.update_vocab)

        self.vocab_label = QLabel('Vocabulary', self)
        self.vocab_label.move(280, 20)
        self.vocab = QScrollArea(self)
        self.vocab.setGeometry(280, 40, 300, 315)
        self.vocab_list = QTableWidget(self)
        self.vocab_list.setFixedWidth(298)
        self.vocab_list.setMinimumHeight(313)
        self.vocab_list.setColumnCount(2)
        self.vocab_list.setHorizontalHeaderLabels(['Entry 1', 'Entry 2'])
        self.vocab_list.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.vocab.setWidget(self.vocab_list)

        self.btn_new_subject = QPushButton('New', self)
        self.btn_new_subject.clicked.connect(self.new_subject)
        self.btn_new_subject.move(15, 360)

        self.btn_import_dict = QPushButton('Import', self)
        self.import_window = DictImporter(self.main_data, self)
        self.btn_import_dict.clicked.connect(self.import_dict_helper)
        self.btn_import_dict.move(145, 360)

        self.btn_learn = QPushButton('Learn', self)
        self.learn_window = LearnWindow(self.main_data, self)
        self.btn_learn.clicked.connect(self.learn_helper)
        self.btn_learn.move(275, 360)

        self.btn_test = QPushButton('Test', self)
        self.test_window = TestWindow(self.main_data, self)
        self.btn_test.clicked.connect(self.test_helper)
        self.btn_test.move(350, 360)

        self.btn_remove = QPushButton('Remove', self)
        self.btn_remove.clicked.connect(self.remove)
        self.btn_remove.move(495, 360)

        self.btn_import_dict = QPushButton('Bulk', self)
        self.bulk_import_window = BulkDictImporter(self.main_data, self)
        self.btn_import_dict.clicked.connect(self.bulk_import_dict_helper)
        self.btn_import_dict.move(215, 360)

        self.report_window = ReportWindow(self.main_data, self)

    def new_subject(self):
        name, ok = QInputDialog.getText(self, 'Create New Subject', 'Subject Name:')
        if ok:
            if name in main_data['Subjects']:
                QMessageBox.warning(self, 'Warning', 'Subject already exists.')
            else:
                main_data['Subjects'][name] = Subject(name)
                os.mkdir('obj/' + name)
        self.update_subjects()

    def bulk_import_dict_helper(self):
        item = self.subject_list.currentItem()
        if item is None:
            QMessageBox.warning(self, 'Warning', 'Select subject first!')
        else:
            subject = item.text()
            self.bulk_import_window.import_dict(subject)

    def import_dict_helper(self):
        item = self.subject_list.currentItem()
        if item is None:
            QMessageBox.warning(self, 'Warning', 'Select subject first!')
        else:
            subject = item.text()
            self.import_window.import_dict(subject)

    def learn_helper(self):
        if not self.valid_lesson():
            QMessageBox.warning(self, 'Warning', 'Select a lesson first!')
            return
        subject = self.subject_list.currentItem().text()
        lesson = self.lesson_list.currentItem().text()
        self.learn_window.setWindowTitle('Learn: ' + subject + ' - ' + lesson)
        self.learn_window.setGeometry(self.geometry())
        lesson_obj = self.main_data["Subjects"][subject].lessons[lesson]
        lesson_obj.load_dict()
        self.learn_window.lesson = lesson_obj
        self.learn_window.init()
        self.learn_window.show()
        self.hide()

    def test_helper(self):
        if not self.valid_lesson():
            QMessageBox.warning(self, 'Warning', 'Select a lesson first!')
            return
        subject = self.subject_list.currentItem().text()
        lesson = self.lesson_list.currentItem().text()
        self.test_window.setWindowTitle('Test: ' + subject + ' - ' + lesson)
        self.test_window.setGeometry(self.geometry())
        lesson_obj = self.main_data["Subjects"][subject].lessons[lesson]
        lesson_obj.load_dict()
        random.shuffle(lesson_obj.dict)
        self.test_window.shuffled_list = lesson_obj.dict
        lesson_obj.clear_dict_ref()
        self.test_window.init()
        self.test_window.show()
        self.hide()

    def valid_lesson(self):
        subject = self.subject_list.currentItem()
        lesson = self.lesson_list.currentItem()
        if subject is None:
            return False
        if lesson is None:
            return False
        return True

    def update_subjects(self):
        self.subject_list.clear()
        for subject in self.main_data['Subjects']:
            item = QListWidgetItem(self.subject_list)
            item.setText(subject)
            self.subject_list.addItem(item)

    def update_lessons(self):
        self.lesson_list.clear()
        subject = self.subject_list.currentItem().text()
        for lesson in self.main_data['Subjects'][subject].lessons:
            item = QListWidgetItem(self.lesson_list)
            item.setText(lesson)
            self.lesson_list.addItem(item)

    def update_vocab(self):
        subject = self.subject_list.currentItem().text()
        lesson_name = self.lesson_list.currentItem().text()
        lesson = self.main_data['Subjects'][subject].lessons[lesson_name]
        lesson.load_dict()
        row_count = 0
        for e in lesson.dict:
            if self.vocab_list.rowCount() <= row_count:
                self.vocab_list.insertRow(row_count)
            ent1 = QTableWidgetItem()
            ent1.setText(e.ent1)
            ent2 = QTableWidgetItem()
            ent2.setText(e.ent2)
            self.vocab_list.setItem(row_count, 0, ent1)
            self.vocab_list.setItem(row_count, 1, ent2)
            row_count = row_count + 1
        while self.vocab_list.rowCount() > row_count:
            self.vocab_list.removeRow(row_count)
        lesson.clear_dict_ref()

    def remove(self):
        if not self.valid_lesson():
            QMessageBox.warning(self, 'Warning', 'Select a lesson first!')
            return
        subject = self.subject_list.currentItem().text()
        lesson_name = self.lesson_list.currentItem().text()
        self.main_data['Subjects'][subject].lessons.pop(lesson_name)
        os.remove('obj/' + subject + '/' + lesson_name + '.pkl')
        self.update_lessons()

    def generate_report(self, mistake):
        subject = self.subject_list.currentItem().text()
        lesson = self.lesson_list.currentItem().text()

        if main_data['Subjects'][subject].is_mistake_book:
            mistake_book_name = subject
        else:
            mistake_book_name = "E - " + subject

        if not mistake_book_name in main_data['Subjects']:
            book = Subject(mistake_book_name)
            main_data['Subjects'][mistake_book_name] = book
            book.is_mistake_book = True
            os.mkdir('obj/' + mistake_book_name)
            self.update_subjects()
        else:
            book = main_data['Subjects'][mistake_book_name]

        if not lesson in book.lessons:
            list = book.new_lesson(lesson)
        else:
            list = book.lessons[lesson]

        list.dict = []
        for e in mistake:
            list.add_entry(e.ent1, e.ent2)
        list.save_dict()
        list.clear_dict_ref()

        self.report_window.setWindowTitle('Report: ' + subject + ' - ' + lesson)
        self.report_window.setGeometry(self.geometry())
        self.report_window.init(mistake)
        self.report_window.show()
        self.hide()

    def closeEvent(self, event):
        self.import_window.hide()
        event.accept()


class DictImporter(QWidget):

    def __init__(self, data, parent):
        super().__init__()
        self.main_data = data
        self.parent = parent
        self.subject = None
        self.picked_file = None
        self.setFixedSize(240, 120)
        self.name_msg = QLabel('Lesson Name:', self)
        self.name_msg.move(20, 20)
        self.name_input = QLineEdit(self)
        self.name_input.setFixedWidth(100)
        self.name_input.move(115, 20)
        self.file_name_prompt = QLabel('File selected:', self)
        self.file_name_prompt.move(20, 50)
        self.file_name = QLabel('None', self)
        self.file_name.setFixedWidth(100)
        self.file_name.move(115, 50)
        self.btn_select_file = QPushButton('Select', self)
        self.btn_select_file.move(30, 75)
        self.btn_select_file.clicked.connect(self.pick_file)
        self.selected_file = QLabel(self)
        self.ok = QPushButton('Import', self)
        self.ok.move(125, 75)
        self.ok.clicked.connect(self.load)

    def import_dict(self, subject):
        self.subject = subject
        self.setWindowTitle('Import to ' + subject)
        self.show()

    def pick_file(self):
        self.picked_file = QFileDialog.getOpenFileName(self, "", "", "(*.xlsx)")[0]
        i = len(self.picked_file) - 1
        while i >= 0:
            if self.picked_file[i] == '/':
                self.file_name.setText(self.picked_file[i + 1:len(self.picked_file) - 1])
                return
            i = i - 1
        self.picked_file = None

    def load(self):
        if self.name_input.text() == '':
            QMessageBox.warning(self, 'Warning', 'Please provide lesson name.')
            return
        if self.picked_file is None:
            QMessageBox.warning(self, 'Warning', 'Please select vocabulary file.')
            return
        self.hide()
        vocab_data = get_data(self.picked_file)['Sheet1']
        lesson = self.main_data['Subjects'][self.subject].new_lesson(self.name_input.text())
        for d in vocab_data:
            if len(d) == 1:
                for i in range(len(d[0])):
                    if d[0][i] == ' ' or d[0][i] == ':':
                        lesson.add_entry(d[0][0:i], d[0][i + 1:])
                        break
            elif len(d) == 2:
                lesson.add_entry(d[0], d[1])
            else:
                QMessageBox.warning(self, 'Warning', 'Failed to process file.')
                return
        lesson.save_dict()
        lesson.clear_dict_ref()
        self.subject = None
        self.picked_file = None
        self.parent.update_lessons()
        self.name_input.setText('')
        self.file_name.setText('None')


class BulkDictImporter(QWidget):

    def __init__(self, data, parent):
        super().__init__()
        self.main_data = data
        self.parent = parent
        self.subject = None
        self.picked_file = None
        self.setFixedSize(240, 120)
        self.name_msg = QLabel('Lesson Prefix:', self)
        self.name_msg.move(20, 20)
        self.name_input = QLineEdit(self)
        self.name_input.setFixedWidth(100)
        self.name_input.move(115, 20)
        self.file_name_prompt = QLabel('File selected:', self)
        self.file_name_prompt.move(20, 50)
        self.file_name = QLabel('None', self)
        self.file_name.setFixedWidth(100)
        self.file_name.move(115, 50)
        self.btn_select_file = QPushButton('Select', self)
        self.btn_select_file.move(30, 75)
        self.btn_select_file.clicked.connect(self.pick_file)
        self.selected_file = QLabel(self)
        self.ok = QPushButton('Import', self)
        self.ok.move(125, 75)
        self.ok.clicked.connect(self.load)

    def import_dict(self, subject):
        self.subject = subject
        self.setWindowTitle('Import to ' + subject)
        self.show()

    def pick_file(self):
        self.picked_file = QFileDialog.getOpenFileName(self, "", "", "(*.xlsx)")[0]
        i = len(self.picked_file) - 1
        while i >= 0:
            if self.picked_file[i] == '/':
                self.file_name.setText(self.picked_file[i + 1:len(self.picked_file) - 1])
                return
            i = i - 1
        self.picked_file = None

    def load(self):
        if self.name_input.text() == '':
            QMessageBox.warning(self, 'Warning', 'Please provide lesson name.')
            return
        if self.picked_file is None:
            QMessageBox.warning(self, 'Warning', 'Please select vocabulary file.')
            return
        self.hide()
        vocab_data = get_data(self.picked_file)['Sheet1']
        lessons = {}
        for d in vocab_data:
            if not d[0] in lessons:
                lessons[d[0]] = self.main_data['Subjects'][self.subject].new_lesson(self.name_input.text() + " " + str(d[0]))
            lessons[d[0]].add_entry(d[1], d[2])
        for lesson in lessons:
            lessons[lesson].save_dict()
            lessons[lesson].clear_dict_ref()
        self.subject = None
        self.picked_file = None
        self.parent.update_lessons()
        self.name_input.setText('')
        self.file_name.setText('None')


class LearnWindow(QWidget):

    def __init__(self, data, parent):
        super().__init__()
        self.main_data = data
        self.parent = parent
        self.setFixedSize(600, 400)
        self.lesson = None
        self.pointer = 0

        self.entry1 = QTextBrowser(self)
        self.entry1.setGeometry(20, 20, 560, 125)
        self.entry1.setFont(QFont("Roman times", 30, QFont.Normal))

        self.entry2 = QTextBrowser(self)
        self.entry2.setGeometry(20, 165, 560, 125)
        self.entry2.setFont(QFont("Roman times", 30, QFont.Normal))

        self.btn_prev = QPushButton("←", self)
        self.btn_prev.setFont(QFont("Roman times", 30, QFont.Normal))
        self.btn_prev.setGeometry(20, 310, 270, 70)
        self.btn_prev.clicked.connect(self.prev)

        self.btn_next = QPushButton("→", self)
        self.btn_next.setFont(QFont("Roman times", 30, QFont.Normal))
        self.btn_next.setGeometry(310, 310, 270, 70)
        self.btn_next.clicked.connect(self.next)

    def init(self):
        entry = self.lesson.dict[self.pointer]
        self.entry1.setText(entry.ent1)
        self.entry2.setText(entry.ent2)
        self.change_btn_state()

    def prev(self):
        self.pointer -= 1
        entry = self.lesson.dict[self.pointer]
        self.entry1.setText(entry.ent1)
        self.entry2.setText(entry.ent2)
        self.change_btn_state()

    def next(self):
        self.pointer += 1
        entry = self.lesson.dict[self.pointer]
        self.entry1.setText(entry.ent1)
        self.entry2.setText(entry.ent2)
        self.change_btn_state()

    def change_btn_state(self):
        if self.pointer == 0:
            self.btn_prev.setEnabled(False)
        else:
            self.btn_prev.setEnabled(True)
        if self.pointer >= len(self.lesson.dict) - 1:
            self.btn_next.setEnabled(False)
        else:
            self.btn_next.setEnabled(True)

    def closeEvent(self, event):
        self.lesson.clear_dict_ref()
        self.parent.setGeometry(self.geometry())
        self.parent.show()
        event.accept()


class TestWindow(QWidget):

    def __init__(self, data, parent):
        super().__init__()
        self.main_data = data
        self.parent = parent
        self.setFixedSize(600, 400)
        self.shuffled_list = None
        self.pointer = 0
        self.mistake = []

        self.entry1 = QTextBrowser(self)
        self.entry1.setGeometry(20, 20, 560, 125)
        self.entry1.setFont(QFont("Roman times", 30, QFont.Normal))

        self.entry2 = QTextBrowser(self)
        self.entry2.setGeometry(20, 165, 560, 125)
        self.entry2.setFont(QFont("Roman times", 30, QFont.Normal))

        self.btn_show = QPushButton("Show Answer", self)
        self.btn_show.setFont(QFont("Roman times", 30, QFont.Normal))
        self.btn_show.setGeometry(20, 310, 270, 70)
        self.btn_show.clicked.connect(self.show_ans)

        self.btn_correct = QPushButton("√", self)
        self.btn_correct.setFont(QFont("Roman times", 30, QFont.Normal))
        self.btn_correct.setGeometry(310, 310, 125, 70)
        self.btn_correct.clicked.connect(self.yes)

        self.btn_error = QPushButton("×", self)
        self.btn_error.setFont(QFont("Roman times", 30, QFont.Normal))
        self.btn_error.setGeometry(455, 310, 125, 70)
        self.btn_error.clicked.connect(self.no)

    def init(self):
        self.mistake = []
        self.pointer = 0
        entry = self.shuffled_list[self.pointer]
        self.entry1.setText(entry.ent1)
        self.entry2.setText('')
        self.change_btn_state()

    def yes(self):
        self.pointer += 1
        if self.pointer == len(self.shuffled_list):
            self.report()
            return
        entry = self.shuffled_list[self.pointer]
        self.entry1.setText(entry.ent1)
        self.entry2.setText('')
        self.change_btn_state()

    def no(self):
        self.mistake.append(self.shuffled_list[self.pointer])
        self.pointer += 1
        if self.pointer == len(self.shuffled_list):
            self.report()
            return
        entry = self.shuffled_list[self.pointer]
        self.entry1.setText(entry.ent1)
        self.entry2.setText('')
        self.change_btn_state()

    def show_ans(self):
        entry = self.shuffled_list[self.pointer]
        self.entry2.setText(entry.ent2)

    def change_btn_state(self):
        if self.pointer >= len(self.shuffled_list):
            self.btn_correct.setEnabled(False)
            self.btn_error.setEnabled(False)

    def report(self):
        self.parent.setGeometry(self.geometry())
        self.parent.generate_report(self.mistake)
        self.hide()

    def closeEvent(self, event):
        self.report()
        event.accept()

class ReportWindow(QWidget):

    def __init__(self, data, parent):
        super().__init__()
        self.main_data = data
        self.parent = parent
        self.setFixedSize(600, 400)

        self.num_label = QLabel('######################', self)
        self.num_label.move(50, 50)

        self.vocab_label = QLabel('Mistakes', self)
        self.vocab_label.move(280, 20)
        self.vocab = QScrollArea(self)
        self.vocab.setGeometry(280, 40, 300, 315)
        self.vocab_list = QTableWidget(self)
        self.vocab_list.setFixedWidth(298)
        self.vocab_list.setMinimumHeight(313)
        self.vocab_list.setColumnCount(2)
        self.vocab_list.setHorizontalHeaderLabels(['Entry 1', 'Entry 2'])
        self.vocab_list.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.vocab.setWidget(self.vocab_list)

        self.mistake = []

    def init(self, mistake):
        self.mistake = mistake
        if not mistake:
            self.num_label.setText('All correct!')
        else:
            self.num_label.setText('Number of mistakes: ' + str(len(mistake)))

        row_count = 0
        for e in mistake:
            if self.vocab_list.rowCount() <= row_count:
                self.vocab_list.insertRow(row_count)
            ent1 = QTableWidgetItem()
            ent1.setText(e.ent1)
            ent2 = QTableWidgetItem()
            ent2.setText(e.ent2)
            self.vocab_list.setItem(row_count, 0, ent1)
            self.vocab_list.setItem(row_count, 1, ent2)
            row_count = row_count + 1
        while self.vocab_list.rowCount() > row_count:
            self.vocab_list.removeRow(row_count)

    def change_btn_state(self):
        if self.pointer >= len(self.shuffled_list):
            self.btn_correct.setEnabled(False)
            self.btn_error.setEnabled(False)

    def closeEvent(self, event):
        self.parent.setGeometry(self.geometry())
        self.parent.update_vocab()
        self.parent.show()
        event.accept()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    try:
        with open('obj/main_data.pkl', 'rb') as file:
            main_data = pickle.load(file)
    except IOError:
        if not os.path.exists('obj'):
            os.mkdir('obj')
        main_data = {'Subjects': {}, 'Tests': {}}
    main = PyVocab(main_data)
    main.show()
    app.exec_()
    with open('obj/main_data.pkl', 'wb') as file:
        pickle.dump(main_data, file, pickle.DEFAULT_PROTOCOL)
    sys.exit()
