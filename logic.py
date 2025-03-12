# logic.py - 修改后的逻辑代码
from PySide6.QtCore import QAbstractListModel, Qt
import pandas as pd
import random


class AttendanceLogic:
    def __init__(self):
        self.students = []  # 所有学生名单
        self.model = StudentListModel()
        self.remaining = []  # 剩余可抽学生名单

    def load_students(self, file_path):
        try:
            df = pd.read_excel(file_path)
            self.students = df.iloc[:, 0].tolist()
            self.remaining = self.students.copy()  # 初始化剩余名单
            self.model.set_data(self.students)
        except Exception as e:
            print(f"读取文件错误: {e}")

    def random_select(self, count, allow_repeat):
        if not self.students:
            return ["请先导入学生名单"]

        # 如果剩余人数不足且不允许重复，则静默重置剩余名单
        if not allow_repeat and count > len(self.remaining):
            self.remaining = self.students.copy()  # 静默重置剩余名单

        if not allow_repeat:
            selected = random.sample(self.remaining, count)
            self.remaining = [s for s in self.remaining if s not in selected]
            return selected

        return random.choices(self.students, k=count)


class StudentListModel(QAbstractListModel):
    def __init__(self):
        super().__init__()
        self._data = []

    def set_data(self, data):
        self.beginResetModel()
        self._data = data.copy()
        self.endResetModel()

    def data(self, index, role):
        if role == Qt.DisplayRole:
            return self._data[index.row()]

    def rowCount(self, index):
        return len(self._data)