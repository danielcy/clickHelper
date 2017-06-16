import datetime
import win32gui
import win32api
import win32con
import sys
import time
from PySide import QtGui, QtCore

pos = (862, 319)

class ClickHelperMain(QtGui.QWidget):
    def __init__(self, parent=None):
        super(ClickHelperMain, self).__init__(parent)
        self.hwnd_list = []
        self.get_hwnd()
        self.setup_ui()
        self.setWindowTitle("Chrome Click Helper")
        self.running = False
        self.proc_thread = None
        self.proc_thread_list = []

    def get_hwnd(self):
        self.hwnd_list = []
        win32gui.EnumWindows(self._hwnd_handler, 0)

    def _hwnd_handler(self, hwnd, i):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            name = win32gui.GetWindowText(hwnd)
            if name.find("Google Chrome") > -1:
                self.hwnd_list.append(hwnd)

    def setup_ui(self):
        self.mainLayout = QtGui.QVBoxLayout()
        self.setLayout(self.mainLayout)

        self.count_label_layout = QtGui.QHBoxLayout()
        self.mainLayout.addLayout(self.count_label_layout)
        self.text_label = QtGui.QLabel("Current Chrome Windows Count: ")
        self.chrome_count_label = QtGui.QLabel(str(len(self.hwnd_list)))
        self.count_label_layout.addWidget(self.text_label)
        self.count_label_layout.addWidget(self.chrome_count_label)

        self.time_group = QtGui.QGroupBox("Time Setting")
        self.mainLayout.addWidget(self.time_group)
        self.time_group_layout = QtGui.QVBoxLayout()
        self.time_group.setLayout(self.time_group_layout)
        self.time_basic_layout = QtGui.QHBoxLayout()
        self.time_group_layout.addLayout(self.time_basic_layout)
        self.hour_label = QtGui.QLabel("Hour:       ")
        self.hour_edit = QtGui.QLineEdit()
        self.min_label = QtGui.QLabel("Min:      ")
        self.min_edit = QtGui.QLineEdit()
        self.time_basic_layout.addWidget(self.hour_label)
        self.time_basic_layout.addWidget(self.hour_edit)
        self.time_basic_layout.addWidget(self.min_label)
        self.time_basic_layout.addWidget(self.min_edit)
        self.time_advance_layout = QtGui.QHBoxLayout()
        self.time_group_layout.addLayout(self.time_advance_layout)
        self.offset_label = QtGui.QLabel("Offset(ms): ")
        self.offset_edit = QtGui.QLineEdit()
        self.step_label = QtGui.QLabel("Step(ms): ")
        self.step_edit = QtGui.QLineEdit()
        self.time_advance_layout.addWidget(self.offset_label)
        self.time_advance_layout.addWidget(self.offset_edit)
        self.time_advance_layout.addWidget(self.step_label)
        self.time_advance_layout.addWidget(self.step_edit)

        self.pos_group = QtGui.QGroupBox("Position Setting(Pixel)")
        self.mainLayout.addWidget(self.pos_group)
        self.pos_group_layout = QtGui.QHBoxLayout()
        self.pos_group.setLayout(self.pos_group_layout)
        self.x1_label = QtGui.QLabel("X1: ")
        self.x1_edit = QtGui.QLineEdit("73")
        self.y1_label = QtGui.QLabel("Y1: ")
        self.y1_edit = QtGui.QLineEdit("42")
        self.x2_label = QtGui.QLabel("X2: ")
        self.x2_edit = QtGui.QLineEdit("931")
        self.y2_label = QtGui.QLabel("Y2: ")
        self.y2_edit = QtGui.QLineEdit("316")
        self.x3_label = QtGui.QLabel("X3: ")
        self.x3_edit = QtGui.QLineEdit("1075")
        self.y3_label = QtGui.QLabel("Y3: ")
        self.y3_edit = QtGui.QLineEdit("430")
        self.x4_label = QtGui.QLabel("X4: ")
        self.x4_edit = QtGui.QLineEdit("966")
        self.y4_label = QtGui.QLabel("Y4: ")
        self.y4_edit = QtGui.QLineEdit("487")
        self.x5_label = QtGui.QLabel("X5: ")
        self.x5_edit = QtGui.QLineEdit("853")
        self.y5_label = QtGui.QLabel("Y5: ")
        self.y5_edit = QtGui.QLineEdit("580")
        self.pos_group_layout.addWidget(self.x1_label)
        self.pos_group_layout.addWidget(self.x1_edit)
        self.pos_group_layout.addWidget(self.y1_label)
        self.pos_group_layout.addWidget(self.y1_edit)
        self.pos_group_layout.addWidget(self.x2_label)
        self.pos_group_layout.addWidget(self.x2_edit)
        self.pos_group_layout.addWidget(self.y2_label)
        self.pos_group_layout.addWidget(self.y2_edit)
        self.pos_group_layout.addWidget(self.x3_label)
        self.pos_group_layout.addWidget(self.x3_edit)
        self.pos_group_layout.addWidget(self.y3_label)
        self.pos_group_layout.addWidget(self.y3_edit)
        self.pos_group_layout.addWidget(self.x4_label)
        self.pos_group_layout.addWidget(self.x4_edit)
        self.pos_group_layout.addWidget(self.y4_label)
        self.pos_group_layout.addWidget(self.y4_edit)
        self.pos_group_layout.addWidget(self.x5_label)
        self.pos_group_layout.addWidget(self.x5_edit)
        self.pos_group_layout.addWidget(self.y5_label)
        self.pos_group_layout.addWidget(self.y5_edit)

        self.button_layout = QtGui.QHBoxLayout()
        self.mainLayout.addLayout(self.button_layout)
        self.refresh_btn = QtGui.QPushButton("Refresh Chrome Count")
        self.do_btn = QtGui.QPushButton("Sunshine!!!")
        self.do_now_btn = QtGui.QPushButton("Do Now")
        self.button_layout.addWidget(self.refresh_btn)
        self.button_layout.addWidget(self.do_btn)
        self.button_layout.addWidget(self.do_now_btn)

        self.refresh_btn.clicked.connect(self.refresh_chrome_count)
        self.do_btn.clicked.connect(self.process)
        self.do_now_btn.clicked.connect(self.process_now)

        self.status_layout = QtGui.QVBoxLayout()
        self.mainLayout.addLayout(self.status_layout)
        self.status_label = QtGui.QLabel("standby.")
        self.status_label.setStyleSheet("QLabel { color : green; }")
        self.status_layout.addWidget(self.status_label)

    def refresh_chrome_count(self):
        self.get_hwnd()
        self.chrome_count_label.setText(str(len(self.hwnd_list)))

    def process(self):
        if not self.running:
            # self.proc_thread = ClickProcessThread(self)
            for i in range(len(self.hwnd_list)):
                cur_hwnd = self.hwnd_list[i]
                thread = PowerUpClickProcessThread(self, cur_hwnd)
                thread.start()
                self.proc_thread_list.append(thread)
            # self.proc_thread.start()
            self.status_label.setText("start processing...")
            self.running = True
        else:
            if len(self.proc_thread_list) > 0:
                for i in range(len(self.proc_thread_list)):
                    self.proc_thread_list[i].exit()
                # self.proc_thread.exit()
                self.status_label.setText("standby")
                self.running = False

    def process_now(self):
        if not self.running:
            # self.proc_thread = ClickProcessThread(self)
            for i in range(len(self.hwnd_list)):
                cur_hwnd = self.hwnd_list[i]
                thread = ClickNowProcessThread(self, cur_hwnd)
                thread.start()

class ClickProcessThread(QtCore.QThread):
    def __init__(self, parent):
        super(ClickProcessThread, self).__init__(parent)
        self.parent = parent
        self.click_hour = int(self.parent.hour_edit.text())
        self.click_min = int(self.parent.min_edit.text())
        self.click_offset = int(self.parent.offset_edit.text())
        self.click_step = int(self.parent.step_edit.text())
        self.get_click_pos()
        self.hwnd_list = self.parent.hwnd_list
        self.target_time = self.get_target_time()

    def run(self):
        count = 0
        exit_flag = False
        while not exit_flag:
            exec_time = self.target_time - datetime.timedelta(milliseconds=self.click_offset)
            curtime = datetime.datetime.now()
            if curtime.hour == exec_time.hour and curtime.min == exec_time.min and curtime.second == exec_time.second and curtime.microsecond >= exec_time.microsecond:
                h1 = self.hwnd_list[count]
                self.do_click(self.click_pos1, h1)
                time.sleep(2)
                self.do_click(self.click_pos2, h1)
                self.do_click(self.click_pos3, h1)
                self.do_click(self.click_pos4, h1)
                time.sleep(0.4)
                self.do_click(self.click_pos4, h1)
                time.sleep(0.4)
                self.do_click(self.click_pos4, h1)
                time.sleep(0.4)
                self.do_click(self.click_pos5, h1)
                print curtime
                self.target_time = datetime.datetime.now() + datetime.timedelta(milliseconds=self.click_step)
                count = count + 1
                if count >= len(self.hwnd_list):
                    exit_flag = True
                    self.parent.running = False
                    self.parent.status_label.setText("complete!")

    def get_target_time(self):
        localtime = datetime.datetime.now()
        target_time = datetime.datetime(localtime.year, localtime.month, localtime.day, self.click_hour, self.click_min, 0, 0)
        return target_time

    def get_click_pos(self):
        x1 = int(self.parent.x1_edit.text())
        y1 = int(self.parent.y1_edit.text())
        x2 = int(self.parent.x2_edit.text())
        y2 = int(self.parent.y2_edit.text())
        x3 = int(self.parent.x3_edit.text())
        y3 = int(self.parent.y3_edit.text())
        x4 = int(self.parent.x4_edit.text())
        y4 = int(self.parent.y4_edit.text())
        x5 = int(self.parent.x5_edit.text())
        y5 = int(self.parent.y5_edit.text())
        self.click_pos1 = x1, y1
        self.click_pos2 = x2, y2
        self.click_pos3 = x3, y3
        self.click_pos4 = x4, y4
        self.click_pos5 = x5, y5

    def do_click(self, click_pos, hwnd):
        client_pos = win32gui.ScreenToClient(hwnd, click_pos)
        tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
        win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, tmp)


class PowerUpClickProcessThread(QtCore.QThread):
    def __init__(self, parent, hwnd):
        super(PowerUpClickProcessThread, self).__init__(parent)
        self.parent = parent
        self.hwnd = hwnd
        self.click_hour = int(self.parent.hour_edit.text())
        self.click_min = int(self.parent.min_edit.text())
        self.click_offset = int(self.parent.offset_edit.text())
        self.click_step = int(self.parent.step_edit.text())
        self.get_click_pos()
        self.target_time = self.get_target_time()

    def run(self):
        exit_flag = False
        while not exit_flag:
            exec_time = self.target_time - datetime.timedelta(milliseconds=self.click_offset)
            curtime = datetime.datetime.now()
            if curtime.hour == exec_time.hour and curtime.min == exec_time.min and curtime.second == exec_time.second and curtime.microsecond >= exec_time.microsecond:
                h1 = self.hwnd
                self.do_click(self.click_pos1, h1)
                time.sleep(3)
                self.do_click(self.click_pos2, h1)
                self.do_click(self.click_pos3, h1)
                self.do_click(self.click_pos4, h1)
                time.sleep(0.4)
                self.do_click(self.click_pos4, h1)
                time.sleep(0.4)
                self.do_click(self.click_pos4, h1)
                time.sleep(0.4)
                self.do_click(self.click_pos5, h1)
                print curtime
                self.target_time = datetime.datetime.now() + datetime.timedelta(milliseconds=self.click_step)
                break

    def get_target_time(self):
        localtime = datetime.datetime.now()
        target_time = datetime.datetime(localtime.year, localtime.month, localtime.day, self.click_hour, self.click_min, 0, 0)
        return target_time

    def get_click_pos(self):
        x1 = int(self.parent.x1_edit.text())
        y1 = int(self.parent.y1_edit.text())
        x2 = int(self.parent.x2_edit.text())
        y2 = int(self.parent.y2_edit.text())
        x3 = int(self.parent.x3_edit.text())
        y3 = int(self.parent.y3_edit.text())
        x4 = int(self.parent.x4_edit.text())
        y4 = int(self.parent.y4_edit.text())
        x5 = int(self.parent.x5_edit.text())
        y5 = int(self.parent.y5_edit.text())
        self.click_pos1 = x1, y1
        self.click_pos2 = x2, y2
        self.click_pos3 = x3, y3
        self.click_pos4 = x4, y4
        self.click_pos5 = x5, y5

    def do_click(self, click_pos, hwnd):
        client_pos = win32gui.ScreenToClient(hwnd, click_pos)
        tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
        win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, tmp)


class ClickNowProcessThread(QtCore.QThread):
    def __init__(self, parent, hwnd):
        super(ClickNowProcessThread, self).__init__(parent)
        self.parent = parent
        self.hwnd = hwnd
        self.get_click_pos()

    def run(self):
        h1 = self.hwnd
        self.do_click(self.click_pos1, h1)
        time.sleep(1.5)
        self.do_click(self.click_pos2, h1)
        self.do_click(self.click_pos3, h1)
        self.do_click(self.click_pos4, h1)
        time.sleep(0.4)
        self.do_click(self.click_pos4, h1)
        time.sleep(0.4)
        self.do_click(self.click_pos4, h1)
        time.sleep(0.4)
        for i in range(5):
            real_pos = self.click_pos5[0], self.click_pos5[1] + i * 25
            self.do_click(real_pos, h1)

    def get_click_pos(self):
        x1 = int(self.parent.x1_edit.text())
        y1 = int(self.parent.y1_edit.text())
        x2 = int(self.parent.x2_edit.text())
        y2 = int(self.parent.y2_edit.text())
        x3 = int(self.parent.x3_edit.text())
        y3 = int(self.parent.y3_edit.text())
        x4 = int(self.parent.x4_edit.text())
        y4 = int(self.parent.y4_edit.text())
        x5 = int(self.parent.x5_edit.text())
        y5 = int(self.parent.y5_edit.text())
        self.click_pos1 = x1, y1
        self.click_pos2 = x2, y2
        self.click_pos3 = x3, y3
        self.click_pos4 = x4, y4
        self.click_pos5 = x5, y5

    def do_click(self, click_pos, hwnd):
        client_pos = win32gui.ScreenToClient(hwnd, click_pos)
        tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
        win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, tmp)

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    chm = ClickHelperMain()
    chm.show()
    sys.exit(app.exec_())
