"""
https://github.com/Return-Log/Mail-to-Excel-Converter
AGPL-3.0 license
coding: UTF-8
"""

import email
import imaplib
import re
import sys
from datetime import datetime, timedelta
from email.header import decode_header
import html2text
import pandas as pd
from PyQt5 import QtCore
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QLineEdit, QComboBox,
                             QDateTimeEdit, QPushButton, QVBoxLayout, QWidget, QProgressBar,
                             QFileDialog, QTextEdit, QHBoxLayout)


# 邮件获取线程类，继承自QThread
# 邮件获取线程类，继承自QThread
class EmailFetchThread(QThread):
    update_progress = pyqtSignal(int)  # 更新进度条的信号
    update_debug = pyqtSignal(str)     # 更新调试日志的信号
    finished = pyqtSignal(pd.DataFrame)  # 获取完成的信号

    # 初始化线程
    def __init__(self, email_address, password, imap_server, filter_type, filter_value, since_date, end_date, language, parent=None):
        super().__init__(parent)
        self.email_address = email_address
        self.password = password
        self.imap_server = imap_server
        self.filter_type = filter_type
        self.filter_value = filter_value
        self.since_date = since_date
        self.end_date = end_date
        self.language = language

    # 线程运行的代码
    def run(self):
        try:
            self.update_debug.emit(
                f"Connecting to {self.imap_server}..." if self.language == 'English' else f"正在连接到 {self.imap_server}...")
            mail = imaplib.IMAP4_SSL(self.imap_server)
            mail.login(self.email_address, self.password)
            self.update_debug.emit("Login successful." if self.language == 'English' else "登录成功。")
            mail.select("inbox")

            search_criteria = self.build_search_criteria(self.since_date, self.end_date)
            self.update_debug.emit(
                f"Searching emails with criteria: {search_criteria}" if self.language == 'English' else f"使用以下条件搜索邮件：{search_criteria}")

            result, data = mail.search(None, search_criteria)
            if result != 'OK':
                self.update_debug.emit(
                    f"Error searching emails: {data}" if self.language == 'English' else f"搜索邮件出错：{data}")
                return

            email_ids = data[0].split()
            self.update_debug.emit(
                f"Found {len(email_ids)} emails." if self.language == 'English' else f"找到 {len(email_ids)} 封邮件。")
            emails = []

            for i, email_id in enumerate(email_ids):
                result, msg_data = mail.fetch(email_id, "(RFC822)")
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                date = msg["Date"]
                sender = self.decode_header_value(msg["From"])
                subject = self.decode_header_value(msg["Subject"])

                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        if content_type == "text/plain":
                            body = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
                            break
                        elif content_type == "text/html":
                            html_content = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
                            body = html2text.html2text(html_content)
                            break
                else:
                    content_type = msg.get_content_type()
                    if content_type == "text/plain":
                        body = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8')
                    elif content_type == "text/html":
                        html_content = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8')
                        body = html2text.html2text(html_content)

                # 本地过滤
                if not self.is_ascii(self.filter_value):
                    if self.filter_type == "From Sender" or self.filter_type == "按发件人" and self.filter_value not in sender:
                        continue
                    if self.filter_type == "By Keyword" or self.filter_type == "按关键字" and self.filter_value not in subject:
                        continue

                emails.append([date, sender, subject, body])

                self.update_progress.emit(int((i + 1) / len(email_ids) * 100))

            df = pd.DataFrame(emails, columns=["Date", "Sender", "Subject", "Content"])
            self.finished.emit(df)
            mail.logout()

        except Exception as e:
            self.update_debug.emit(str(e))

    # 解码邮件头信息
    def decode_header_value(self, value):
        decoded_bytes, charset = decode_header(value)[0]
        if isinstance(decoded_bytes, bytes):
            return decoded_bytes.decode(charset or 'utf-8')
        return decoded_bytes

    # 构建搜索条件
    def build_search_criteria(self, since_date, end_date):
        criteria = [f'SINCE {since_date.strftime("%d-%b-%Y")}', f'BEFORE {end_date.strftime("%d-%b-%Y")}']
        if self.is_ascii(self.filter_value):
            if self.filter_type == "From Sender" or self.filter_type == "按发件人":
                criteria.append(f'FROM "{self.filter_value}"')
            elif self.filter_type == "By Keyword" or self.filter_type == "按关键字":
                criteria.append(f'SUBJECT "{self.filter_value}"')
        return ' '.join(criteria)

    # 检查字符串是否为ASCII
    def is_ascii(self, s):
        return all(ord(c) < 128 for c in s)



# 主应用类，继承自QMainWindow
class MailToExcelApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mail to Excel Converter")
        self.setGeometry(100, 100, 800, 600)

        self.language = 'English'  # 默认语言

        self.initUI()

    # 初始化UI
    def initUI(self):
        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)

        layout = QVBoxLayout()

        self.emailLabel = QLabel("Email Address:" if self.language == 'English' else "邮箱地址：")
        self.emailInput = QLineEdit()
        self.passwordLabel = QLabel("Password:" if self.language == 'English' else "密码：")
        self.passwordInput = QLineEdit()
        self.passwordInput.setEchoMode(QLineEdit.Password)
        self.imapLabel = QLabel("IMAP Server:" if self.language == 'English' else "IMAP服务器：")
        self.imapInput = QLineEdit()
        self.imapInput.setText("outlook.office365.com")

        self.filterTypeLabel = QLabel("Filter Type:" if self.language == 'English' else "过滤类型：")
        self.filterTypeCombo = QComboBox()
        self.filterTypeCombo.addItems(["All Emails", "From Sender", "By Keyword"] if self.language == 'English' else ["所有邮件", "按发件人", "按关键字"])
        self.filterTypeCombo.currentIndexChanged.connect(self.onFilterTypeChange)
        self.filterValueInput = QLineEdit()
        self.filterValueInput.setPlaceholderText("Enter sender email or keyword" if self.language == 'English' else "输入发件人邮箱或关键字")
        self.filterValueInput.setEnabled(False)

        self.timeRangeLabel = QLabel("Time Range:" if self.language == 'English' else "时间范围：")
        self.timeRangeCombo = QComboBox()
        self.timeRangeCombo.addItems(["Last Week", "Last Month", "Last Year", "All Time", "Custom"] if self.language == 'English' else ["上周", "上个月", "去年", "所有时间", "自定义"])
        self.timeRangeCombo.currentIndexChanged.connect(self.onTimeRangeChange)

        self.startDateLabel = QLabel("Start Date:" if self.language == 'English' else "开始日期：")
        self.startDateEdit = QDateTimeEdit()
        self.startDateEdit.setDateTime(QtCore.QDateTime.currentDateTime().addDays(-7))
        self.startDateEdit.setEnabled(False)

        self.endDateLabel = QLabel("End Date:" if self.language == 'English' else "结束日期：")
        self.endDateEdit = QDateTimeEdit()
        self.endDateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.endDateEdit.setEnabled(False)

        self.exportButton = QPushButton("Export to Excel" if self.language == 'English' else "导出到Excel")
        self.exportButton.clicked.connect(self.exportToExcel)

        self.progressBar = QProgressBar()

        self.debugOutput = QTextEdit()
        self.debugOutput.setReadOnly(True)
        self.debugOutput.setFixedHeight(100)

        layout.addWidget(self.emailLabel)
        layout.addWidget(self.emailInput)
        layout.addWidget(self.passwordLabel)
        layout.addWidget(self.passwordInput)
        layout.addWidget(self.imapLabel)
        layout.addWidget(self.imapInput)
        layout.addWidget(self.filterTypeLabel)
        layout.addWidget(self.filterTypeCombo)
        layout.addWidget(self.filterValueInput)
        layout.addWidget(self.timeRangeLabel)
        layout.addWidget(self.timeRangeCombo)
        layout.addWidget(self.startDateLabel)
        layout.addWidget(self.startDateEdit)
        layout.addWidget(self.endDateLabel)
        layout.addWidget(self.endDateEdit)
        layout.addWidget(self.exportButton)
        layout.addWidget(self.progressBar)
        layout.addWidget(self.debugOutput)

        self.infoLayout = QHBoxLayout()
        self.infoLabel = QLabel('<a href="https://github.com/Return-Log/Mail-to-Excel-Converter">Mail to Excel '
                                'Converter v1.1 | Copyright © 2024 Return-Log</a>' if self.language == 'English' else '<a'
                                                                                              'href="https://github'
                                                                                              '.com/Return-Log/Mail'
                                                                                              '-to-Excel-Converter'
                                                                                              '">邮件到Excel转换器 v1.1 | '
                                                                                              'Copyright © 2024 Return-Log</a>')
        self.infoLabel.setOpenExternalLinks(True)
        self.languageComboBox = QComboBox()
        self.languageComboBox.addItems(["English", "中文"])
        self.languageComboBox.currentIndexChanged.connect(self.changeLanguage)
        self.infoLayout.addWidget(self.infoLabel)
        self.infoLayout.addWidget(self.languageComboBox)
        layout.addLayout(self.infoLayout)

        self.centralWidget.setLayout(layout)

    # 过滤类型改变事件处理
    def onFilterTypeChange(self, index):
        self.filterValueInput.setEnabled(index != 0)

    # 时间范围改变事件处理
    def onTimeRangeChange(self, index):
        self.startDateEdit.setEnabled(index == 4)
        self.endDateEdit.setEnabled(index == 4)

    # 切换语言
    def changeLanguage(self, index):
        self.language = self.languageComboBox.currentText()
        self.updateUIText()

    # 更新UI文本
    def updateUIText(self):
        self.emailLabel.setText("Email Address:" if self.language == 'English' else "邮箱地址：")
        self.passwordLabel.setText("Password:" if self.language == 'English' else "密码：")
        self.imapLabel.setText("IMAP Server:" if self.language == 'English' else "IMAP服务器：")
        self.filterTypeLabel.setText("Filter Type:" if self.language == 'English' else "过滤类型：")
        self.filterTypeCombo.setItemText(0, "All Emails" if self.language == 'English' else "所有邮件")
        self.filterTypeCombo.setItemText(1, "From Sender" if self.language == 'English' else "按发件人")
        self.filterTypeCombo.setItemText(2, "By Keyword" if self.language == 'English' else "按关键字")
        self.filterValueInput.setPlaceholderText("Enter sender email or keyword" if self.language == 'English' else "输入发件人邮箱或关键字")
        self.timeRangeLabel.setText("Time Range:" if self.language == 'English' else "时间范围：")
        self.timeRangeCombo.setItemText(0, "Last Week" if self.language == 'English' else "上周")
        self.timeRangeCombo.setItemText(1, "Last Month" if self.language == 'English' else "上个月")
        self.timeRangeCombo.setItemText(2, "Last Year" if self.language == 'English' else "去年")
        self.timeRangeCombo.setItemText(3, "All Time" if self.language == 'English' else "所有时间")
        self.timeRangeCombo.setItemText(4, "Custom" if self.language == 'English' else "自定义")
        self.startDateLabel.setText("Start Date:" if self.language == 'English' else "开始日期：")
        self.endDateLabel.setText("End Date:" if self.language == 'English' else "结束日期：")
        self.exportButton.setText("Export to Excel" if self.language == 'English' else "导出到Excel")
        self.infoLabel.setText('<a href="https://github.com/Return-Log/Mail-to-Excel-Converter">Mail to Excel '
                               'Converter v1.1 | Copyright © 2024 Return-Log</a>' if self.language == 'English' else '<a '
                                                                                             'href="https://github'
                                                                                             '.com/Return-Log/Mail-to'
                                                                                             '-Excel-Converter'
                                                                                             '">邮件到Excel转换器 v1.1 | '
                                                                                             'Copyright © 2024 Return-Log</a>')
        self.infoLabel.setOpenExternalLinks(True)

    # 记录调试信息
    def logDebug(self, message):
        self.debugOutput.append(message)

    # 验证邮箱格式
    def validateEmail(self, email):
        email_regex = r'^\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        if not re.match(email_regex, email):
            raise ValueError("Invalid email address format" if self.language == 'English' else "邮箱地址格式无效")

    # 导出到Excel
    def exportToExcel(self):
        email_address = self.emailInput.text()
        password = self.passwordInput.text()
        imap_server = self.imapInput.text()
        filter_type = self.filterTypeCombo.currentText()
        filter_value = self.filterValueInput.text()
        time_range = self.timeRangeCombo.currentText()

        try:
            if filter_type == "From Sender" or filter_type == "按发件人":
                self.validateEmail(filter_value)

            if time_range == "Last Week" or time_range == "上周":
                since_date = datetime.now() - timedelta(weeks=1)
            elif time_range == "Last Month" or time_range == "上个月":
                since_date = datetime.now() - timedelta(days=30)
            elif time_range == "Last Year" or time_range == "去年":
                since_date = datetime.now() - timedelta(days=365)
            elif time_range == "All Time" or time_range == "所有时间":
                since_date = datetime(1970, 1, 1)
            elif time_range == "Custom" or time_range == "自定义":
                since_date = self.startDateEdit.dateTime().toPyDateTime()
            end_date = self.endDateEdit.dateTime().toPyDateTime()

            self.email_fetch_thread = EmailFetchThread(email_address, password, imap_server, filter_type, filter_value, since_date, end_date, self.language)
            self.email_fetch_thread.update_progress.connect(self.progressBar.setValue)
            self.email_fetch_thread.update_debug.connect(self.logDebug)
            self.email_fetch_thread.finished.connect(self.saveToExcel)
            self.email_fetch_thread.start()

        except ValueError as ve:
            self.logDebug(str(ve))

    # 保存到Excel
    def saveToExcel(self, df):
        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx)")
        if save_path:
            df.to_excel(save_path, index=False)
            self.logDebug(f"Exported to {save_path}" if self.language == 'English' else f"导出到 {save_path}")

        self.progressBar.setValue(100)
        self.logDebug("Export completed." if self.language == 'English' else "导出完成。")
        self.progressBar.setValue(0)


# 主函数，运行应用
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MailToExcelApp()
    window.show()
    sys.exit(app.exec_())
