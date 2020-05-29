#!/usr/bin/python3
#
#  Copyright (C) 2019-2020 Angus King
#
#  displayobject.py - This file is part of catalogue.
#
#  catalogue is free software: you can redistribute it and/or modify
#  it under the terms of the GNU Affero General Public License as
#  published by the Free Software Foundation, either version 3 of
#  the License, or (at your option) any later version.
#
#  catalogue is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU Affero General Public License for more details.
#
#  You should have received a copy of the GNU Affero General
#  Public License along with catalogue.  If not, see
#  <http://www.gnu.org/licenses/>.
#

import os
import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.Qt import *

from functions import *


class GrowingTextEdit(QPlainTextEdit):
# From: https://stackoverflow.com/questions/11851020/a-qwidget-like-qtextedit-that-wraps-its-height-automatically-to-its-contents
    def __init__(self, *args, **kwargs):
        super(GrowingTextEdit, self).__init__(*args, **kwargs)
        self.document().contentsChanged.connect(self.sizeChange)

        self.heightMin = 0
        self.heightMax = 65000

    def sizeChange(self):
        docHeight = self.document().size().height()
    #    print('(27)', self.document().size().height(), self.document().size().width())
        if self.heightMin <= docHeight <= self.heightMax:
            self.setMinimumHeight(docHeight)


class AnObject(QDialog):
    procStart = pyqtSignal(str)

    def resizeEvent(self, event):
   #     print('(36) resize (%d x %d)' % (event.size().width(), event.size().height()))
        QWidget.resizeEvent(self, event)
        w = event.size().width()
        h = event.size().height()

    def resize(self, w, h):
    #    # Pass through to Qt to resize the widget.
        QWidget.resize( self, w, h )

    def __init__(self, dialog, anobject, readonly=True, title=None, section=None,
                 textedit=True, duplicate=None, combolist=None, multi=False, locnlist=None):
        super(AnObject, self).__init__()
        self.anobject = anobject
        self.readonly = readonly
        self.title = title
        self.section = section
        self.textedit = textedit
        self.duplicate = duplicate
        self.combolist = combolist
        self.locnlist = locnlist
        self.multi = multi
        dialog.setObjectName('Dialog')
        self.initUI()

    def set_stuff(self, grid, widths, heights, i):
        if widths[1] > 0:
            grid.setColumnMinimumWidth(0, widths[0] + 10)
            grid.setColumnMinimumWidth(1, widths[1] + 10)
        i += 1
        if isinstance(self.anobject, dict) and self.textedit:
            self.message = QLabel('') #str(heights))
            msg_font = self.message.font()
            msg_font.setBold(True)
            self.message.setFont(msg_font)
            msg_palette = QPalette()
            msg_palette.setColor(QPalette.Foreground, Qt.red)
            self.message.setPalette(msg_palette)
            grid.addWidget(self.message, i + 1, 1)
            i += 1
        if isinstance(self.anobject, str):
            quit = QPushButton('Close', self)
        else:
            quit = QPushButton('Quit', self)
        width = quit.fontMetrics().boundingRect('Close').width() + 10
        quit.setMaximumWidth(width)
        grid.addWidget(quit, i + 1, 0)
        quit.clicked.connect(self.quitClicked)
        if not self.readonly:
            save = QPushButton('Save', self)
            save.setMaximumWidth(width)
            grid.addWidget(save, i + 1, 1)
            save.clicked.connect(self.saveClicked)
        frame = QFrame()
        frame.setLayout(grid)
        frame.setFrameShape(QFrame.NoFrame)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(frame)
        layout = QVBoxLayout(self)
        layout.addWidget(scroll)
        self.setLayout(layout)
        screen = QDesktopWidget().availableGeometry()
        h = heights * i * 2.5
        w = widths[0] + widths[1] + 80
        if w > screen.width():
            w = int(screen.width() * .85)
        if h > screen.height():
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                pct = 0.85
            else:
                pct = 0.90
            h = int(screen.height() * pct)
        if w < 300:
            w = 300
        if h < 160:
            h = 160
        self.resize(w, h)
        if self.title is not None:
            self.setWindowTitle(self.title)
        elif isinstance(self.anobject, str) or isinstance(self.anobject, dict):
            self.setWindowTitle('?')
        else:
            self.setWindowTitle('Review ' + getattr(self.anobject, '__module__'))
        self.setWindowIcon(QIcon('books.png'))

    def initUI(self):
        label = []
        self.edit = []
        self.field_type = []
        metrics = []
        widths = [50, 50]
        heights = 0
        rows = 0
        i = -1
        grid = QGridLayout()
        if isinstance(self.anobject, str):
            self.web = QTextEdit()
            if os.path.exists(self.anobject):
                htf = open(self.anobject, 'r')
                html = htf.read()
                htf.close()
                if self.anobject[-5:].lower() == '.html' or \
                    self.anobject[-4:].lower() == '.htm' or \
                    html[:5] == '<html':
                    try:
                        html = html.replace('[VERSION]', fileVersion())
                    except:
                        pass
                    if self.section is not None:
                        line = html.split('\n')
                        html = ''
                        for i in range(len(line)):
                            html += line[i] + '\n'
                            if line[i].strip() == '<body>':
                               break
                        for i in range(i, len(line)):
                            if line[i][:2] == '<h':
                                if line[i].find('id="' + self.section + '"') > 0:
                                    break
                        for i in range(i, len(line)):
                            if line[i].find('Back to top<') > 0:
                                break
                            j = line[i].find(' (see <a href=')
                            if j > 0:
                                k = line[i].find('</a>)', j)
                                line[i] = line[i][:j] + line[i][k + 5:]
                            html += line[i] + '\n'
                        for i in range(i, len(line)):
                            if line[i].strip() == '</body>':
                                break
                        for i in range(i, len(line)):
                            html += line[i] + '\n'
                    self.web.setHtml(html)
                else:
                    self.web.setPlainText(html)
            else:
                html = self.anobject
                if self.anobject[:5] == '<html':
                    self.anobject = self.anobject.replace('[VERSION]', fileVersion())
                    self.web.setHtml(self.anobject)
                else:
                    self.web.setPlainText(self.anobject)
            metrics.append(self.web.fontMetrics())
            try:
                widths[0] = metrics[0].boundingRect(self.web.text()).width()
                heights = metrics[0].boundingRect(self.web.text()).height()
            except:
                bits = html.split('\n')
                for lin in bits:
                    if len(lin) > widths[0]:
                        widths[0] = len(lin)
                heights = len(bits)
                fnt = self.web.fontMetrics()
                widths[0] = (widths[0]) * fnt.maxWidth()
                heights = (heights) * fnt.height()
                screen = QDesktopWidget().availableGeometry()
                if widths[0] > screen.width() * .67:
                    heights = int(heights / .67)
                    widths[0] = int(screen.width() * .67)
            if self.readonly:
                self.web.setReadOnly(True)
            i = 1
            grid.addWidget(self.web, 0, 0)
            self.set_stuff(grid, widths, heights, i)
        elif isinstance(self.anobject, dict):
            if self.textedit: # probably only this for catalogue app
                self.keys = []
                metrics = [QLabel('text').fontMetrics(), GrowingTextEdit('text').fontMetrics()]
                rows = {}
                heights = 18
                for key, value in self.anobject.items():
                    if value is None:
                        value = ''
                    widths[0] = max(widths[0], metrics[0].boundingRect(key).width())
                    try:
                        bits = value.split('\n')
                        for lin in bits:
                            widths[1] = max(widths[1], metrics[1].boundingRect(lin).width())
                        if self.combolist is not None and key == self.combolist[0]:
                            max_val = ''
                            for val in self.combolist[1]:
                                if (len(val) + 4) > len(max_val):
                                    max_val = val + 'xxxx'
                            widths[1] = max(widths[1], metrics[1].boundingRect(max_val).width())
                        rows[key] = [max(metrics[0].boundingRect(key).height(), metrics[1].boundingRect(value).height()),
                                     len(bits)]
                        heights = max(heights, rows[key][0])
                    except:
                        rows[key] = [metrics[0].boundingRect(key).height(), 1]
                for key, value in self.anobject.items():
                    if value is None:
                        value = ''
                    self.field_type.append('str')
                    label.append(QLabel(key + ':'))
                    self.keys.append(key)
                    if self.duplicate is not None and len(self.keys) == 1:
                        self.edit.append(QLineEdit())
                        self.edit[-1].resize(widths[1], rows[key][0])
                        self.edit[-1].setText(value)
                    else:
                        if self.combolist is not None and key == self.combolist[0]:
                            label[-1] = QLabel('<strong>' + key + ':</strong>')
                            if self.multi:
                                self.metacombo = ClickableQLabel()
                                frameStyle = QFrame.Sunken | QFrame.Panel
                                self.metacombo.setFrameStyle(frameStyle)
                                self.metacombo.setStyleSheet("background-color:#ffffff;")
                                self.chosen = value
                                txt = ''
                                for valu in value:
                                    txt += ';' + valu
                                txt = txt[1:]
                                self.metacombo.setText(txt)
                                self.metacombo.clicked.connect(self.comboSelected)
                                self.edit.append(self.metacombo)
                            else:
                                self.metacombo = QComboBox(self)
                                j = 0
                                for val in self.combolist[1]:
                                    if value == val:
                                        j = self.metacombo.count()
                                    self.metacombo.addItem(val)
                                self.edit.append(self.metacombo)
                                self.metacombo.setCurrentIndex(j)
                        else:
                            if self.locnlist is not None and key == self.locnlist[0]:
                                self.locncombo = QComboBox(self)
                                j = 0
                                for val in self.locnlist[1]:
                                    if value == val:
                                        j = self.locncombo.count()
                                    self.locncombo.addItem(val)
                                self.edit.append(self.locncombo)
                                self.locncombo.setCurrentIndex(j)
                                self.locncombo.setEditable(True)
                            else:
        #                    self.edit.append(QTextEdit())
                                self.edit.append(GrowingTextEdit())
                                self.edit[-1].resize(widths[1], rows[key][0])
                                self.edit[-1].setPlainText(value)
                    if self.readonly:
                        self.edit[-1].setReadOnly(True)
                    i += 1
                    grid.addWidget(label[-1], i + 1, 0)
                    grid.addWidget(self.edit[-1], i + 1, 1)
                    grid.setRowMinimumHeight(grid.rowCount() - 1, rows[key][0])
                    grid.setRowStretch(grid.rowCount() - 1, rows[key][1])
                self.edit[0].setFocusPolicy(Qt.StrongFocus)
            else:
                print('(226) Not been here before')
                self.keys = []
                for key, value in self.anobject.items():
                    if value is None:
                        value = ''
                    self.field_type.append('str')
                    label.append(QLabel(key + ':'))
                    self.keys.append(key)
                    self.edit.append(QLineEdit())
                    self.edit[-1].setText(value)
                    if i < 0:
                        metrics.append(label[-1].fontMetrics())
                        metrics.append(self.edit[-1].fontMetrics())
                    chars = (len(value) + 5) * metrics[0].maxWidth()
                    rows = metrics[0].height()
                    self.edit[-1].resize(chars, rows)
                    if metrics[0].boundingRect(label[-1].text()).width() > widths[0]:
                        widths[0] = metrics[0].boundingRect(label[-1].text()).width()
                    try:
                        if metrics[1].boundingRect(self.edit[-1].text()).width() > widths[1]:
                            widths[1] = metrics[1].boundingRect(self.edit[-1].text()).width()
                    except:
                        widths[1] = chars
                    if self.readonly:
                        self.edit[-1].setReadOnly(True)
                    i += 1
                    grid.addWidget(label[-1], i, 0)
                    grid.addWidget(self.edit[-1], i, 1)
            self.set_stuff(grid, widths, heights, i)
        else:
            print('(256) Not been here before')
            for prop in dir(self.anobject):
                if prop[:2] != '__' and prop[-2:] != '__':
                    attr = getattr(self.anobject, prop)
                    if isinstance(attr, int):
                         self.field_type.append('int')
                    elif isinstance(attr, float):
                         self.field_type.append('float')
                    else:
                         self.field_type.append('str')
                    label.append(QLabel(prop.title() + ':'))
                    if self.field_type[-1] != 'str':
                        self.edit.append(QLineEdit(str(attr)))
                    else:
                        self.edit.append(QLineEdit(attr))
                    if i < 0:
                        metrics.append(label[-1].fontMetrics())
                        metrics.append(self.edit[-1].fontMetrics())
                    if metrics[0].boundingRect(label[-1].text()).width() > widths[0]:
                        widths[0] = metrics[0].boundingRect(label[-1].text()).width()
                    if metrics[1].boundingRect(self.edit[-1].text()).width() > widths[1]:
                        widths[1] = metrics[1].boundingRect(self.edit[-1].text()).width()
                    for j in range(2):
                        if metrics[j].boundingRect(label[-1].text()).height() > heights:
                            heights = metrics[j].boundingRect(label[-1].text()).height()
                    if self.readonly:
                        self.edit[-1].setReadOnly(True)
                    i += 1
                    grid.addWidget(label[-1], i + 1, 0)
                    grid.addWidget(self.edit[-1], i + 1, 1)
            self.set_stuff(grid, widths, heights, i)
        QShortcut(QKeySequence('q'), self, self.quitClicked)

    def quitClicked(self):
        self.anobject = None
        self.close()

    def comboSelected(self):
        chosen = selectMulti(self.combolist[1], self.chosen, self.combolist[0])
        chosen.setWindowModality(Qt.WindowModal)
        chosen.setWindowFlags(Qt.WindowStaysOnTopHint)
        chosen.exec_()
        selected = chosen.getValues()
        del chosen
        if selected is None:
            return
        if len(selected) == 0:
            txt = '?'
            self.chosen = [txt]
        else:
            self.chosen = selected
            txt = ''
            for valu in self.chosen:
                txt += ';' + valu
            txt = txt[1:]
        self.metacombo.setText(txt)

    def saveClicked(self):
        if isinstance(self.anobject, dict):
            if self.textedit:
                for i in range(len(self.keys)):
                    if self.combolist is not None and self.keys[i] == self.combolist[0]:
                        if self.multi:
                            self.anobject[self.keys[i]] = self.chosen
                        else:
                            self.anobject[self.keys[i]] = self.metacombo.currentText()
                    elif self.locnlist is not None and self.keys[i] == self.locnlist[0]:
                        self.anobject[self.keys[i]] = self.locncombo.currentText()
                    else:
                        try:
                            self.anobject[self.keys[i]] = str(self.edit[i].toPlainText())
                        except:
                            self.anobject[self.keys[i]] = str(self.edit[i].text())
                if self.duplicate is not None:
                    for dup in self.duplicate:
                        if self.anobject[self.keys[0]] == dup:
                            self.message.setText('Duplicate value not allowed for ' + self.keys[0])
                            return
            else:
                for i in range(len(self.keys)):
                    self.anobject[self.keys[i]] = str(self.edit[i].text())
        else:
            i = -1
            for prop in dir(self.anobject):
                if prop[:2] != '__' and prop[-2:] != '__':
                    i += 1
                    if self.field_type[i] == 'int':
                        setattr(self.anobject, prop, int(self.edit[i].text()))
                    elif self.field_type[i] == 'float':
                        setattr(self.anobject, prop, float(self.edit[i].text()))
                    else:
                        setattr(self.anobject, prop, str(self.edit[i].text()))
        self.close()

    def getValues(self):
        return self.anobject


class ClickableQLabel(QLabel):
    clicked=pyqtSignal()
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QApplication.widgetAt(event.globalPos()).setFocus()
        self.clicked.emit()


class selectMulti(QDialog):
    def __init__(self, combos, chosen, multi):
        super(selectMulti, self).__init__()
        self.combos = combos
        self.chosen = chosen
        self.grid = QGridLayout()
#       use this if you want a checkbox
#        j = 0
#        self.checkbox = []
#        self.checkbox.append(QCheckBox('Check / Uncheck all', self))
#        self.grid.addWidget(self.checkbox[-1], 0, 0)
#        i = 0
#        c = 0
#        for combo in sorted(self.combos):
#            self.checkbox.append(QCheckBox(combo, self))
#            if combo in chosen:
#                self.checkbox[-1].setCheckState(Checked)
#            i += 1
#            self.grid.addWidget(self.checkbox[-1], i, c)
#            if i > 25:
#                i = 0
#                c += 1
#        self.checkbox[0].stateChanged.connect(self.check_all)
#        selectc = QPushButton('Select', self)
#        selectc.clicked.connect(self.selectcClicked)
#        self.grid.addWidget(select, i + 1, c)
        self.multiListBox = QListWidget()
     #   self.multiListBox.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.multiListBox.setSelectionMode(QAbstractItemView.MultiSelection)
        i = 0
        for combo in sorted(self.combos):
            self.multiListBox.addItem(combo)
            if combo in chosen:
             #   self.multiListBox.setItemSelected(self.multiListBox.item(i), True)
                self.multiListBox.item(i).setSelected(True)
            i += 1
        j = int((self.multiListBox.fontMetrics().height() + 2.5) * i - self.multiListBox.sizeHint().height())
        self.multiListBox.resize(self.multiListBox.sizeHint().width(), self.multiListBox.fontMetrics().height() * i)
        self.grid.addWidget(self.multiListBox, 0, 0)
        selectl = QPushButton('Select', self)
        selectl.clicked.connect(self.selectlClicked)
        self.grid.addWidget(selectl, 1, 0)
        self.setLayout(self.grid)
        if j > 0:
            h1 = QDesktopWidget().availableGeometry().height() * 0.85
            h = self.sizeHint().height() + j
            if h > h1:
                h = h1
            self.resize(self.sizeHint().width(), h)
        self.setWindowTitle('Choose ' + multi + ' values')
        QShortcut(QKeySequence('q'), self, self.quitClicked)
      #  self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.setWindowState(self.windowState() & Qt.WindowActive)
        self.activateWindow()
        self.show_them = False
        self.setGeometry(100, 100, self.width(), self.height())
        self.show()

    def check_all(self):
        if self.checkbox[0].isChecked():
            for i in range(len(self.checkbox)):
                self.checkbox[i].setCheckState(Checked)
        else:
            for i in range(len(self.checkbox)):
                self.checkbox[i].setCheckState(Unchecked)

    def closeEvent(self, event):
        if not self.show_them:
            self.chosen = None
        event.accept()

    def quitClicked(self):
        self.close()

#    def selectcClicked(self):
#        self.chosen = []
#        for combo in range(1, len(self.checkbox)):
#            if self.checkbox[combo].checkState() == Checked:
#                self.chosen.append(str(self.checkbox[combo].text()))
#        self.show_them = True
#        self.close()

    def selectlClicked(self):
        chosen = []
        for item in self.multiListBox.selectedItems():
            chosen.append(item.text())
        self.chosen = sorted(chosen)
        self.show_them = True
        self.close()

    def getValues(self):
        return self.chosen
