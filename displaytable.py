#!/usr/bin/python3
#
#  Copyright (C) 2019 Angus King
#
#  displaytable.py - This file is part of catalogue.
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
from PyQt4 import QtCore
from PyQt4 import QtGui

import displayobject

class FakeObject:
    def __init__(self, fake_object, fields):
        f = -1
        if not isinstance(fake_object, list) and len(fields) > 1:
            f += 1
            setattr(self, fields[f], fake_object)
            for f in range(1, len(fields)):
                setattr(self, fields[f], '')
            return
        for i in range(len(fake_object)):
            if isinstance(fake_object[i], list):
                for j in range(len(fake_object[i])):
                    f += 1
                    setattr(self, fields[f], fake_object[i][j])
            else:
                f += 1
                if fake_object[i] is None:
                    fake_object[i] = ''
                setattr(self, fields[f], fake_object[i])


class Table(QtGui.QDialog):
    def __init__(self, objects, parent=None, fields=None, title=None, edit=False, decpts=None, sortby=None):
        super(Table, self).__init__(parent)
        if isinstance(objects, list):
            self.ois = 'list'
            fakes = []
            if len(objects) == 0 or not isinstance(objects[0], list):
                for row in objects:
                    fakes.append(FakeObject([row], fields))
            else:
                for row in objects:
                    fakes.append(FakeObject(row, fields))
            self.objects = fakes
        elif len(objects) == 0:
            buttonLayout = QtGui.QVBoxLayout()
            buttonLayout.addWidget(QtGui.QLabel('Nothing to display.'))
            self.quitButton = QtGui.QPushButton(self.tr('&Quit'))
            buttonLayout.addWidget(self.quitButton)
            self.connect(self.quitButton, QtCore.SIGNAL('clicked()'),
                        self.quit)
            QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
            self.setLayout(buttonLayout)
            return
        elif isinstance(objects, tuple):
            self.ois = 'list'
            fakes = []
            for row in objects:
                fakes.append(FakeObject(row, fields))
            self.objects = fakes
        elif isinstance(objects, dict):
            self.ois = 'dict'
            fakes = []
            if fields is None: # assume we have some class objects
                fields = []
                if hasattr(objects[list(objects.keys())[0]], 'name'):
                    fields.append('name')
                for prop in dir(objects[list(objects.keys())[0]]):
                    if prop[:2] != '__' and prop[-2:] != '__':
                        if prop != 'name':
                            fields.append(prop)
                for key, value in objects.items():
                     values = []
                     for field in fields:
                         values.append(getattr(value, field))
                     fakes.append(FakeObject(values, fields))
            else:
                for key, value in objects.items():
                    fakes.append(FakeObject([key, value], fields))
            self.objects = fakes
        else:
            self.ois = 'obj'
            self.objects = objects
        self.selection = None
        self.recur = False
        self.fields = fields
        self.title = title
        self.edit_table = edit
        self.decpts = decpts
        self.replaced = None
        if self.edit_table:
            self.title_word = ['Edit', 'Export']
        else:
            self.title_word = ['Display', 'Save']
        if self.title is None:
            try:
                self.setWindowTitle(self.title_word[0] + ' ' + self.fields[0] + ' values')
            except:
                self.setWindowTitle(self.title_word[0] + ' values')
        else:
            self.setWindowTitle(self.title)
        self.setWindowIcon(QtGui.QIcon('books.png'))
        msg = '(Right click column header to sort)'
        buttonLayout = QtGui.QHBoxLayout()
        self.message = QtGui.QLabel(msg)
        self.quitButton = QtGui.QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(self.quitButton)
        self.connect(self.quitButton, QtCore.SIGNAL('clicked()'),
                    self.quit)
        if self.edit_table:
            self.addButton = QtGui.QPushButton(self.tr('Add'))
            buttonLayout.addWidget(self.addButton)
            self.connect(self.addButton, QtCore.SIGNAL('clicked()'),
                         self.addtotbl)
            self.replaceButton = QtGui.QPushButton(self.tr('Save'))
            buttonLayout.addWidget(self.replaceButton)
            self.connect(self.replaceButton, QtCore.SIGNAL('clicked()'),
                        self.replacetbl)
        self.saveButton = QtGui.QPushButton(self.tr(self.title_word[1]))
        buttonLayout.addWidget(self.saveButton)
        self.connect(self.saveButton, QtCore.SIGNAL('clicked()'),
                    self.saveit)
        buttons = QtGui.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtGui.QGridLayout()
        self.table = QtGui.QTableWidget()
        self.populate()
        self.table.setRowCount(len(self.entry))
        self.table.setColumnCount(len(self.labels))
        if self.fields is None:
            labels = sorted(self.labels.keys())
        else:
            labels = self.fields
        self.table.setHorizontalHeaderLabels(labels)
        for cl in range(self.table.columnCount()):
            self.table.horizontalHeaderItem(cl).setIcon(QtGui.QIcon('blank.png'))
        self.headers = self.table.horizontalHeader()
        self.headers.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.headers.customContextMenuRequested.connect(self.header_click)
        self.headers.setSelectionMode(QtGui.QAbstractItemView.SingleSelection)
        self.rh = self.table.verticalHeader().sizeHint().height()
        if self.edit_table:
            self.table.setEditTriggers(QtGui.QAbstractItemView.CurrentChanged)
            self.rows = self.table.verticalHeader()
            self.rows.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            self.rows.customContextMenuRequested.connect(self.row_click)
            self.rows.setSelectionMode(QtGui.QAbstractItemView.SingleSelection)
            self.table.verticalHeader().setVisible(True)
        else:
            self.table.setEditTriggers(QtGui.QAbstractItemView.SelectedClicked)
            self.table.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        layout.addWidget(self.table, 0, 0)
        layout.addWidget(self.message, 1, 0)
        layout.addWidget(buttons, 2, 0)
        self.sort_col = 0
        self.sort_asc = False
        if sortby is None:
            xtra_rows = self.order(0)
        elif sortby == '':
            xtra_rows = self.order(-1)
        else:
            xtra_rows = self.order(self.fields.index(sortby))
        self.table.resizeColumnsToContents()
        width = 0
        for cl in range(self.table.columnCount()):
            width += self.table.columnWidth(cl)
        width += 50
        height = self.rh * (self.table.rowCount() + xtra_rows) + 200
        screen = QtGui.QDesktopWidget().availableGeometry()
        if height > (screen.height() - 70):
            height = screen.height() - 70
        self.setLayout(layout)
        size = QtCore.QSize(int(width), int(height))
        self.resize(size)
        self.updated = QtCore.pyqtSignal(QtGui.QLabel)   # ??
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        if self.edit_table:
            self.table.cellChanged.connect(self.item_changed)
        else:
            self.table.cellClicked.connect(self.item_selected)

    def populate(self):
        self.labels = {}
        self.lens = {}
        try:
            self.hdrs
        except:
            self.hdrs = {}
        if len(self.objects) == 0:
            self.entry = []
            for field in self.fields:
                self.labels[field] = 'str'
            return
        for thing in self.objects:
            for prop in dir(thing):
                if prop[:2] != '__' and prop[-2:] != '__':
                    if self.fields is not None:
                        if prop not in self.fields:
                            continue
                    attr = getattr(thing, prop)
                    if attr is None:
                        continue
                    if prop not in self.labels:
                        if isinstance(attr, int):
                            self.labels[prop] = 'int'
                        elif isinstance(attr, float):
                            self.labels[prop] = 'float'
                        else:
                            self.labels[prop] = 'str'
                    if isinstance(attr, int):
                        if self.labels[prop] == 'str':
                            self.labels[prop] = 'int'
                        if prop in self.lens:
                            if len(str(attr)) > self.lens[prop][0]:
                                self.lens[prop][0] = len(str(attr))
                        else:
                            self.lens[prop] = [len(str(attr)), 0]
                    elif isinstance(attr, float):
                        if self.labels[prop] == 'str':
                            self.labels[prop] = 'float'
                        a = str(attr)
                        bits = a.split('.')
                        if self.decpts is None:
                            if prop in self.lens:
                                for i in range(2):
                                    if len(bits[i]) > self.lens[prop][i]:
                                        self.lens[prop][i] = len(bits[i])
                            else:
                                self.lens[prop] = [len(bits[0]), len(bits[1])]
                        else:
                            pts = self.decpts[self.fields.index(prop)]
                            if prop in self.lens:
                                if len(bits[0]) > self.lens[prop][0]:
                                    self.lens[prop][0] = len(bits[0])
                                if len(bits[1]) > self.lens[prop][1]:
                                    if len(bits[1]) > pts:
                                        self.lens[prop][1] = pts
                                    else:
                                        self.lens[prop][1] = len(bits[1])
                            else:
                                if len(bits[1]) > pts or self.edit_table:
                                    self.lens[prop] = [len(bits[0]), pts]
                                else:
                                    self.lens[prop] = [len(bits[0]), len(bits[1])]
        if self.fields is None:
            self.fields = []
            for key, value in iter(sorted(self.labels.items())):
                self.fields.append(key)
        self.entry = []
        try:
            iam = getattr(self.objects[0], '__module__')
        except:
            iam = 'list'
        for obj in range(len(self.objects)):
            values = {}
            for key, value in self.labels.items():
                try:
                    if key == '#':
                        attr = obj
                    elif key[0] == '%':
                        attr = ''
                    else:
                        attr = getattr(self.objects[obj], key)
                    if value != 'str':
                        fmat_str = '{: >' + str(self.lens[key][0] + self.lens[key][1] + 1) + ',.' + str(self.lens[key][1]) + 'f}'
                        values[key] = fmat_str.format(attr)
                    else:
                        values[key] = attr
                except:
                    pass
            self.entry.append(values)
        return len(self.objects)

    def order(self, col):
        if len(self.entry) == 0:
            return 0
        xtra_rows = 0
        if col < 0:
            torder = []
            for rw in range(len(self.entry)):
                torder.append(rw)
        else:
            numbrs = True
            orderd = {}
            norderd = {}   # minus
            key = self.fields[col]
            max_l = 0
            for rw in range(len(self.entry)):
                if key in self.entry[rw]:
                    try:
                        max_l = max(max_l, len(self.entry[rw][key]))
                        try:
                            txt = str(self.entry[rw][key]).strip().replace(',', '')
                            nmbr = float(txt)
                        except:
                            numbrs = False
                    except:
                        pass
            if numbrs:
                try:
                    fmat_str = '{:0>' + str(self.lens[key][0] + self.lens[key][1] + 1) + '.' + str(self.lens[key][1]) + 'f}'
                except:
                    numbrs = False
            for rw in range(len(self.entry)):
                if numbrs:
                    if key in self.entry[rw]:
                        txt = str(self.entry[rw][key]).strip().replace(',', '')
                        nmbr = float(txt)
                        if nmbr == 0:
                            orderd['<' + fmat_str.format(nmbr) + self.entry[rw][self.fields[0]]] = rw
                        elif nmbr < 0:
                            norderd['<' + fmat_str.format(-nmbr) + self.entry[rw][self.fields[0]]] = rw
                        else:
                            orderd['>' + fmat_str.format(nmbr) + self.entry[rw][self.fields[0]]] = rw
                    else:
                        orderd[' ' + self.entry[rw][self.fields[0]]] = rw
                else:
                    try:
                        orderd[str(self.entry[rw][key]) + self.entry[rw][self.fields[0]]] = rw
                    except:
                        orderd[' ' + self.entry[rw][self.fields[0]]] = rw
            torder = []
            if col != self.sort_col:
                self.table.horizontalHeaderItem(self.sort_col).setIcon(QtGui.QIcon('blank.png'))
                self.sort_asc = False
            self.sort_col = col
            if self.sort_asc:   # swap order
                for key, value in iter(sorted(iter(orderd.items()), reverse=True)):
                    torder.append(value)
                for key, value in iter(sorted(norderd.items())):
                    torder.append(value)
                self.sort_asc = False
                self.table.horizontalHeaderItem(col).setIcon(QtGui.QIcon('arrowd.png'))
            else:
                self.sort_asc = True
                for key, value in iter(sorted(iter(norderd.items()), reverse=True)):
                    torder.append(value)
                for key, value in iter(sorted(orderd.items())):
                    torder.append(value)
                self.table.horizontalHeaderItem(col).setIcon(QtGui.QIcon('arrowu.png'))
        self.entry = [self.entry[i] for i in torder]
        for rw in range(len(self.entry)):
            for cl in range(self.table.columnCount()):
                self.table.setItem(rw, cl, QtGui.QTableWidgetItem(''))
            for key, value in sorted(list(self.entry[rw].items()), key=lambda i: self.fields.index(i[0])):
                cl = self.fields.index(key)
                if value is not None:
                    if isinstance(value, list):
                        fld = str(value[0])
                        for i in range(1, len(value)):
                            fld = fld + ',' + str(value[i])
                        self.table.setItem(rw, cl, QtGui.QTableWidgetItem(fld))
                    else:
                        if cl > 0 and self.labels[key] == 'str':
                            value_pt = QtGui.QPlainTextEdit()
                            value_pt.setPlainText(value)
                            self.table.setCellWidget(rw, cl, value_pt)
                            if value.find('\n') >= 0:
                                self.table.setRowHeight(rw, self.rh * 2)
                                xtra_rows += 1
                        else:
                            self.table.setItem(rw, cl, QtGui.QTableWidgetItem(value))
                            if self.labels[key] != 'str':
                                self.table.item(rw, cl).setTextAlignment(130)   # x'82'
                if not self.edit_table:
                    self.table.item(rw, cl).setFlags(QtCore.Qt.ItemIsEnabled)
        return xtra_rows

    def showit(self):
        self.show()

    def header_click(self, position):
        column = self.headers.logicalIndexAt(position)
        self.order(column)

    def row_click(self, position):
        row = self.rows.logicalIndexAt(position)
        msgbox = QtGui.QMessageBox()
        msgbox.setWindowTitle('Delete item')
        msgbox.setText("Press Yes to delete '" + str(self.table.item(row, 0).text()) + "'")
        msgbox.setIcon(QtGui.QMessageBox.Question)
        msgbox.setStandardButtons(QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        reply = msgbox.exec_()
        if reply == QtGui.QMessageBox.Yes:
            for i in range(len(self.entry)):
                if self.entry[i][self.fields[0]] == str(self.table.item(row, 0).text()):
                    del self.entry[i]
                    break
            self.table.removeRow(row)

    def addtotbl(self):
        if self.message.text() != ' ' and self.message.text()[0] != '(':
            return
        addproperty = {}
        for field in self.fields:
            addproperty[field] = ''
        if len(self.fields) == 2:
            duplicate = []
            for entry in self.entry:
                duplicate.append(entry[self.fields[0]])
        else:
            duplicate = None
        dialog = displayobject.AnObject(QtGui.QDialog(), addproperty, readonly=False,
                 textedit=True, title='Add ' + self.fields[0].title() + ' value', duplicate=duplicate)
        try:
            dialog.edit[0].setFocus()
        except:
            pass
        dialog.exec_()
        if dialog.getValues()[self.fields[0]] != '':
            self.recur = True
            self.entry.append(addproperty)
            self.table.setRowCount(self.table.rowCount() + 1)
            self.sort_col = 1
            self.order(0)
        del dialog
        self.recur = False

    def item_selected(self, row, col):
        self.selection = self.entry[row][self.fields[0]]
        self.close()
        return
        for thing in self.e:
            try:
                attr = getattr(thing, self.fields[0])
                if attr == self.entry[row][self.fields[0]]:
                    self.selection = attr
                    self.close()
                    return
                    dialog = displayobject.AnObject(QtGui.QDialog(), thing)
                    dialog.exec_()
                    break
            except:
                pass

    def item_changed(self, row, col):
        if self.recur:
            return
        self.message.setText(' ')
        self.entry[row][self.fields[col]] = str(self.table.item(row, col).text())
        if col == 0 and len(self.fields) == 2 and self.labels[self.fields[col]] == 'str':
            newvalue = str(self.table.item(row, col).text())
            for rw in range(len(self.entry)):
                if rw == row:
                    continue
                if self.entry[rw][self.fields[col]] == newvalue:
                    self.message.setText('Duplicate value not allowed for ' + self.fields[col])
                    self.recur = False
                    break
        else:
            if self.labels[self.fields[col]] == 'int' or self.labels[self.fields[col]] == 'float':
                self.recur = True
                tst = str(self.table.item(row, col).text().replace(',', ''))
                if len(tst) < 1:
                    return
                if self.labels[self.fields[col]] == 'int':
                    try:
                        tst = int(tst)
                        fmat_str = '{: >' + str(self.lens[self.fields[col]][0] + self.lens[self.fields[col]][1] + 1) + ',.' \
                                   + str(self.lens[self.fields[col]][1]) + 'f}'
                        self.table.setItem(row, col, QtGui.QTableWidgetItem(fmat_str.format(tst)))
                        self.table.item(row, col).setTextAlignment(130)  # x'82'
                        self.recur = False
                        return
                    except:
                        self.message.setText('Error with ' + self.fields[col].title() + ' field - ' + tst)
                        self.recur = False
                else:
                    try:
                        tst = float(tst)
                        fmat_str = '{: >' + str(self.lens[self.fields[col]][0] + self.lens[self.fields[col]][1] + 1) + ',.' \
                                   + str(self.lens[self.fields[col]][1]) + 'f}'
                        self.table.setItem(row, col, QtGui.QTableWidgetItem(fmat_str.format(tst)))
                        self.table.item(row, col).setTextAlignment(130)  # x'82'
                        self.recur = False
                        return
                    except:
                        self.message.setText('Error with ' + self.fields[col].title() + ' field - ' + tst)
                        self.recur = False
        msg_font = self.message.font()
        msg_font.setBold(True)
        self.message.setFont(msg_font)
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.message.setPalette(msg_palette)
        self.table.resizeRowsToContents()

    def quit(self):
        self.close()

    def saveit(self):
        if self.title is None:
            iam = getattr(self.objects[0], '__module__')
        else:
            iam = self.title
        data_file = '%s_Table_%s.xls' % (iam,
                    str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm')))
        data_file = QtGui.QFileDialog.getSaveFileName(None, 'Save ' + iam + ' Table',
                    self.save_folder + data_file, 'Excel Files (*.xls*);;CSV Files (*.csv)')
        if data_file == '':
            return
        data_file = str(data_file)
        if data_file[-4:] == '.csv' or data_file[-4:] == '.xls' or data_file[-5:] == '.xlsx':
            pass
        else:
            data_file += '.xls'
        if os.path.exists(data_file):
            if os.path.exists(data_file + '~'):
                os.remove(data_file + '~')
            os.rename(data_file, data_file + '~')
        if data_file[-4:] == '.csv':
            tf = open(data_file, 'w')
            hdr_types = []
            line = ''
            for cl in range(self.table.columnCount()):
                if cl > 0:
                    line += ','
                hdr = str(self.table.horizontalHeaderItem(cl).text())
                if hdr[0] != '%':
                    txt = hdr
                    if ',' in txt:
                        line += '"' + txt + '"'
                    else:
                        line += txt
                txt = self.hdrs[hdr]
                try:
                    hdr_types.append(self.labels[txt.lower()])
                except:
                    hdr_types.append(self.labels[txt])
            tf.write(line + '\n')
            for rw in range(self.table.rowCount()):
                line = ''
                for cl in range(self.table.columnCount()):
                    if cl > 0:
                        line += ','
                    if self.table.item(rw, cl) is not None:
                        txt = str(self.table.item(rw, cl).text())
                        if hdr_types[cl] == 'int':
                            try:
                                txt = str(self.table.item(rw, cl).text()).strip()
                            except:
                                pass
                        elif hdr_types[cl] == 'float':
                            try:
                                txt = str(self.table.item(rw, cl).text()).strip()
                                txt = txt.replace(',', '')
                            except:
                                pass
                        if ',' in txt:
                            line += '"' + txt + '"'
                        else:
                            line += txt
                tf.write(line + '\n')
            tf.close()
        else:
            wb = xlwt.Workbook()
            ws = wb.add_sheet(str(iam))
            hdr_types = []
            dec_fmts = []
            xl_lens = []
            for cl in range(self.table.columnCount()):
                hdr = str(self.table.horizontalHeaderItem(cl).text())
                if hdr[0] != '%':
                    ws.write(0, cl, hdr)
                txt = self.hdrs[hdr]
                try:
                    hdr_types.append(self.labels[txt.lower()])
                    txt = txt.lower()
                except:
                    try:
                        hdr_types.append(self.labels[txt])
                    except:
                        hdr_types.append('str')
                style = xlwt.XFStyle()
                try:
                    if self.lens[txt][1] > 0:
                        style.num_format_str = '#,##0.' + '0' * self.lens[txt][1]
                    elif self.labels[txt] == 'int' or self.labels[txt] == 'float':
                        style.num_format_str = '#,##0'
                except:
                    pass
                dec_fmts.append(style)
                xl_lens.append(len(hdr))
            for rw in range(self.table.rowCount()):
                for cl in range(self.table.columnCount()):
                    if self.table.item(rw, cl) is not None:
                        valu = str(self.table.item(rw, cl).text())
                        if hdr_types[cl] == 'int':
                            try:
                                val1 = str(self.table.item(rw, cl).text()).strip()
                                val1 = val1.replace(',', '')
                                valu = int(val1)
                            except:
                                pass
                        elif hdr_types[cl] == 'float':
                            try:
                                val1 = str(self.table.item(rw, cl).text()).strip()
                                val1 = val1.replace(',', '')
                                valu = float(val1)
                            except:
                                pass
                        xl_lens[cl] = max(xl_lens[cl], len(str(valu)))
                        ws.write(rw + 1, cl, valu, dec_fmts[cl])
            for cl in range(self.table.columnCount()):
                if xl_lens[cl] * 275 > ws.col(cl).width:
                    ws.col(cl).width = xl_lens[cl] * 275
            ws.set_panes_frozen(True)   # frozen headings instead of split panes
            ws.set_horz_split_pos(1)   # in general, freeze after last heading row
            ws.set_remove_splits(True)   # if user does unfreeze, don't leave a split there
            wb.save(data_file)
            self.savedfile = data_file
        self.close()

    def replacetbl(self):
        if self.message.text() != ' ' and self.message.text()[0] != '(':
            return
        if self.ois == 'list':
            self.replaced = []
            if self.table.columnCount() == 1:
                cl = 0
                for rw in range(self.table.rowCount()):
                    if self.table.item(rw, cl) is None:
                        if self.labels[self.fields[cl]] == 'int' or self.labels[self.fields[cl]] == 'float':
                            self.replaced.append(0)
                        else:
                            self.replaced.append('')
                    else:
                        if self.labels[self.fields[cl]] == 'int':
                            self.replaced.append(int(self.table.item(rw, cl).text()).strip().replace(',', ''))
                        elif self.labels[self.fields[cl]] == 'float':
                            self.replaced.append(float(self.table.item(rw, cl).text()).strip().replace(',', ''))
                        else:
                            self.replaced.append(str(self.table.item(rw, cl).text()))
                self.close()
                return
            for rw in range(self.table.rowCount()):
                self.replaced.append([])
                for cl in range(self.table.columnCount()):
                    if self.table.item(rw, cl) is None:
                        if self.labels[self.fields[cl]] == 'int' or self.labels[self.fields[cl]] == 'float':
                            self.replaced[-1].append(0)
                        else:
                            self.replaced[-1].append('')
                    else:
                        if self.labels[self.fields[cl]] == 'int':
                            self.replaced[-1].append(int(self.table.item(rw, cl).text()).strip().replace(',', ''))
                        elif self.labels[self.fields[cl]] == 'float':
                            self.replaced[-1].append(float(self.table.item(rw, cl).text()).strip().replace(',', ''))
                        else:
                            try:
                                self.replaced[-1].append(self.table.cellWidget(rw, cl).toPlainText())
                            except:
                                self.replaced[-1].append(str(self.table.item(rw, cl).text()))
            self.close()
            return
        self.replaced = {}
        for rw in range(self.table.rowCount()):
            key = str(self.table.item(rw, 0).text())
            values = []
            for cl in range(1, self.table.columnCount()):
                if self.table.item(rw, cl) is not None:
                    if self.labels[self.fields[cl]] == 'int' or self.labels[self.fields[cl]] == 'float':
                        tst = str(self.table.item(rw, cl).text()).strip().replace(',', '')
                        if tst == '':
                            valu = tst
                        else:
                            if self.labels[self.fields[cl]] == 'int':
                                try:
                                    valu = int(tst)
                                    if valu == 0:
                                        valu = ''
                                    else:
                                        valu = str(valu)
                                except:
                                    self.message.setText('Error with ' + self.fields[cl].title() + ' field - ' + tst)
                                    self.replaced = None
                                    return
                            else:
                                try:
                                    valu = float(tst)
                                    if valu == 0:
                                        valu = ''
                                    elif len(str(valu).split('.')[1]) > 1:
                                        valu = str(valu)
                                except:
                                    self.message.setText('Error with ' + self.fields[cl].title() + ' field - ' + tst)
                                    self.replace = None
                                    return
                    else:
                        valu = str(self.table.item(rw, cl).text())
                else:
                    valu = ''
                if valu != '':
                    values.append(self.fields[cl] + '=' + valu)
            self.replaced[key] = values
        self.close()

    def getValues(self):
        if self.edit_table:
            return self.replaced
        return None

    def getChoice(self):
        return self.selection
