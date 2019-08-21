#!/usr/bin/python3
#
#  Copyright (C) 2019 Angus King
#
#  catalogue.py - This file is part of catalogue.
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

import datetime
from functools import partial
import os
try:
    import pwd
except:
    pass
from PyQt4 import QtCore, QtGui
import sqlite3
from sqlite3 import Error
import sys
import tempfile
import time
import webbrowser

import displayobject
import displaytable
from functions import *


class TabDialog(QtGui.QMainWindow):

    def wheelEvent(self, event):
        if len(self.rows) == 0:
            return
        self.wrapmsg.setText('')
        if event.delta() < 0:
            self.row += 1
            if self.row >= len(self.rows):
                self.row = 0
                self.wrapmsg.setText('Wrapped to top')
        else:
            self.row -= 1
            if self.row < 0:
                self.row = len(self.rows) - self.pagerows.value()
                if self.row < 0:
                    self.row = 0
                else:
                    self.wrapmsg.setText('Wrapped to bottom')
        self.getRows()

    def resizeEvent(self, event):
    #    print('(43) resize (%d x %d)' % (event.size().width(), event.size().height()))
        QtGui.QWidget.resizeEvent(self, event)
        w = event.size().width()
        h = event.size().height()

    def __init__(self, parent=None):
        QtGui.QMainWindow.__init__(self, parent)
        self.find_file = False
        self.conn = None
        self.db = None
        self.dbs = configFile()
        self.cattitle = QtGui.QLabel('')
        self.items = QtGui.QLabel('')
        self.metacombo = QtGui.QComboBox(self)
        self.category = 'Category'
        self.catcombo = QtGui.QComboBox(self)
        self.catcombo.setHidden(True)
        self.attr_cat = QtGui.QAction(QtGui.QIcon('copy.png'), self.category, self)
        self.attr_cat.setShortcut('Ctrl+C')
        self.attr_cat.setStatusTip('Edit category values')
        self.attr_cat.triggered.connect(self.editFields)
        if len(self.dbs) > 0:
            self.conn = create_connection(self.dbs[0])
            if self.conn is not None:
                self.db = self.dbs[0]
                self.updDetails()
        self.me = '/' + sys.argv[0][:sys.argv[0].rfind('.')]
        self.me = self.me[self.me.rfind('/') + 1:].title()
        self.setWindowTitle(self.me + ' (' + fileVersion() + ') - Simple Catalogue')
        self.setWindowIcon(QtGui.QIcon('books.png'))
        buttonLayout = QtGui.QHBoxLayout()
        quitButton = QtGui.QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(quitButton)
        self.connect(quitButton, QtCore.SIGNAL('clicked()'), self.quit)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        addButton = QtGui.QPushButton(self.tr('&Add'))
        buttonLayout.addWidget(addButton)
        self.connect(addButton, QtCore.SIGNAL('clicked()'), self.addItem)
        addfile = QtGui.QPushButton(self.tr('&Add File'))
        buttonLayout.addWidget(addfile)
        self.connect(addfile, QtCore.SIGNAL('clicked()'), self.addFile)
        QtGui.QShortcut(QtGui.QKeySequence('pgdown'), self, self.nextRows)
        QtGui.QShortcut(QtGui.QKeySequence('pgup'), self, self.prevRows)
        buttons = QtGui.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtGui.QGridLayout()
        layout.setVerticalSpacing(10)
        layout.addWidget(QtGui.QLabel('Catalogue:'), 0, 0)
        layout.addWidget(self.cattitle, 0, 1, 1, 2)
        layout.addWidget(self.items, 0, 4)
        layout.addWidget(QtGui.QLabel('Search by'), 1, 0)
        layout.addWidget(self.metacombo, 1, 1)
        self.search = QtGui.QLineEdit()
        self.filter = QtGui.QComboBox(self)
        self.filter.addItem('equals')
        self.filter.addItem('starts with')
        self.filter.addItem('contains')
        self.filter.addItem('missing')
        self.filter.addItem('duplicate')
        self.filter.setCurrentIndex(1)
        layout.addWidget(self.filter, 1, 2)
        layout.addWidget(self.catcombo, 1, 3)
        layout.addWidget(self.search, 1, 3)
        srchb = QtGui.QPushButton('Search')
        layout.addWidget(srchb, 1, 4)
        self.connect(srchb, QtCore.SIGNAL('clicked()'), self.do_search)
        enter = QtGui.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Return), self)
        enter.activated.connect(self.do_search)
        layout.addWidget(buttons, 2, 0, 1, 2)
        msgLayout = QtGui.QHBoxLayout()
        msgLayout.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.wrapmsg = QtGui.QLabel('')
        msg_font = self.wrapmsg.font()
        msg_font.setBold(True)
        self.wrapmsg.setFont(msg_font)
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.wrapmsg.setPalette(msg_palette)
        msgLayout.addWidget(self.wrapmsg)
        msgLayout.addWidget(QtGui.QLabel('Rows per page:'))
        self.pagerows = QtGui.QSpinBox()
        msgLayout.addWidget(self.pagerows)
        self.srchrng = QtGui.QLabel('')
        self.srchrng.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        msgLayout.addWidget(self.srchrng)
        self.srchmsg = QtGui.QLabel('')
        msgLayout.addWidget(self.srchmsg)
        msgs = QtGui.QFrame()
        msgs.setLayout(msgLayout)
        layout.addWidget(msgs, 2, 2, 1, 3)
        layout.setColumnStretch(3, 10)
        menubar = QtGui.QMenuBar()
        layout.setMenuBar(menubar)
        db_opn = QtGui.QAction(QtGui.QIcon('open.png'), 'Open', self)
        db_opn.setShortcut('Ctrl+O')
        db_opn.setStatusTip('Open Catalogue')
        db_opn.triggered.connect(self.openDB)
        db_add = QtGui.QAction(QtGui.QIcon('plus.png'), 'Add', self)
        db_add.setShortcut('Ctrl+A')
        db_add.setStatusTip('Add Catalogue')
        db_add.triggered.connect(self.addDB)
        db_new = QtGui.QAction(QtGui.QIcon('books.png'), 'New', self)
        db_new.setShortcut('Ctrl+N')
        db_new.setStatusTip('Create Catalogue')
        db_new.triggered.connect(self.newDB)
        db_lod = QtGui.QAction(QtGui.QIcon('load.png'), 'Load', self)
        db_lod.setShortcut('Ctrl+L')
        db_lod.setStatusTip('Load Catalogue')
        db_lod.triggered.connect(self.loadDB)
        db_rem = QtGui.QAction(QtGui.QIcon('cancel.png'), 'Remove', self)
        db_rem.setStatusTip('Remove Catalogue')
        db_rem.triggered.connect(self.remDB)
        db_qui = QtGui.QAction(QtGui.QIcon('quit.png'), 'Quit', self)
        db_qui.setShortcut('Ctrl+Q')
        db_qui.setStatusTip('Quit')
        db_qui.triggered.connect(self.quit)
        dbMenu = menubar.addMenu('&Catalogue')
        dbMenu.addAction(db_opn)
        dbMenu.addAction(db_add)
        dbMenu.addAction(db_new)
        dbMenu.addAction(db_lod)
        dbMenu.addAction(db_rem)
        dbMenu.addAction(db_qui)
        attrMenu = menubar.addMenu('&Attributes')
        attr_inf = QtGui.QAction(QtGui.QIcon('edit.png'), 'Info', self)
        attr_inf.setShortcut('Ctrl+E')
        attr_inf.setStatusTip('Edit Catalogue info')
        attr_inf.triggered.connect(self.editFields)
        attr_met = QtGui.QAction(QtGui.QIcon('list.png'), 'Meta', self)
        attr_met.setShortcut('Ctrl+M')
        attr_met.setStatusTip('Edit Meta values')
        attr_met.triggered.connect(self.editFields)
        attr_set = QtGui.QAction(QtGui.QIcon('edit.png'), 'Settings', self)
        attr_set.setShortcut('Ctrl+S')
        attr_set.setStatusTip('Edit Catalogue setting')
        attr_set.triggered.connect(self.editFields)
        attrMenu.addAction(self.attr_cat)
        attrMenu.addAction(attr_inf)
        attrMenu.addAction(attr_met)
        attrMenu.addAction(attr_set)
        help = QtGui.QAction(QtGui.QIcon('help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        about = QtGui.QAction(QtGui.QIcon('info.png'), 'About', self)
        about.setShortcut('Ctrl+I')
        about.setStatusTip('About')
        about.triggered.connect(self.showAbout)
        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(help)
        helpMenu.addAction(about)
        # assume height of system window titlebar at 50
        toplen = 50 + buttonLayout.sizeHint().height() + msgLayout.sizeHint().height() + \
                 dbMenu.geometry().height() + layout.sizeHint().height() + \
                 layout.verticalSpacing() * 3 + 10
        self.table = QtGui.QTableWidget()
        self.table.setRowCount(1)
        self.table.setItem(0, 0, QtGui.QTableWidgetItem(''))
        rh = self.table.rowHeight(0)
        screen = QtGui.QDesktopWidget().availableGeometry()
        table_rows = int((screen.height() - toplen) / rh)
        self.pagerows.setRange(1, table_rows)
        self.pagerows.setValue(table_rows)
        self.table.setRowCount(table_rows)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['', '', 'Title', self.category])
        self.table.setColumnWidth(0, 30)
        self.table.setColumnWidth(1, 0)
        self.table.setColumnWidth(2, 800)
        self.table.setColumnWidth(3, 200)
        self.table.setEditTriggers(QtGui.QAbstractItemView.SelectedClicked)
        self.table.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
  #      self.table.setSizePolicy(QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        rows = self.table.verticalHeader()
        rows.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        rows.customContextMenuRequested.connect(self.row_click)
        rows.setSelectionMode(QtGui.QAbstractItemView.SingleSelection)
        self.table.verticalHeader().setVisible(True)
        self.table.cellClicked.connect(self.item_selected)
        layout.addWidget(self.table, 3, 0, 1, 5)
        centralWidget = QtGui.QWidget(self)
        self.setCentralWidget(centralWidget)
        centralWidget.setLayout(layout)
        self.metacombo.currentIndexChanged.connect(self.metaChanged)
        self.catcombo.currentIndexChanged.connect(self.catChanged)
        size = QtCore.QSize(screen.width(), screen.height())
        size = QtCore.QSize(min(screen.width(), 1100), screen.height() - 50)
        self.resize(size)
        if self.conn is not None:
            self.updDetails()
            self.do_search()

    def openDB(self):
        if len(self.dbs) < 1:
            self.newDB()
            return
        fields = ['database']
        dialog = displaytable.Table(self.dbs, fields=fields)
        dialog.exec_()
        db = dialog.getChoice()
        if db is not None:
            if not os.path.exists(db):
                self.wrapmsg.setText(db + ' not found')
                return
            if self.conn is not None:
                self.conn.close()
            self.conn = create_connection(db)
            self.db = db
            self.wrapmsg.setText('Catalogue opened')
            self.updDetails()
            self.do_search()

    def addDB(self):
        db_file = 'catalogue.db'
        db_file = QtGui.QFileDialog.getOpenFileName(None, 'Add Catalogue',
                  db_file, 'Database Files (*.db)')
        if db_file == '':
            return
        if self.conn is not None:
            self.conn.close()
        self.conn = create_connection(db_file)
        self.db = db_file
        self.dbs.append(db_file)
        self.wrapmsg.setText('Catalogue added')
        self.updDetails()
        self.do_search()

    def remDB(self):
        fields = ['database']
        dialog = displaytable.Table(self.dbs, fields=fields)
        dialog.exec_()
        db = dialog.getChoice()
        if db is not None:
            if self.db == db:
                if self.conn is not None:
                    self.conn.close()
            dell = '.'
            if os.path.exists(db):
                msgbox = QtGui.QMessageBox()
                msgbox.setText('Do you want to DELETE the database (Y)?')
                msgbox.setIcon(QtGui.QMessageBox.Question)
                msgbox.setStandardButtons(QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
                reply = msgbox.exec_()
                if reply == QtGui.QMessageBox.Yes:
                    os.remove(db)
                    dell = ' and deleted.'
            for f in range(len(self.dbs)):
                if self.dbs[f] == db:
                    del self.dbs[f]
                    break
            self.wrapmsg.setText(db[db.rfind('/')+1:] + ' removed' + dell)
            if self.db == db:
                self.db = None
                self.openDB()

    def newDB(self):
        db_file = 'catalogue.db'
        db_file = QtGui.QFileDialog.getSaveFileName(None, 'Create Catalogue',
                  db_file, 'Database Files (*.db)')
        if db_file == '':
            return
        conn = create_catalogue(db_file)
        if conn is None:
            return
        conn.close()
        if self.conn is not None:
            self.conn.close()
        self.conn = create_connection(db_file)
        self.db = db_file
        self.updDetails()
        if db_file not in self.dbs:
            self.dbs.append(db_file)
        self.wrapmsg.setText('Catalogue created')
        self.loadDB()

    def loadDB(self):
        data_file = QtGui.QFileDialog.getOpenFileName(None, 'Choose data to load',
                  '', 'Excel Files (*.xls*);;CSV files (*.csv)')
        if data_file == '':
            return
        message = load_catalogue(self, self.db, data_file)
        self.wrapmsg.setText(message)
        self.updDetails()
        self.do_search()

    def updDetails(self):
        cur = self.conn.cursor()
        cur.execute("select description from fields where typ = 'Info' and field = 'Catalogue'" \
                    + " or field = 'Description' order by field")
        try:
            self.cattitle.setText('<strong>' + cur.fetchone()[0] + '</strong> - ' + cur.fetchone()[0])
        except:
            self.cattitle.setText('<strong>' + self.db + '</strong>')
        self.category = 'Category'
        self.pdf_decrypt = 'qpdf'
        self.launcher = ''
        self.translate_user = '$USER$'
        self.category_multi = False
        self.url_field = ''
        cur.execute("select field, description from fields where typ = 'Settings'")
        row = cur.fetchone()
        while row is not None:
            if row[0] == 'Category Choice':
                if row[1][:5].lower() == 'multi':
                    self.category_multi = True
            if row[0] == 'Category Field':
                self.category = row[1].title()
            elif row[0] == 'Decrypt PDF':
                self.pdf_decrypt = row[1]
            elif row[0] == 'Launch File':
                self.launcher = row[1]
            elif row[0] == 'Translate Userid':
                self.translate_user = row[1]
            elif row[0] == 'URL Field':
                self.url_field = row[1].upper()
            row = cur.fetchone()
        self.attr_cat.setText(self.category)
        try:
            if self.launcher == '':
                self.table.setColumnHidden(0, True)
            else:
                self.table.setColumnHidden(0, False)
        except:
            pass
        cur.execute("select count(*) from items")
        cnt = cur.fetchone()[0]
        if cnt > 0:
            self.items.setText('(' + str(cnt) + ' items)')
        else:
            self.items.setText('')
        self.metacombo.clear()
        self.metacombo.addItem('All')
        self.metacombo.addItem(self.category)
        self.metacombo.addItem('Title')
        self.metacombo.addItem('Filename')
        self.metacombo.addItem('Location')
        cur.execute("select field from fields where typ = 'Meta' order by field")
        row = cur.fetchone()
        cnt = 0
        while row is not None:
            if row[0] == 'keyword':
                cnt = self.metacombo.count()
            if row[0] != self.category:
                self.metacombo.addItem(row[0])
            row = cur.fetchone()
        self.metacombo.setCurrentIndex(cnt)
        cur.close()

    def editFields(self):
        if self.conn is None:
            return
        cur = self.conn.cursor()
        cur.execute("select field, description from fields where typ = '" + self.sender().text() + "'")
        rows = []
        row = cur.fetchone()
        while row is not None:
            rows.append([row[0], row[1]])
            row = cur.fetchone()
        cur.close()
        fields = [self.sender().toolTip()]
        if self.sender().text() == 'Settings':
            fields.append('Value')
        else:
            fields.append('Description')
        dialog = displaytable.Table(rows, fields=fields, edit=True)
        dialog.exec_()
        if dialog.getValues() is None:
            return
        meta = False
        info = False
        if self.sender().text() == 'Meta':
            self.metacombo.clear()
            self.metacombo.addItem('All')
            self.metacombo.addItem(self.category)
            self.metacombo.addItem('Title')
            meta = True
            cnt = 0
        elif self.sender().text() == 'Info':
            info = True
            cat_title = ['', '']
        upd_rows = dict(dialog.getValues())
        cur = self.conn.cursor()
        updcur = self.conn.cursor()
        cur.execute("select field, description from fields where typ = '" + self.sender().text() + "'")
        row = cur.fetchone()
        while row is not None:
            if row[0] in upd_rows:
                if row[1] != upd_rows[row[0]]:
                    sql = "update fields set description = ? where typ = ? and field = ?"
                    updcur.execute(sql, (upd_rows[row[0]], self.sender().text(), row[0]))
                if info:
                    if row[0] == 'Catalogue':
                        cat_title[0] = '<strong>' + upd_rows[row[0]] + '</strong>'
                    elif row[0] == 'Description':
                        cat_title[1] = ' - ' + upd_rows[row[0]]
                        self.cattitle.setText('<strong>' + upd_rows[row[0]] + '</strong>')
                del upd_rows[row[0]]
                if meta:
                    if row[0] == 'Keyword':
                        cnt = self.metacombo.count()
                    if row[0] != self.category:
                        self.metacombo.addItem(row[0])
            else:
                sql = "select count(*) from meta where field = ?"
                updcur.execute(sql, (row[0], ))
                i = updcur.fetchone()[0]
                if i > 0:
                    msgbox = QtGui.QMessageBox()
                    msgbox.setWindowTitle('Delete ' + self.sender().text() + ' value')
                    msgbox.setText(str(i) + ' rows for ' + row[0] \
                                   + '. Do you still want to delete it?')
                    msgbox.setIcon(QtGui.QMessageBox.Question)
                    msgbox.setStandardButtons(QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
                    reply = msgbox.exec_()
                if i == 0 or reply == QtGui.QMessageBox.Yes:
                   sql = "delete from fields where typ = ? and field = ?"
                   updcur.execute(sql, (self.sender().text(), row[0]))
                   if i > 0:
                       msgbox.setText('Do you want to delete the Meta rows?')
                       reply = msgbox.exec_()
                       if reply == QtGui.QMessageBox.Yes:
                           sql = "delete from meta where field = ?"
                           updcur.execute(sql, (row[0], ))
                   if row[0].upper() == self.url_field:
                       sql = "delete from fields where typ = 'Settings' and field = 'URL Field'"
                       updcur.execute(sql)
            row = cur.fetchone()
        for key, value in upd_rows.items():
            if key == '':
                continue
            sql = "insert into fields (typ, field, description) values (?, ?, ?)"
            updcur.execute(sql, (self.sender().text(), key, value))
            if meta:
                if key == 'Keyword':
                    cnt = self.metacombo.count()
                elif key.upper() == 'URL':
                    cur.execute("select description from fields where typ = 'Settings'" + \
                                " and field = 'URL Field'")
                    row = cur.fetchone()
                    if row is None:
                        sql = "insert into fields (typ, field, description) values (?, ?, ?)"
                        updcur.execute(sql, ('Settings', 'URL Field', 'URL'))
                        self.url_field = 'URL'
                self.metacombo.addItem(key)
            elif info:
                if key == 'Catalogue':
                    cat_title[0] = '<strong>' + value + '</strong>'
                elif row[0] == 'Description':
                    cat_title[1] = ' - ' + value
        if info:
            self.cattitle.setText(cat_title[0] + cat_title[1])
        cur.close()
        updcur.close()
        self.conn.commit()
        if meta:
            self.metacombo.setCurrentIndex(cnt)
        elif self.sender().text() == 'Settings':
            self.updDetails()

    def addFile(self):
        if self.conn is None:
            return
        new_file = QtGui.QFileDialog.getOpenFileName(None, 'Add File',
                   self.db[:self.db.rfind('/')], 'All Files (*.*)')
        if new_file == '':
            return
        properties = {}
        folder = new_file[:new_file.rfind('/') + 1]
        properties['Title'] = new_file[len(folder):]
        properties['Filename'] = new_file[len(folder):]
        stat = os.stat(new_file)
        properties['Filesize'] = str(stat.st_size)
        if self.translate_user != '':
            folder = folder.replace(getUser(), self.translate_user)
        properties['Location'] = folder
        properties['Acquired'] = time.strftime('%Y-%m-%d', time.localtime(stat.st_mtime))
        if new_file[-4:].lower() == '.pdf':
            properties = getPDFInfo(new_file, properties=properties, decrypt=self.pdf_decrypt)
        self.addItem(properties=properties)

    def addItem(self, properties=None):
        if self.conn is None:
            return
        addproperty = {}
        fields = ['Title', 'Filename', 'Location']
        for field in fields:
            if properties is not None and field in properties.keys():
                addproperty[field] = properties[field]
            else:
                addproperty[field] = ''
        cur = self.conn.cursor()
        cur.execute("select field from fields where typ = 'Meta' order by field")
        row = cur.fetchone()
        add_cat_field = True
        while row is not None:
            if row[0] == self.category:
                add_cat_field = False
            if properties is not None and row[0] in properties.keys():
                addproperty[row[0]] = properties[row[0]]
            else:
                addproperty[row[0]] = ''
            row = cur.fetchone()
        combolist = [self.category, []]
        cur.execute("select field from fields where typ = '" + self.category + "' order by field")
        row = cur.fetchone()
        while row is not None:
            combolist[1].append(row[0])
            row = cur.fetchone()
        dialog = displayobject.AnObject(QtGui.QDialog(), addproperty, readonly=False,
                 textedit=True, title='Add Item', combolist=combolist, multi=self.category_multi)
        dialog.exec_()
        if dialog.getValues() is None or dialog.getValues()['Title'] == '':
            cur.close()
            return
        folder = dialog.getValues()['Location']
        if self.translate_user != '':
            folder = folder.replace(getUser(), self.translate_user)
        titl = dialog.getValues()['Title'].replace('\n', ' ').replace('\r', '')
        sql = "insert into items (Title, Filename, Location) values (?, ?, ?)"
        cur.execute(sql, (titl, dialog.getValues()['Filename'] , folder))
        sql = "select last_insert_rowid()"
   #     " + self.category + ",
        cur.execute(sql)
        iid = cur.fetchone()[0]
        if add_cat_field:
            sql = "insert into fields (typ, field, description) values (?, ?, ?)"
            cur.execute(sql, ('Meta', self.category, 'Category field'))
        sql = "insert into meta (field, item_id, value) values (?, ?, ?)"
        for key, value in dialog.getValues().items():
            if key in fields:
                continue
            if value == '':
                if key == 'Acquired':
                    value = datetime.now().strftime('%Y-%m-%d')
                else:
                    continue
            if isinstance(value, list):
                for valu in value:
                    cur.execute(sql, (key, iid, valu))
            else:
                cur.execute(sql, (key, iid, value))
        cur.execute("select count(*) from items")
        cnt = cur.fetchone()[0]
        if cnt > 0:
            self.items.setText('(' + str(cnt) + ' items)')
        else:
            self.items.setText('')
        cur.close()
        self.conn.commit()
        self.do_search()

    def metaChanged(self):
        if self.metacombo.currentText() == self.category:
            self.catcombo.clear()
            cur = self.conn.cursor()
            cur.execute("select field from fields where typ = '" + self.category.title() + "' order by field")
            row = cur.fetchone()
            while row is not None:
                self.catcombo.addItem(row[0])
                row = cur.fetchone()
            cur.close()
            self.filter.setCurrentIndex(0)
            self.search.setHidden(True)
            self.catcombo.setHidden(False)
        else:
            self.catcombo.setHidden(True)
            self.search.setHidden(False)

    def catChanged(self):
        self.search.setText(self.catcombo.currentText())
        self.do_search()

    def do_search(self):
        # think about like options - '%S', '%S%', '%S_S%' '[a-zA-Z0-9_]%'
        self.field = self.metacombo.currentText()
        self.rows = []
        search = self.search.text()
        self.find_file = False
        if self.filter.currentText() == 'missing' and self.field == 'Filename':
            where = "<> ''"
            self.find_file = True
        elif self.filter.currentText() == 'missing' and self.field not in ['All', 'Title', 'Location']:
            if self.field == self.category:
                pass
            else:
                search = '_%'
                where = "like ?"
        elif self.filter.currentText() == 'duplicate':
            if search != '':
                search = search + '%'
            where = "like ?"
        elif self.filter.currentText() == 'equals':
            where = "= ?"
        elif self.filter.currentText() == 'starts with':
            search = search + '%'
            where = "like ?"
        elif self.filter.currentText() == 'contains':
            search = '%' + search + '%'
            where = "like ?"
        else:
            return
        cur = self.conn.cursor()
        if self.field == 'All':
            sql = "select (id) from items where title " + where + \
                  " or filename " + where + \
                  " or location " + where + \
                  " or id in (select (item_id) from meta where value " + where + \
                  " ) order by title"
            cur.execute(sql, (search, search, search, search))
        elif self.filter.currentText() == 'duplicate':
            if self.field in ['Title', 'Filename', 'Location']:
                if search != '':
                    sql = "select o.id from items o inner join (select " + self.field + \
                      " from items where " + self.field + " " + where + " group by " + \
                      self.field + " having count(" + self.field + ") > 1) i on o." + \
                      self.field + " = i." + self.field + " order by o." + self.field
                    print(sql, search)
                    cur.execute(sql, (search, ))
                else:
                    sql = "select o.id from items o inner join (select " + self.field + \
                          " from items group by " + self.field + " having count(" + self.field + \
                          ") > 1) i on o." + self.field + " = i." + self.field + " order by o." + self.field
                    cur.execute(sql)
            else:
                if search != '':
                    sql = "select (item_id) from meta where field = ? and value in (select value from meta where field = ? " + \
                          " and value " + where + " group by value having count(*) > 1) order by value"
                    print(self.field, search, sql)
                    cur.execute(sql, (self.field, self.field, search))
                else:
                    sql = "select (item_id) from meta where field = ? and value in (select value from meta where field = ? " + \
                          " group by value having count(*) > 1) order by value"
                    cur.execute(sql, (self.field, self.field))
        elif self.filter.currentText() == 'missing' and self.field == 'Filename':
            sql = "select id, location, filename from items where " + self.field + " " + where + " order by " + self.field
            cur.execute(sql)
        elif self.field in ['Title', 'Filename', 'Location']:
            sql = "select (id) from items where " + self.field + " " + where + " order by " + self.field
            cur.execute(sql, (search, ))
        elif self.filter.currentText() == 'missing' and self.field == self.category:
            sql = "select (id) from items where id not in (select item_id from meta where field = ?)"
            cur.execute(sql, (self.field, ))
        else:
            sql = "select (item_id) from meta where field = ? and value " + where + \
                  " order by value"
            if search == '_%':
                sql = "select (id) from items where id not in (" + sql + ") order by title"
            cur.execute(sql, (self.field, search))
        row = cur.fetchone()
        if self.find_file:
            while row is not None:
                if self.translate_user != '':
                    folder = row[1].replace(self.translate_user, getUser())
                else:
                    folder = row[1]
                if folder[:7] == 'file://':
                    folder = folder[7:]
                if not os.path.exists(folder + row[2]):
                    if row[0] not in self.rows:
                        self.rows.append(row[0])
                row = cur.fetchone()
        else:
            while row is not None:
                if row[0] not in self.rows:
                    self.rows.append(row[0])
                row = cur.fetchone()
        cur.close()
        self.srchmsg.setText(str(len(self.rows)) + ' items')
        self.srchrng.setText('')
        self.table.clear()
        self.table.setHorizontalHeaderLabels(['', self.field, 'Title', self.category])
        if len(self.rows) == 0:
            return
        if self.field in ['All', self.category, 'Title'] or search == '_%':
            self.table.setColumnWidth(1, 0)
            self.col_2 = False
        else:
            self.table.setColumnWidth(1, 200)
            self.col_2 = True
        self.row = 0
        self.getRows()
        return

    def getRows(self):
        self.table.clear()
        self.table.setRowCount(self.pagerows.value())
        self.table.setHorizontalHeaderLabels(['', self.field, 'Title', self.category])
        sql = "select title, filename from items where id = ?"
        sql1 = "select value from meta where item_id = ? and field = ? order by value"
        sql2 = "select " + self.field + " from items where id = ?"
        cur = self.conn.cursor()
        for trw in range(self.pagerows.value()):
            if self.row + trw >= len(self.rows):
                trw = trw - 1
                break
            cur.execute(sql, (self.rows[self.row + trw], ))
            row = cur.fetchone()
            if row[1] == '' and self.url_field != '':
                button = QtGui.QPushButton('', self.table)
                button.setIcon(QtGui.QIcon('url.png'))
                button.setFlat(True)
                self.table.setCellWidget(trw, 0, button)
                button.clicked.connect(partial(self._buttonItemClicked, trw))
            elif self.launcher != '':
                button = QtGui.QPushButton('', self.table)
                button.setIcon(QtGui.QIcon('open.png'))
                button.setFlat(True)
                self.table.setCellWidget(trw, 0, button)
                button.clicked.connect(partial(self._buttonItemClicked, trw))
            self.table.setItem(trw, 2, QtGui.QTableWidgetItem(row[0]))
            cur.execute(sql1, (self.rows[self.row + trw], self.category))
            txt = ''
            row = cur.fetchone()
            while row is not None:
                txt += ';' + row[0]
                row = cur.fetchone()
            txt = txt[1:]
            self.table.setItem(trw, 3, QtGui.QTableWidgetItem(txt))
            if self.col_2:
                if self.field in ['Filename', 'Location']:
                    cur.execute(sql2, (self.rows[self.row + trw], ))
                else:
                    cur.execute(sql1, (self.rows[self.row + trw], self.field))
                row = cur.fetchone()
                if row is not None:
                    self.table.setItem(trw, 1, QtGui.QTableWidgetItem(row[0]))
        self.srchrng.setText(str(self.row + 1) + ' to ' + str(self.row + 1 + trw) + ' of ')
        cur.close()

    def nextRows(self):
        if len(self.rows) < self.pagerows.value():
            return
        self.row += self.pagerows.value()
        if self.row >= len(self.rows):
            self.row = 0
            self.wrapmsg.setText('Wrapped to top')
        else:
            self.wrapmsg.setText('')
        self.getRows()

    def prevRows(self):
        if len(self.rows) < self.pagerows.value():
            return
        self.wrapmsg.setText('')
        if self.row - self.pagerows.value() < 0:
            if self.row > 0:
                self.row = 0
            else:
                self.row = len(self.rows) - self.pagerows.value()
                self.wrapmsg.setText('Wrapped to bottom')
        else:
             self.row = self.row - self.pagerows.value()
        self.getRows()

    def _buttonItemClicked(self, row):
        if self.row + row >= len(self.rows):
            return
        if self.launcher == '':
            return
        self.wrapmsg.setText('')
        sql = "select location, filename from items where id = ?"
        cur = self.conn.cursor()
        cur.execute(sql, (self.rows[self.row + row], ))
        item = cur.fetchone()
        cur.close()
        if item[1] == '':
            if self.url_field != '': # maybe we have a url
                sql = "select value from meta where item_id = ? and field = ?"
                cur = self.conn.cursor()
                cur.execute(sql, (self.rows[self.row + row], self.url_field))
                item = cur.fetchone()
                cur.close()
                if item is not None:
                    if item[0] != '':
                        webbrowser.open_new(item[0])
            return
        folder = item[0].replace(self.translate_user, getUser())
        if folder[:7] == 'file://':
            folder = folder[7:]
        if os.path.exists(folder + item[1]):
            os.system(self.launcher + ' "' + folder + item[1] + '"')
        elif self.find_file:
            self.findFolder(self.rows[self.row + row], folder, item[1])
        else:
            self.wrapmsg.setText(folder + item[1] + ' not found.')

    def item_selected(self, row, col):
        self.wrapmsg.setText('')
        if self.row + row >= len(self.rows):
            return
        fields = ['Title', 'Filename', 'Location']
        itmproperty = {}
        sql = "select title, filename, location from items where id = ?"
        cur = self.conn.cursor()
        cur.execute(sql, (self.rows[self.row + row], ))
        srow = cur.fetchone()
        for f in range(len(fields)):
            itmproperty[fields[f]] = srow[f]
        sql = "select field, value from meta where item_id = ? order by field, value"
        cur.execute(sql, (self.rows[self.row + row], ))
        srow = cur.fetchone()
        while srow is not None:
            if self.category_multi and srow[0] == self.category:
                if srow[0] not in itmproperty.keys():
                    itmproperty[srow[0]] = [srow[1]]
                else:
                    itmproperty[srow[0]].append(srow[1])
            else:
                itmproperty[srow[0]] = srow[1]
            srow = cur.fetchone()
        sql = "select field from fields where typ = 'Meta' order by field"
        cur.execute(sql)
        srow = cur.fetchone()
        while srow is not None:
            if srow[0] not in itmproperty.keys():
                itmproperty[srow[0]] = ''
            srow = cur.fetchone()
        combolist = [self.category, []]
        cur.execute("select field from fields where typ = '" + self.category.title() + "' order by field")
        srow = cur.fetchone()
        while srow is not None:
            combolist[1].append(srow[0])
            srow = cur.fetchone()
        cur.close()
        dialog = displayobject.AnObject(QtGui.QDialog(), itmproperty, readonly=False,
                 textedit=True, title='Edit Item (' + str(self.rows[self.row + row]) + ')',
                 combolist=combolist, multi=self.category_multi)
        dialog.exec_()
        if dialog.getValues() is None:
            return
        if dialog.getValues()['Title'] == '':
            return
        try:
            if isinstance(dialog.getValues()[self.category], list):
                txt = ''
                for valu in sorted(dialog.getValues()[self.category]):
                    txt += ';' + valu
                txt = txt[1:]
                self.table.setItem(row, 3, QtGui.QTableWidgetItem(txt))
            else:
                self.table.setItem(row, 3, QtGui.QTableWidgetItem(dialog.getValues()[self.category]))
            self.table.setItem(row, 2, QtGui.QTableWidgetItem(dialog.getValues()['Title']))
            if self.field != '' and self.field not in fields:
                self.table.setItem(row, 1, QtGui.QTableWidgetItem(dialog.getValues()[self.field]))
        except:
            pass
        cur = self.conn.cursor()
        updcur = self.conn.cursor()
        folder = dialog.getValues()['Location']
        if self.translate_user != '':
            folder = folder.replace(getUser(), self.translate_user)
        sql = "update items set title = ?, filename = ?, location = ? where id = ?"
        titl = dialog.getValues()['Title'].replace('\n', ' ').replace('\r', '')
        cur.execute(sql, (titl, dialog.getValues()['Filename'] , folder, self.rows[self.row + row]))
        sqlu = "update meta set value = ? where id = ?"
        sqld = "delete from meta where id = ?"
        sqli = "insert into meta (item_id, field, value) values (?, ?, ?)"
        sql = "select id, field, value from meta where item_id = ?"
        cur.execute(sql, (self.rows[self.row + row], ))
        srow = cur.fetchone()
        while srow is not None:
            if srow[1] in dialog.getValues().keys():
                if isinstance(dialog.getValues()[srow[1]], list):
                    try:
                        i = dialog.getValues()[srow[1]].index(srow[2])
                        del dialog.getValues()[srow[1]][i]
                    except:
                        updcur.execute(sqld, (srow[0],))
                else:
                    if dialog.getValues()[srow[1]].strip() == '':
                        updcur.execute(sqld, (srow[0],))
                    elif srow[2] != dialog.getValues()[srow[1]].strip():
                        updcur.execute(sqlu, (dialog.getValues()[srow[1]].strip(), srow[0]))
                    del dialog.getValues()[srow[1]]
            else:
                if dialog.getValues()[srow[1]].strip() != '':
                    updcur.execute(sqli, (self.rows[self.row + row], srow[1], dialog.getValues()[srow[1]].strip()))
                del dialog.getValues()[srow[1]]
            srow = cur.fetchone()
        for key, value in dialog.getValues().items():
            if key in fields:
                continue
            if isinstance(value, list):
                for valu in value:
                    if valu.strip() == '':
                        continue
                    updcur.execute(sqli, (self.rows[self.row + row], key, valu.strip()))
            else:
                if value.strip() == '':
                    continue
                updcur.execute(sqli, (self.rows[self.row + row], key, value.strip()))
        cur.close()
        updcur.close()
        self.conn.commit()

    def row_click(self, position):
        row = self.table.indexAt(position).row()
        if row < 0:
            return
        msgbox = QtGui.QMessageBox()
        msgbox.setWindowTitle('Delete item')
        msgbox.setText("Press Yes to delete '" + str(self.table.item(row, 2).text()) + "'")
        msgbox.setIcon(QtGui.QMessageBox.Question)
        msgbox.setStandardButtons(QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        reply = msgbox.exec_()
        if reply == QtGui.QMessageBox.Yes:
            cur = self.conn.cursor()
            iid = self.rows[self.row + row]
            sql = "delete from meta where item_id = ?"
            cur.execute(sql, (iid, ))
            sql = "delete from items where id = ?"
            cur.execute(sql, (iid, ))
            cur.execute("select count(*) from items")
            cnt = cur.fetchone()[0]
            if cnt > 0:
                self.items.setText('(' + str(cnt) + ' items)')
            else:
                self.items.setText('')
            cur.close()
            self.conn.commit()
        #    self.table.removeRow(row)
            self.do_search()

    def showAbout(self):
        about = 'A simple catalogue for documents, books, whatever...'
        dialog = displayobject.AnObject(QtGui.QDialog(), about, title='About ' + self.me)
        dialog.exec_()

    def showHelp(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), 'help.html', title='Help for ' + self.me)
        dialog.exec_()

    def header_click(self, position):
        column = self.headers.logicalIndexAt(position)
        self.order(column)

    def quit(self):
        self.close()

    def closeEvent(self, event):
        if len(self.dbs) > 0:
            if self.db is not None:
                self.dbs.insert(0, self.dbs.pop(self.dbs.index(self.db)))
            configFile(data=self.dbs)
        event.accept()

    def findFolder(self, itemid, folder, filename):
        folder = QtGui.QFileDialog.getExistingDirectory(None,
                 'Folder to start searching for ' + filename, folder,
                 QtGui.QFileDialog.ShowDirsOnly)
        if folder != '':
            possibles = []
            folders = [f[0] for f in os.walk(folder)]
            for fldr in folders:
                if os.path.exists(fldr + '/' + filename):
                    possibles.append(fldr)
            del folders
            if len(possibles) == 0:
                self.wrapmsg.setText(filename + ' not found')
                return
            if len(possibles) == 1:
                folder = possibles[0] + '/'
                print(folder)
                if self.translate_user != '':
                    folder = folder.replace(getUser(), self.translate_user)
                sql = "update items set location = ? where id = ?"
                updcur = self.conn.cursor()
                updcur.execute(sql, (folder , itemid))
                updcur.close()
                self.conn.commit()
                os.system(self.launcher + ' "' + possibles[0] + '/' + filename + '"')
                self.wrapmsg.setText(filename + ' found at ' + folder)
                return
            for f in range(len(possibles)):
                stat = os.stat(possibles[f] + '/' + filename)
                possibles[f] = [possibles[f], stat.st_size]
            fields = ['Location', 'Size']
            dialog = displaytable.Table(possibles, fields=fields)
            dialog.exec_()
            folder = dialog.getChoice()
            if folder is not None:
                sql = "update items set location = ? where id = ?"
                updcur = self.conn.cursor()
                if self.translate_user != '':
                    fldr = folder.replace(getUser(), self.translate_user)
                updcur.execute(sql, (fldr + '/', itemid))
                updcur.close()
                self.conn.commit()
                os.system(self.launcher + ' "' + folder + '/' + filename + '"')
                self.wrapmsg.setText(filename + ' found at ' + folder + '/')
                return


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))


if '__main__' == __name__:
    app = QtGui.QApplication(sys.argv)
    tabdialog = TabDialog()
    tabdialog.show()
    sys.exit(app.exec_())
