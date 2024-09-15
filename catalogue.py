#!/usr/bin/python3
#
#  Copyright (C) 2019-2024 Angus King
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
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sqlite3
from sqlite3 import Error
import sys
import tempfile
import time
import webbrowser

import displayobject
import displaytable
from functions import *


class TabDialog(QMainWindow):

    def wheelEvent(self, event):
        if len(self.rows) == 0:
            return
        self.wrapmsg.setText('')
        if event.angleDelta().y() < 0:
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

    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        i = sys.argv[0].rfind('/')
        if i >= 0:
            self.mydir = sys.argv[0][: i + 1]
        else:
            self.mydir = ''
        self.find_file = False
        self.conn = None
        self.db = None
        self.dbs = configFile()
        self.last_locn = '' # remember last location when adding books
        self.cattitle = QLabel('')
        self.items = QLabel('')
        self.metacombo = QComboBox(self)
        self.category = 'Category'
        self.catcombo = QComboBox(self)
        self.catcombo.setHidden(True)
        self.attr_cat = QAction(QIcon(self.mydir + 'copy.png'), self.category, self)
        self.attr_cat.setShortcut('Ctrl+C')
        self.attr_cat.setStatusTip('Edit category values')
        self.attr_cat.triggered.connect(self.editFields)
        if len(self.dbs) > 0:
            self.conn = create_connection(self.dbs[0])
            if self.conn is not None:
                self.db = self.dbs[0]
                self.updDetails()
        cur = self.conn.cursor()
        cur.execute("select field from fields where typ = 'Meta' and description = 'Item Date'")
        try:
            self.item_date = cur.fetchone()[0].title()
        except:
            self.item_date = 'n/a'
        cur.close()
        self.me = '/' + sys.argv[0][:sys.argv[0].rfind('.')]
        self.me = self.me[self.me.rfind('/') + 1:].title()
        self.setWindowTitle(self.me + ' (' + fileVersion() + ') - Simple Catalogue')
        self.setWindowIcon(QIcon(self.mydir + 'books.png'))
        buttonLayout = QHBoxLayout()
        quitButton = QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(quitButton)
        quitButton.clicked.connect(self.quit)
        QShortcut(QKeySequence('q'), self, self.quit)
        addButton = QPushButton(self.tr('&Add Item'))
        buttonLayout.addWidget(addButton)
        addButton.clicked.connect(self.addItem)
        addfile = QPushButton(self.tr('&Add File'))
        buttonLayout.addWidget(addfile)
        addfile.clicked.connect(self.addFile)
        addfiles = QPushButton(self.tr('&Add Files'))
        buttonLayout.addWidget(addfiles)
        addfiles.clicked.connect(self.addFiles)
        self.addisbn = QPushButton(self.tr('&Add ISBN'))
        buttonLayout.addWidget(self.addisbn)
        self.addisbn.clicked.connect(self.addISBN)
        if self.isbn_field == '':
            self.addisbn.setVisible(False)
        self.ignore_expired = QCheckBox('Ignore ' + self.expired_category + '?', self)
        buttonLayout.addWidget(self.ignore_expired)
        if self.expired_category != '':
            self.ignore_expired.setVisible(True)
        QShortcut(QKeySequence('pgdown'), self, self.nextRows)
        QShortcut(QKeySequence('pgup'), self, self.prevRows)
        buttons = QFrame()
        buttons.setLayout(buttonLayout)
        layout = QGridLayout()
        layout.setVerticalSpacing(10)
        layout.addWidget(QLabel('Catalogue:'), 0, 0)
        layout.addWidget(self.cattitle, 0, 1, 1, 2)
        layout.addWidget(self.items, 0, 4)
        layout.addWidget(QLabel('Search by'), 1, 0)
        layout.addWidget(self.metacombo, 1, 1)
        self.search = QLineEdit()
        self.filter = QComboBox(self)
        self.filter.addItem('equals')
        self.filter.addItem('starts with')
        self.filter.addItem('contains')
        self.filter.addItem('missing')
        self.filter.addItem('duplicate')
        self.filter.setCurrentIndex(2)
        layout.addWidget(self.filter, 1, 2)
        layout.addWidget(self.catcombo, 1, 3)
        layout.addWidget(self.search, 1, 3)
        srchb = QPushButton('Search')
        layout.addWidget(srchb, 1, 4)
        srchb.clicked.connect(self.do_search)
        enter = QShortcut(QKeySequence('Return'), self)
        enter.activated.connect(self.do_search)
        layout.addWidget(buttons, 2, 0, 1, 2)
        msgLayout = QHBoxLayout()
        msgLayout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.wrapmsg = QLabel('')
        msg_font = self.wrapmsg.font()
        msg_font.setBold(True)
        self.wrapmsg.setFont(msg_font)
        msg_palette = QPalette()
        msg_palette.setColor(QPalette.Foreground, Qt.red)
        self.wrapmsg.setPalette(msg_palette)
        msgLayout.addWidget(self.wrapmsg)
        msgLayout.addWidget(QLabel('Rows per page:'))
        self.pagerows = QSpinBox()
        msgLayout.addWidget(self.pagerows)
        self.srchrng = QLabel('')
        self.srchrng.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        msgLayout.addWidget(self.srchrng)
        self.srchmsg = QLabel('')
        msgLayout.addWidget(self.srchmsg)
        msgs = QFrame()
        msgs.setLayout(msgLayout)
        layout.addWidget(msgs, 2, 2, 1, 3)
        layout.setColumnStretch(3, 10)
        menubar = QMenuBar()
        layout.setMenuBar(menubar)
        db_opn = QAction(QIcon(self.mydir + 'open.png'), 'Open', self)
        db_opn.setShortcut('Ctrl+O')
        db_opn.setStatusTip('Open Catalogue')
        db_opn.triggered.connect(self.openDB)
        db_add = QAction(QIcon(self.mydir + 'plus.png'), 'Add', self)
        db_add.setShortcut('Ctrl+A')
        db_add.setStatusTip('Add Catalogue')
        db_add.triggered.connect(self.addDB)
        db_new = QAction(QIcon(self.mydir + 'books.png'), 'New', self)
        db_new.setShortcut('Ctrl+N')
        db_new.setStatusTip('Create Catalogue')
        db_new.triggered.connect(self.newDB)
        db_lod = QAction(QIcon(self.mydir + 'load.png'), 'Load', self)
        db_lod.setShortcut('Ctrl+L')
        db_lod.setStatusTip('Load Catalogue')
        db_lod.triggered.connect(self.loadDB)
        db_xpt = QAction(QIcon(self.mydir + 'save.png'), 'Export', self)
        db_xpt.setStatusTip('Export Catalogue')
        db_xpt.triggered.connect(self.exportDB)
        db_rem = QAction(QIcon(self.mydir + 'cancel.png'), 'Remove', self)
        db_rem.setStatusTip('Remove Catalogue')
        db_rem.triggered.connect(self.remDB)
        db_qui = QAction(QIcon(self.mydir + 'quit.png'), 'Quit', self)
        db_qui.setShortcut('Ctrl+Q')
        db_qui.setStatusTip('Quit')
        db_qui.triggered.connect(self.quit)
        dbMenu = menubar.addMenu('&Catalogue')
        dbMenu.addAction(db_opn)
        dbMenu.addAction(db_add)
        dbMenu.addAction(db_new)
        dbMenu.addAction(db_lod)
        dbMenu.addAction(db_xpt)
        dbMenu.addAction(db_rem)
        dbMenu.addAction(db_qui)
        attrMenu = menubar.addMenu('&Attributes')
        attr_inf = QAction(QIcon(self.mydir + 'edit.png'), 'Info', self)
        attr_inf.setShortcut('Ctrl+E')
        attr_inf.setStatusTip('Edit Catalogue info')
        attr_inf.triggered.connect(self.editFields)
        attr_met = QAction(QIcon(self.mydir + 'list.png'), 'Meta', self)
        attr_met.setShortcut('Ctrl+M')
        attr_met.setStatusTip('Edit Meta values')
        attr_met.triggered.connect(self.editFields)
        attr_set = QAction(QIcon(self.mydir + 'edit.png'), 'Settings', self)
        attr_set.setShortcut('Ctrl+S')
        attr_set.setStatusTip('Edit Catalogue setting')
        attr_set.triggered.connect(self.editFields)
        attrMenu.addAction(self.attr_cat)
        attrMenu.addAction(attr_inf)
        attrMenu.addAction(attr_met)
        attrMenu.addAction(attr_set)
        help = QAction(QIcon(self.mydir + 'help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        about = QAction(QIcon(self.mydir + 'info.png'), 'About', self)
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
        self.table = QTableWidget()
        self.table.setRowCount(1)
        self.table.setItem(0, 0, QTableWidgetItem(''))
        rh = self.table.rowHeight(0)
        screen = QDesktopWidget().availableGeometry()
        table_rows = int((screen.height() - toplen) / rh)
        self.pagerows.setRange(1, table_rows)
        self.pagerows.setValue(table_rows)
        self.table.setRowCount(table_rows)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(['', '', 'Title', self.category, self.item_date])
        self.table.setColumnWidth(0, 20)
        self.table.setColumnWidth(1, 0)
        self.table.setColumnWidth(2, 800)
        self.table.setColumnWidth(3, 200)
        self.table.setColumnWidth(4, 100)
        self.table.setEditTriggers(QAbstractItemView.SelectedClicked)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
  #      self.table.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Expanding)
        rows = self.table.verticalHeader()
        rows.setContextMenuPolicy(Qt.CustomContextMenu)
        rows.customContextMenuRequested.connect(self.row_click)
        rows.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.verticalHeader().setVisible(True)
        self.table.cellClicked.connect(self.item_selected)
        layout.addWidget(self.table, 3, 0, 1, 5)
        centralWidget = QWidget(self)
        size = QSize(screen.width() - 50, screen.height() - 50)
        centralWidget.setMaximumSize(size)
        self.setCentralWidget(centralWidget)
        centralWidget.setLayout(layout)
        self.metacombo.currentIndexChanged.connect(self.metaChanged)
        self.catcombo.currentIndexChanged.connect(self.catChanged)
        size = QSize(min(screen.width() - 50, 1200), screen.height() - 50)
        self.resize(size)
        goto_top = QShortcut(QKeySequence('Ctrl+Home'), self)
        goto_top.activated.connect(lambda: self.nextRows(top=True))
        goto_end = QShortcut(QKeySequence('Ctrl+End'), self)
        goto_end.activated.connect(lambda: self.prevRows(bottom=True))
        delete = QShortcut(QKeySequence('Ctrl+Del'), self)
        delete.activated.connect(lambda: self.delete_items())
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
            self.search.setText('')
            self.do_search()

    def addDB(self):
        db_file = 'catalogue.db'
        db_file = QFileDialog.getOpenFileName(None, 'Add Catalogue',
                  db_file, 'Database Files (*.db)')[0]
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
                msgbox = QMessageBox()
                msgbox.setText('Do you want to DELETE the database (Y)?')
                msgbox.setIcon(QMessageBox.Question)
                msgbox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                reply = msgbox.exec_()
                if reply == QMessageBox.Yes:
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
        db_file = QFileDialog.getSaveFileName(None, 'Create Catalogue',
                  db_file, 'Database Files (*.db)')[0]
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
        self.table.clear()
        self.table.setHorizontalHeaderLabels(['', self.field, 'Title', self.category, self.item_date])
        self.rows = 0
        if db_file not in self.dbs:
            self.dbs.append(db_file)
        self.wrapmsg.setText('Catalogue created')
        self.loadDB()

    def loadDB(self):
        fldr = self.db[:self.db.rfind('/')]
        data_file = QFileDialog.getOpenFileName(None, 'Choose data to load',
                    fldr, 'Data Files (*.xls* *.csv)')[0]
        if data_file == '':
            return
        message = load_catalogue(self, self.db, data_file)
        self.wrapmsg.setText(message)
        self.updDetails()
        self.do_search()

    def exportDB(self):
        fldr = self.db[:self.db.rfind('/') + 1] + self.metacombo.currentText() + '_' \
               + self.filter.currentText() + '_' + self.search.text() + '.csv'
        data_file = QFileDialog.getSaveFileName(None, 'Choose export file',
                    fldr, 'Data Files (*.xls* *.csv)')[0]
        if data_file == '':
            return
        message = export_catalogue(self, self.db, data_file, self.rows)
        self.wrapmsg.setText(message)

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
        self.location_list = False
        self.translate_user = '$USER$'
        self.category_multi = False
        self.url_field = ''
        self.isbn_field = ''
        self.item_date = ''
        self.dewey_field = ''
        self.default_category = '?'
        self.expired_category = ''
        self.set_ignore_expired = False
        cur.execute("select field, description from fields where typ = 'Settings'")
        row = cur.fetchone()
        while row is not None:
            if row[0] == 'Category Choice':
                if row[1][:5].lower() == 'multi':
                    self.category_multi = True
            elif row[0] == 'Category Field':
                self.category = row[1].title()
            elif row[0] == 'Decrypt PDF':
                self.pdf_decrypt = row[1]
            elif row[0] == 'Default Category':
                self.default_category = row[1].title()
            elif row[0] == 'Dewey Field':
                self.dewey_field = row[1].title()
            elif row[0] == 'Expired Category':
                self.expired_category = row[1].title()
            elif row[0] == 'Ignore Expired':
                if row[1].lower() in ['true', 'on', 'yes']:
                    self.set_ignore_expired = True
            elif row[0] == 'ISBN Field':
                self.isbn_field = row[1].upper()
            elif row[0] == 'Item Date Field':
                self.item_date = row[1]
            elif row[0] == 'Launch File':
                self.launcher = row[1]
            elif row[0] == 'Location Choice':
                if row[1].lower() == 'list':
                    self.location_list = True
            elif row[0] == 'Translate Userid':
                self.translate_user = row[1]
            elif row[0] == 'URL Field':
                self.url_field = row[1].upper()
            row = cur.fetchone()
        try:
            if self.isbn_field != '':
                 self.addisbn.setVisible(True)
            else:
                 self.addisbn.setVisible(False)
        except:
            pass
        try:
            if self.expired_category != '':
                self.ignore_expired.setText('Ignore ' + self.expired_category + '?')
                self.ignore_expired.setVisible(True)
                if self.set_ignore_expired:
                    self.ignore_expired.setCheckState(Qt.Checked)
                else:
                    self.ignore_expired.setCheckState(Qt.Unchecked)
            else:
                self.ignore_expired.setVisible(False)
        except:
            pass
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
            self.items.setText('({:,d} items)'.format(cnt))
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
            if row[0] == 'Keyword':
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
        folder = self.db[:self.db.rfind('/') + 1]
        dialog = displaytable.Table(rows, fields=fields, title=self.sender().text(), edit=True,
                                    save_folder=folder)
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
        if self.sender().text() == 'Settings':
            set_field = ''
            set_default = ''
            set_expired = ''
            try:
                set_field = upd_rows['Category Field']
                try:
                    set_default = upd_rows['Default Category']
                except:
                    pass
                try:
                    set_expired = upd_rows['Expired Category']
                except:
                    pass
            except:
                pass
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
                    msgbox = QMessageBox()
                    msgbox.setWindowTitle('Delete ' + self.sender().text() + ' value')
                    msgbox.setText(str(i) + ' rows for ' + row[0] \
                                   + '. Do you still want to delete it?')
                    msgbox.setIcon(QMessageBox.Question)
                    msgbox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                    reply = msgbox.exec_()
                if i == 0 or reply == QMessageBox.Yes:
                   sql = "delete from fields where typ = ? and field = ?"
                   updcur.execute(sql, (self.sender().text(), row[0]))
                   if i > 0:
                       msgbox.setText('Do you want to delete the Meta rows?')
                       reply = msgbox.exec_()
                       if reply == QMessageBox.Yes:
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
        if self.sender().text() == 'Settings' and set_field != '':
            if set_default != '':
                sql = "select field from fields where typ = ? and field = ?"
                cur.execute(sql, (set_field, set_default))
                if row is None:
                    sql = "insert into fields (typ, field, description) values (?, ?, ?)"
                    updcur.execute(sql, (set_field, set_default, 'Default Category'))
            if set_expired != '':
                sql = "select field from fields where typ = ? and field = ?"
                cur.execute(sql, (set_field, set_expired))
                if row is None:
                    sql = "insert into fields (typ, field, description) values (?, ?, ?)"
                    updcur.execute(sql, (set_field, set_expired, 'Expired Category'))
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
        new_file = QFileDialog.getOpenFileName(None, 'Add File',
                   self.db[:self.db.rfind('/')], 'All Files (*.*)')[0]
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
            properties = getPDFInfo(new_file, properties=properties, decrypt=self.pdf_decrypt,
                                    translate_user=self.translate_user)
        self.addItem(properties=properties)

    def addFiles(self):
        if self.conn is None:
            return
        folder = QFileDialog.getExistingDirectory(None,
                 'Folder to start searching for files',
                 self.db[:self.db.rfind('/')], QFileDialog.ShowDirsOnly)
        if folder == '':
            return
        filetypes = ['pdf', 'html', 'doc', 'docx', 'xls', 'xlsx']
        cur = self.conn.cursor()
        cur.execute("select description from fields where typ = 'Settings'" + \
                    " and field like 'File Type%'")
        row = cur.fetchone()
        while row is not None:
            typs = row[0].split(' ')
            for typ in typs:
                if typ.lower() not in filetypes:
                    filetypes.append(typ.lower())
            row = cur.fetchone()
        add_limit = 20
        cur.execute("select description from fields where typ = 'Settings'" + \
                    " and field == 'Add Limit'")
        row = cur.fetchone()
        while row is not None:
            try:
                add_limit = int(row[1])
            except:
                pass
            row = cur.fetchone()
        possibles = []
        sql = "select location from items where filename = ?"
        for top, dirs, files in os.walk(folder):
            for name in files:
                typ = name[name.rfind('.') + 1:]
                if typ.lower() not in filetypes:
                    continue
                cur.execute(sql, (name, ))
                row = cur.fetchone()
                if row is None:
                    possibles.append([top, name])
            if len(possibles) >= add_limit:
                break
        if len(possibles) == 0:
            self.wrapmsg.setText('No files to add')
            return
        selected = whatFiles(possibles, self.launcher)
        selected.exec_()
        selected = selected.getValues()
        if selected is None:
            return
        for item in selected:
            properties = {}
            item[0] = item[0] + '/'
            properties['Title'] = item[1]
            properties['Filename'] = item[1]
            stat = os.stat(item[0] + item[1])
            properties['Filesize'] = str(stat.st_size)
            if self.translate_user != '':
                item[0] = item[0].replace(getUser(), self.translate_user)
            properties['Location'] = item[0]
            properties['Acquired'] = time.strftime('%Y-%m-%d', time.localtime(stat.st_mtime))
            if item[1][-4:].lower() == '.pdf':
                properties = getPDFInfo(item[0] + item[1], properties=properties,
                                        decrypt=self.pdf_decrypt,
                                        translate_user=self.translate_user)
            self.addItem(properties=properties)
        return

    def addISBN(self):
        if self.conn is None:
            return
        isbn, ok = QInputDialog.getText(None, 'Get ISBN Details', 'Enter ISBN:',
                   QLineEdit.Normal, '978')
        if not ok:
            return
        cur = self.conn.cursor()
        sql = "select (item_id) from meta where field = ? and value = ?"
        cur.execute(sql, (self.isbn_field, isbn))
        row = cur.fetchone()
        cur.close()
        if row is not None:
            msgbox = QMessageBox()
            msgbox.setWindowTitle('Add ISBN')
            msgbox.setText('ISBN ' + isbn + ' in Catalogue. Press Yes to add a duplicate')
            msgbox.setIcon(QMessageBox.Question)
            msgbox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            reply = msgbox.exec_()
            if reply != QMessageBox.Yes:
                return
        properties = getISBNInfo(isbn, self.conn)
        if len(properties) > 1: # got some stuff
            try:
                properties['Notes'] = properties['Notes'] + '\n(info derived from openlibrary.org)'
            except:
                properties['Notes'] = '(info derived from openlibrary.org)'
        else:
            QApplication.clipboard().setText(isbn)
        self.addItem(properties=properties)

    def addItem(self, properties=None):
        if self.conn is None:
            return
        addproperty = {}
        fields = ['Title', 'Filename', 'Location']
        for field in fields:
            if properties and field in properties.keys():
                addproperty[field] = properties[field]
            else:
                addproperty[field] = ''
        if addproperty['Location'] == '' and self.last_locn != '':
            addproperty['Location'] = self.last_locn
        cur = self.conn.cursor()
        cur.execute("select field from fields where typ = 'Meta' order by field")
        row = cur.fetchone()
        add_cat_field = True
        while row is not None:
            if row[0] == self.category:
                add_cat_field = False
            if properties and row[0] in properties.keys():
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
        if self.location_list:
            locnlist = ['Location', []]
            cur.execute("select distinct location from items where filename = '' order by location")
            row = cur.fetchone()
            while row is not None:
                locnlist[1].append(row[0])
                row = cur.fetchone()
        else:
            locnlist = None
        if self.default_category != '':
            addproperty[self.category] = self.default_category
        dialog = displayobject.AnObject(QDialog(), addproperty, readonly=False,
                 textedit=True, title='Add Item', combolist=combolist, multi=self.category_multi,
                 locnlist=locnlist, default=self.default_category)
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
            self.items.setText('({:,d} items)'.format(cnt))
        else:
            self.items.setText('')
        cur.close()
        self.conn.commit()
        self.last_locn = dialog.getValues()['Location']
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
            self.filter.setCurrentIndex(2) # contains
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
    #    self.wrapmsg.setText('')
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
        no_expired = []
        if self.ignore_expired.isChecked() and self.expired_category != '':
            cur = self.conn.cursor()
            sql = "select (item_id) from meta where field = ? and value = ?"
            cur.execute(sql, (self.category, self.expired_category))
            row = cur.fetchone()
            while row is not None:
                no_expired.append(row[0])
                row = cur.fetchone()
            cur.close()
        cur = self.conn.cursor()
        if self.field == 'All':
            sql = "select (id) from items where title " + where + \
                  " or filename " + where + \
                  " or location " + where + \
                  " or id in (select (item_id) from meta where value " + where + ")" + \
                  " order by title"
            cur.execute(sql, (search, search, search, search))
        elif self.filter.currentText() == 'duplicate':
            if self.field in ['Title', 'Filename', 'Location']:
                if search != '':
                    sql = "select o.id from items o inner join (select " + self.field + \
                      " from items where " + self.field + " " + where + " group by " + \
                      self.field + " having count(" + self.field + ") > 1) i on o." + \
                      self.field + " = i." + self.field + " order by o." + self.field
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
            elif self.field == self.category:
                sql = "select (id) from items where id in (" + sql + ") order by title"
            cur.execute(sql, (self.field, search))
        row = cur.fetchone()
        if self.find_file:
            if row is not None:
                self.wrapmsg.setText('Files missing. Click Launch icon(s) to search for new location.')
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
                if row[0] not in self.rows and row[0] not in no_expired:
                    self.rows.append(row[0])
                row = cur.fetchone()
        cur.close()
        self.srchmsg.setText('{:,d} items'.format(len(self.rows)))
        self.srchrng.setText('')
        self.table.clear()
        self.table.setHorizontalHeaderLabels(['', self.field, 'Title', self.category, self.item_date])
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
        self.table.setHorizontalHeaderLabels(['', self.field, 'Title', self.category, self.item_date])
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
            if row[1] == '':
                if self.url_field != '':
                    # need to check here if they have a url
                    cur.execute(sql1, (self.rows[self.row + trw], self.url_field))
                    urow = cur.fetchone()
                    if urow is not None:
                        button = QPushButton('', self.table)
                        button.setIcon(QIcon(self.mydir + 'url.png'))
                        button.setFlat(True)
                        self.table.setCellWidget(trw, 0, button)
                        button.clicked.connect(partial(self._buttonItemClicked, trw))
            elif self.launcher != '':
                button = QPushButton('', self.table)
                button.setIcon(QIcon(self.mydir + 'open.png'))
                button.setFlat(True)
                self.table.setCellWidget(trw, 0, button)
                button.clicked.connect(partial(self._buttonItemClicked, trw))
            self.table.setItem(trw, 2, QTableWidgetItem(row[0]))
            cur.execute(sql1, (self.rows[self.row + trw], self.category))
            txt = ''
            row = cur.fetchone()
            while row is not None:
                txt += ';' + row[0]
                row = cur.fetchone()
            txt = txt[1:]
            self.table.setItem(trw, 3, QTableWidgetItem(txt))
            if self.col_2:
                if self.field in ['Filename', 'Location']:
                    cur.execute(sql2, (self.rows[self.row + trw], ))
                else:
                    cur.execute(sql1, (self.rows[self.row + trw], self.field))
                row = cur.fetchone()
                if row is not None:
                    self.table.setItem(trw, 1, QTableWidgetItem(row[0]))
            cur.execute(sql1, (self.rows[self.row + trw], self.item_date))
            row = cur.fetchone()
            if row is not None:
                self.table.setItem(trw, 4, QTableWidgetItem(row[0]))
        self.srchrng.setText('{:,d} to  {:,d} of'.format(self.row + 1, self.row + 1 + trw))
        cur.close()

    def nextRows(self, top=False):
        if len(self.rows) < self.pagerows.value():
                return
        self.wrapmsg.setText('')
        if top:
            self.row = 0
        else:
            self.row += self.pagerows.value()
            if self.row >= len(self.rows):
                self.row = 0
                self.wrapmsg.setText('Wrapped to top')
        self.getRows()

    def prevRows(self, bottom=False):
        if len(self.rows) < self.pagerows.value():
            return
        self.wrapmsg.setText('')
        if bottom:
            self.row = len(self.rows) - self.pagerows.value()
        elif self.row - self.pagerows.value() < 0:
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
        if self.translate_user != '':
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
        if itmproperty['Filename'] == '' and self.location_list:
            locnlist = ['Location', []]
            cur.execute("select distinct location from items where filename = '' order by location")
            lrow = cur.fetchone()
            while lrow is not None:
                locnlist[1].append(lrow[0])
                lrow = cur.fetchone()
        else:
            locnlist = None
        cur.close()
        dialog = displayobject.AnObject(QDialog(), itmproperty, readonly=False,
                 textedit=True, title='Edit Item (' + str(self.rows[self.row + row]) + ')',
                 combolist=combolist, multi=self.category_multi, locnlist=locnlist,
                 default=self.default_category)
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
                self.table.setItem(row, 3, QTableWidgetItem(txt))
            else:
                self.table.setItem(row, 3, QTableWidgetItem(dialog.getValues()[self.category]))
            self.table.setItem(row, 2, QTableWidgetItem(dialog.getValues()['Title']))
            if self.field != '' and self.field not in fields:
                self.table.setItem(row, 1, QTableWidgetItem(dialog.getValues()[self.field]))
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
        msgbox = QMessageBox()
        msgbox.setWindowTitle('Delete item')
        msgbox.setText("Press Yes to delete '" + str(self.table.item(row, 2).text()) + "'")
        msgbox.setIcon(QMessageBox.Question)
        msgbox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        reply = msgbox.exec_()
        if reply == QMessageBox.Yes:
            cur = self.conn.cursor()
            iid = self.rows[self.row + row]
            sql = "delete from meta where item_id = ?"
            cur.execute(sql, (iid, ))
            sql = "delete from items where id = ?"
            cur.execute(sql, (iid, ))
            cur.execute("select count(*) from items")
            cnt = cur.fetchone()[0]
            if cnt > 0:
                self.items.setText('({:,d} items)'.format(cnt))
            else:
                self.items.setText('')
            cur.close()
            self.conn.commit()
        #    self.table.removeRow(row)
            self.do_search()

    def delete_items(self):
        delrows = []
        for item in self.table.selectedItems():
            if item.row() not in delrows and self.row + item.row() < len(self.rows):
                delrows.append(item.row())
        if len(delrows) == 0:
            return
        msgbox = QMessageBox()
        msgbox.setWindowTitle('Delete items')
        msgbox.setText("Press Yes to delete " + str(len(delrows)) + " items.")
        msgbox.setIcon(QMessageBox.Question)
        msgbox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        reply = msgbox.exec_()
        if reply == QMessageBox.Yes:
            sqlm = "delete from meta where item_id = ?"
            sqli = "delete from items where id = ?"
            cur = self.conn.cursor()
            for row in delrows:
                iid = self.rows[self.row + row]
                cur.execute(sqlm, (iid, ))
                cur.execute(sqli, (iid, ))
            cur.execute("select count(*) from items")
            cnt = cur.fetchone()[0]
            if cnt > 0:
                self.items.setText('({:,d} items)'.format(cnt))
            else:
                self.items.setText('')
            cur.close()
            self.conn.commit()
            self.wrapmsg.setText(str(len(delrows)) + ' items deleted.')
            self.do_search()

    def showAbout(self):
        about = '<html>' + \
                '<h2>Catalogue</h2>' + \
                '<p>A simple catalogue for documents, books, whatever...</p>\n' + \
                '<p>Copyright © 2019-2023 Angus King</p>\n' + \
                '<p>This program is free software: you can redistribute it and/or modify\n' + \
                ' it under the terms of the GNU Affero General Public License as published\n' + \
                ' by the Free Software Foundation, either version 3 of the License, or\n' + \
                ' (at your option) any later version.</p>\n' + \
                '<p>This program is distributed in the hope that it will be useful, but WITHOUT\n' + \
                ' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or\n' + \
                ' FITNESS FOR A PARTICULAR PURPOSE. See the GNU Affero General Public\n' + \
                ' License for more details.</p>\n' + \
                '<p>You should have received a copy of the GNU Affero General Public License\n' + \
                ' along with this program. If not, see <a href="http://www.gnu.org/licenses/">\n' + \
                'http://www.gnu.org/licenses/</a></html>'
        dialog = displayobject.AnObject(QDialog(), about, title='About ' + self.me)
        dialog.exec_()

    def showHelp(self):
        dialog = displayobject.AnObject(QDialog(), 'catalogue.html', title='Help for ' + self.me)
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
        folder = QFileDialog.getExistingDirectory(None,
                 'Folder to start searching for ' + filename, folder,
                 QFileDialog.ShowDirsOnly)
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


class whatFiles(QDialog):
    def __init__(self, files, launcher=''):
        super(whatFiles, self).__init__()
        self.files = files
        self.launcher = launcher
        self.chosen = []
        self.grid = QGridLayout()
        self.checkbox = []
        self.checkbox.append(QCheckBox('Check / Uncheck all', self))
        self.grid.addWidget(self.checkbox[-1], 0, 0)
        i = 0
        c = 0
        for fil in range(len(self.files)):
            self.checkbox.append(QCheckBox(self.files[fil][1]))
            self.checkbox[-1].setObjectName(str(fil))
            i += 1
            self.grid.addWidget(self.checkbox[-1], i, c)
            if i > 25:
                i = 0
                c += 1
        self.checkbox[0].stateChanged.connect(self.check_all)
        show = QPushButton('Choose', self)
        self.grid.addWidget(show, i + 1, c)
        show.clicked.connect(self.showClicked)
        self.setLayout(self.grid)
        self.setWindowTitle('Select files to add - Right-click to open')
        QShortcut(QKeySequence('q'), self, self.quitClicked)
        self.show_them = False
        self.show()

    def check_all(self):
        if self.checkbox[0].isChecked():
            for i in range(len(self.checkbox)):
                self.checkbox[i].setCheckState(Qt.Checked)
        else:
            for i in range(len(self.checkbox)):
                self.checkbox[i].setCheckState(Qt.Unchecked)

    def closeEvent(self, event):
        if not self.show_them:
            self.chosen = None
        event.accept()

    def quitClicked(self):
        self.close()

    def showClicked(self):
        for fil in range(1, len(self.checkbox)):
            if self.checkbox[fil].checkState() == Qt.Checked:
                self.chosen.append(self.files[fil - 1])
        self.show_them = True
        self.close()

    def getValues(self):
        return self.chosen

    def mousePressEvent(self, QMouseEvent):
        if self.launcher == '':
            return
        cursor = QCursor()
        widget = qApp.widgetAt(cursor.pos())
        fil = int(widget.objectName())
        if os.path.exists(self.files[fil][0] + '/' + self.files[fil][1]):
            os.system(self.launcher + ' "' + self.files[fil][0] + '/' + self.files[fil][1] + '"')


if '__main__' == __name__:
    app = QApplication(sys.argv)
    tabdialog = TabDialog()
    tabdialog.show()
    sys.exit(app.exec_())
