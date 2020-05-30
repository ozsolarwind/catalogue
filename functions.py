#!/usr/bin/python3
#
#  Copyright (C) 2019 Angus King
#
#  functions.py - This file is part of catalogue.
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

import csv
from datetime import datetime
import json
import http.client
import os
import sqlite3
from sqlite3 import Error
import sys
import tempfile
from PyPDF2 import PdfFileReader
if sys.platform == 'win32' or sys.platform == 'cygwin':
    from win32api import GetFileVersionInfo, LOWORD, HIWORD
import pyexcel as pxl

import displaytable

def configFile(data=None):
    if sys.platform == 'win32' or sys.platform == 'cygwin':
        try:
            config = '~\\.catalogue.data'.replace('~', os.environ['HOME'])
        except:
            config = '.catalogue.data'
    else:
        config = '~/.user/catalogue.data'.replace('~', os.path.expanduser('~'))
    usr = getUser()
    if data is None:
        dbs = []
        if os.path.exists(config):
            cf = open(config, 'r')
            lines = cf.readlines()
            cf.close()
            for line in lines:
                lin = line.strip()
                linu = lin.replace('$USER$', usr)
                if linu not in dbs:
                    dbs.append(linu)
            return dbs
        else:
            return []
    else:
        cf = open(config, 'w')
        for db in data:
            i = db.find(usr)
            if i >= 0:
                db = db.replace(usr, '$USER$')
            cf.write(db + '\n')
        cf.close()

def fileVersion(program=None, year=False):
    ver = '?'
    ver_yr = '????'
    if program == None:
        check = sys.argv[0]
    else:
        s = program.rfind('.')
        if s > 0 and program[s:] == '.html':
            check = program
        elif s < len(program) - 4:
            check = program + sys.argv[0][sys.argv[0].rfind('.'):]
        else:
            check = program
    if check[-3:] == '.py':
        try:
            modtime = datetime.fromtimestamp(os.path.getmtime(check))
            ver = '0.1.%04d.%d%02d' % (modtime.year, modtime.month, modtime.day)
            ver_yr = '%04d' % modtime.year
        except:
            pass
    elif check[-5:] == '.html':
        try:
            modtime = datetime.fromtimestamp(os.path.getmtime(check))
            ver = '0.1.%04d.%d%02d' % (modtime.year, modtime.month, modtime.day)
            ver_yr = '%04d' % modtime.year
        except:
            pass
    else:
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                if check.find('\\') >= 0:  # if full path
                    info = GetFileVersionInfo(check, '\\')
                else:
                    info = GetFileVersionInfo(os.getcwd() + '\\' + check, '\\')
                ms = info['ProductVersionMS']
              #  ls = info['FileVersionLS']
                ls = info['ProductVersionLS']
                ver = str(HIWORD(ms)) + '.' + str(LOWORD(ms)) + '.' + str(HIWORD(ls)) + '.' + str(LOWORD(ls))
                ver_yr = str(HIWORD(ls))
            except:
                try:
                    info = os.path.getmtime(os.getcwd() + '\\' + check)
                    ver = '0.1.' + datetime.datetime.fromtimestamp(info).strftime('%Y.%m%d')
                    ver_yr = datetime.datetime.fromtimestamp(info).strftime('%Y')
                    if ver[9] == '0':
                        ver = ver[:9] + ver[10:]
                except:
                    pass
    if year:
        return ver_yr
    else:
        return ver

def create_connection(db_file, create=False):
    """ create a database connection to the SQLite database
        specified by db_file
    :param db_file: database file
    :return: Connection object or None
    """
    if not create:
        if not os.path.exists(db_file):
            return None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)
    return None

def create_table(conn, create_table_sql):
    """ create a table from the create_table_sql statement
    :param conn: Connection object
    :param create_table_sql: a CREATE TABLE statement
    :return:
    """
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
    except Error as e:
        print(e)
    return

def insert_field(conn, typ, field, description):
    """
    Insert a new field into the fields table
    :param conn:
    :param field:
    :param description
    :return: row
    """
    sql = ''' INSERT INTO fields(typ, field, description)
              VALUES(?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, (typ, field, description))
    return

def create_catalogue(database, category='Category'):
    sql_create_fields_table = """ CREATE TABLE IF NOT EXISTS fields (
                                        id integer PRIMARY KEY,
                                        typ text NOT NULL COLLATE NOCASE,
                                        field text NOT NULL COLLATE NOCASE,
                                        description text COLLATE NOCASE
                                    ); """

    sql_create_items_table = """CREATE TABLE IF NOT EXISTS items (
                                    id integer PRIMARY KEY,
                                    filename text COLLATE NOCASE,
                                    location text COLLATE NOCASE,
                                    title text NOT NULL COLLATE NOCASE
                                );"""

    sql_create_meta_table = """CREATE TABLE IF NOT EXISTS meta (
                                    id integer PRIMARY KEY,
                                    item_id integer NOT NULL,
                                    field text NOT NULL COLLATE NOCASE,
                                    value text NOT NULL COLLATE NOCASE,
                                    FOREIGN KEY (item_id) REFERENCES items (id)
                                );"""
    # create a database connection
    conn = create_connection(database, create=True)
    if conn is not None:
        # create items table
        create_table(conn, sql_create_items_table)
        # create fields table
        create_table(conn, sql_create_fields_table)
        # create meta table
        create_table(conn, sql_create_meta_table)
        i = database.rfind('/')
        if i > 0:
            insert_field(conn, 'Info', 'Catalogue', database[i + 1:])
        else:
            insert_field(conn, 'Info', 'Catalogue', database)
        insert_field(conn, 'Info', 'Description', 'Description for ' + database)
        insert_field(conn, 'Meta', 'Keyword', '')
        insert_field(conn, 'Settings', 'Category Choice', 'Single')
        insert_field(conn, 'Settings', 'Category Field', category)
        insert_field(conn, 'Settings', 'Translate Userid', '$USER$')
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            insert_field(conn, 'Settings', 'Launch File', 'start ""')
        else:
            insert_field(conn, 'Settings', 'Decrypt PDF', 'qpdf')
            insert_field(conn, 'Settings', 'Launch File', '/usr/bin/xdg-open')
        conn.commit()
        return conn
    else:
        print("Error! cannot create the database connection.")
    return conn

def load_catalogue(him, database, datafile):
    if not os.path.exists(database):
        return 'Database not found'
    if not os.path.exists(datafile):
        return 'Data file not found'
    items = pxl.get_records(file_name=datafile)
    if len(items) == 0:
        return 'No items to load'
    # get column names
    hdrs = []
    cat_not_defined = False
    conn = create_connection(database)
    cur = conn.cursor()
    cur.execute("select description from fields where typ = 'Settings' and field = 'Category Field'")
    try:
        fields = [cur.fetchone()[0].title()]
    except:
        cat_not_defined = True
        fields = ['Category']
    fields = fields + ['title', 'filename', 'location']
    fldkey = ['', '', '', '']
    for f in range(len(fields)):
        for key in items[0].keys():
            if key.lower() == fields[f].lower():
                fldkey[f] = key
                break
        else:
            try:
                dialog = displaytable.Table(list(items[0].keys()), fields=['attribute'],
                         title='Choose ' + fields[f].title() + ' column')
                dialog.exec_()
                fld = dialog.getChoice()
                if fld is None:
                    raise Exception('Invalid field')
                fldkey[f] = fld
                if f == 0:
                    sqlc = "select count(*) from meta where field = ?"
                    cur.execute(sqlc, (fields[0], ))
                    if cur.fetchone()[0] == 0:
                        fields[0] = fld.title()
                        if not cat_not_defined:
                            sqlu = "update fields set description = ? where typ = 'Settings' and field = 'Category Field'"
                            cur.execute(sqlu, (fields[0], ))
            except:
                if f < 2:
                    cur.close()
                    conn.close()
                    return 'Invalid field selected for ' + fields[f].title()
    cats = []
    sqlc = "select count(*) from fields where typ = 'Meta' and field = ?"
    sql = "insert into fields (typ, field) values ('Meta', ?)"
    if cat_not_defined:
        fields[0] = fldkey[0].title()
        sqls = "insert into fields (typ, field, description) values ('Settings', 'Category Field', ?)"
        cur.execute(sqls, (fields[0], ))
        cur.execute(sqlc, (fields[0], ))
        if cur.fetchone()[0] == 0:
            cur.execute(sql, (fields[0].title(), ))
    for key in items[0].keys():
        if key in fldkey:
            pass
        else:
            cur.execute(sqlc, (key, ))
            if cur.fetchone()[0] == 0:
                cur.execute(sql, (key.title(), ))
    sql = 'insert into items (title, filename, location) values (?, ?, ?)'
    sqlr = 'select last_insert_rowid()'
    sqli = 'insert into meta (field, item_id, value) values (?, ?, ?)'
    ctr = 0
    multi = False
    for item in items:
        if item[fldkey[1]] == '':
            continue
        values = ['?', '', '', '']
        for col in range(len(fldkey)):
            if fldkey[col] == '':
                pass
            elif item[fldkey[col]] != '':
                values[col] = item[fldkey[col]]
        valu = values[0].split(';')
        if len(valu) > 1:
            multi = True
        for val in valu:
            if val not in cats:
                cats.append(val)
        cur.execute(sql, (values[1], values[2], values[3]))
        cur.execute(sqlr)
        iid = cur.fetchone()[0]
        for key, value in item.items():
            if key in fldkey[1:]:
                pass
            elif key == fldkey[0]:
                if item[key] == '':
                    cur.execute(sqli, (fields[0].title(), iid, '?'))
                else:
                    cur.execute(sqli, (fields[0].title(), iid, value))
            elif value != '':
                    cur.execute(sqli, (key.title(), iid, value))
        ctr += 1
    if multi:
        sqlu = "update fields set description = 'Multi' where typ = 'Settings' and field = 'Category Choice'"
        cur.execute(sqlu)
    sqlc = "select count(*) from fields where typ = ? and field = ?"
    sql = "insert into fields (typ, field) values (?, ?)"
    for cat in cats:
        cur.execute(sqlc, (fields[0] , cat))
        if cur.fetchone()[0] == 0:
            cur.execute(sql, (fields[0].title(), cat))
    cur.close()
    conn.commit()
    conn.close()
    pxl.free_resources()
    return '{:,d} items loaded to {}'.format(ctr, database[database.rfind('/') + 1:])

def export_catalogue(him, database, datafile, rows):
    if not os.path.exists(database):
        return 'Database not found'
    conn = create_connection(database)
    cur = conn.cursor()
    cur.execute("select description from fields where typ = 'Settings' and field = 'Category Field'")
    try:
        fields = [cur.fetchone()[0].title()]
    except:
        fields = ['Category']
    fields = fields + ['title', 'filename', 'location']
    cur.execute("select field from fields where typ = 'Meta' order by field")
    row = cur.fetchone()
    while row is not None:
        if row[0] != fields[0]:
            fields.append(row[0])
        row = cur.fetchone()
    sql = "select title, filename, location from items where id = ?"
    sql2 = "select field, value from meta where item_id = ?"
    items = []
    for iid in rows:
        item = {}
        cur.execute(sql, (iid, ))
        row = cur.fetchone()
        item['title'] = row[0]
        item['filename'] = row[1]
        item['location'] = row[2]
        cats = []
        cur.execute(sql2, (iid, ))
        row = cur.fetchone()
        while row is not None:
            if row[0] == fields[0]:
                cats.append(row[1])
            else:
                item[row[0]] = row[1]
            row = cur.fetchone()
        try:
            cat = cats[0]
            for c in range(1, len(cats)):
                cat += ';' + cats[c]
        except:
            cat = '?'
        item[fields[0]] = cat
        for key in fields:
            if key in item.keys():
                pass
            else:
                item[key] = ''
        items.append(item)
    cur.close()
    conn.commit()
    conn.close()
    pxl.save_as(records=items, dest_file_name=datafile)
    pxl.free_resources()
    return '{:,d} items exported'.format(len(rows))

def getUser():
    if sys.platform == 'win32' or sys.platform == 'cygwin':   # windows
        return os.environ.get("USERNAME")

    elif sys.platform == 'darwin':   # osx64
        return os.environ.get("USERNAME")

    elif sys.platform == 'linux2':   # linux
        return pwd.getpwuid(os.geteuid()).pw_name
    else:
        return os.environ.get("USERNAME")

def getPDFInfo(filename, properties=None, decrypt=None):
    if filename is None:
        return
    if properties is None:
        properties = {}
    columns = ['Category', 'Title', 'Filename' , 'Filesize', 'Location', 'Acquired',
           'Author', 'Date', 'Keyword', 'Publisher', 'Subject', 'URL']
    infos = {}
    infos['/Title'] = columns.index('Title')
    infos['/Author'] = columns.index('Author')
    infos['/ModDate'] = columns.index('Date')
    infos['/Keywords'] = columns.index('Keyword')
    infos['/Publisher'] = columns.index('Publisher')
    infos['/Subject'] = columns.index('Subject')
    try:
        pdf_toread = PdfFileReader(open(filename, 'rb'))
    except:
        return properties
    # from https://stackoverflow.com/questions/26242952/pypdf-2-decrypt-not-working
    if pdf_toread.isEncrypted:
        try:
            pdf_toread.decrypt('')
        except:
            if decrypt is not None and decrypt == 'qpdf':
                dfile = filename
                dfile = dfile.replace(' ', '\\ ')
                tfile = tempfile.gettempdir() + '/temp.pdf'
                tfile2 = tempfile.gettempdir() + '/temp2.pdf'
                command = ('cp ' + dfile + ' ' + tfile + "; qpdf --password=''" + \
                           ' --decrypt ' + tfile + ' ' + tfile2 + \
                           '; rm ' + tfile)
                try:
                    os.system(command)
                    pdf_toread = PdfFileReader(open(tfile2, 'rb'))
                    os.remove(tfile2)
                except:
                    return properties
            else:
                return properties
    try:
        pdf_info = pdf_toread.getDocumentInfo()
    except:
        pdf_info = None
    if pdf_info is not None:
        for key, value in pdf_info.items():
            if key in infos.keys():
                 if key == '/ModDate':
                     try:
                         if value[:2] == 'D:':
                             properties[columns[infos[key]]] = \
                                  value[2:6] + '-' + value[6:8] + '-' + value[8:10]
                     except:
                         pass
                 else:
                     try:
                         properties[columns[infos[key]]] = value.decode()
                     except AttributeError:
                         try:
                             properties[columns[infos[key]]] = value
                         except:
                             pass
                     except UnicodeDecodeError:
                         value_bytes = bytearray(value)
                         for b in range(len(value_bytes) -1, -1, -1):
                             if value_bytes[b] > 127:
                                 if value_bytes[b] == 144: # 0x90
                                     value_bytes[b] = 39 # '
                                 elif value_bytes[b] == 139: # 0x85
                                     value_bytes[b] = 45 # - or _ # '
                                 else:
                                     value_bytes[b] = 46 # .
                             elif value_bytes[b] in [10, 13]:
                                 del value_bytes[b]
                         value = b"".join([value_bytes]).decode()
                         properties[columns[infos[key]]] = value
                     except:
                         pass
    return properties

def getISBNInfo(isbn, db_conn):
    cur = db_conn.cursor()
    cur.execute("select description from fields where typ = 'Settings' and field = 'ISBN Field'")
    try:
        isbn_field = cur.fetchone()[0]
    except:
        isbn_field = 'ISBN'
    cur.execute("select description from fields where typ = 'Settings' and field = 'Dewey Field'")
    try:
        dewey_field = cur.fetchone()[0]
    except:
        dewey_field = 'Dewey Decimal'
    cur.close()
    properties = {isbn_field: isbn}
    conn = http.client.HTTPConnection('openlibrary.org')
    conn.request('GET', '/api/books?bibkeys=ISBN:' + isbn + '&jscmd=data&format=json')
    response = conn.getresponse()
    if response.status == 200 and response.reason == 'OK':
        data_dict = json.loads(response.read())
        for isbn, data in data_dict.items():
            by_statement = ''
            for key, value in data.items():
            # deal with the keys I'm interested in
                if key == 'title':
                    properties['Title'] = value
                elif key == 'authors':
                    authors = ''
                    for who in value:
                       authors += who['name'] + ', '
                    authors = authors[:-2]
                    properties['Author'] = authors
                elif key == 'by_statement':
                    by_statement = value
                elif key == 'classifications':
                    dewey = ''
                    try:
                        for clas in value['dewey_decimal_class']:
                            dewey += clas + ', '
                        dewey = dewey[:-2]
                    except:
                        pass
                    properties[dewey_field] = dewey
                elif key == 'identifiers':
                    try:
                        properties[isbn_field] = value['isbn_13'][0]
                    except:
                        pass
                elif key == 'publishers':
                    publishers = ''
                    for who in value:
                        publishers += who['name'] + ', '
                    publishers = publishers[:-2]
                    properties['Publisher'] = publishers
                elif key == 'publish_date':
                    properties['Date'] = value
                elif key == 'notes':
                    properties['Notes'] = value
            if by_statement != '':
                if 'Author' not in properties.keys():
                    properties['Author'] = by_statement
                else:
                    if 'Notes' in properties.keys():
                        properties['Notes'] = properties['Notes'] + '\n' + by_statement
                    else:
                        properties['Notes'] = by_statement
    else:
        print(str(response.status) + ' ' + response.reason)
    conn.close()
    return properties
