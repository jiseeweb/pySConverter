import pyodbc
import os
import shutil
from datetime import datetime

class Converter:


    def __init__(self):
        self.default_save_path = r'C:\Users\{}\Desktop\Convert'.format(os.getlogin())
        self.folder_dest_path = 'C:\\'
        self.dbq_path = ''
        self.autosave = False
        self.is_connected = False
        self.driver_name = 'Microsoft Access Driver (*.mdb, *.accdb)'
        self.conn_str = tuple()
        self.q_str = ''
        self.is_cursor_active = False
        self.connection = None
        self.cursor = None
        self.user_id = 'USERID'
        self.badgenum = 'Badgenumber'
        self.default_table_target = 'CHECKINOUT'


    def save_dest(self, folder_path):
        #Create a folder for the modified databases.
        if os.path.exists(folder_path):
            self.default_save_path = folder_path
            return folder_path
        elif not folder_path:
            return self.default_save_path
        else:
            try:
                os.makedirs(folder_path)
                self.default_save_path = folder_path
                return folder_path
            except PermissionError as err:
                print('Permission Denied: ', err)
                return None


    def set_driver_name(self, driver_name):
        #Set what database driver should we use.
        #Uses Microsoft Access Driver as default.
        if driver_name: self.driver_name = driver_name


    def set_dbq_path(self, dbq_path):
        #Set the database filepath to use when building connection.
        self.dbq_path = dbq_path


    def create_sub_folder(self):
        now = datetime.now()
        folder_name = os.path.join(self.default_save_path, now.strftime('%m-%d-%y %H_%M_%S %p'))
        os.makedirs(folder_name)
        return folder_name


    def build_connection(self):
        #Builds connection string for use of connection by pyodbc.
        try:
            filepath = os.path.join(self.create_sub_folder(), os.path.basename(self.dbq_path))
        except FileNotFoundError as err:
            return 'File or directory not found.\nPlease check your specified paths. again.'
        shutil.copy(self.dbq_path, filepath)
        conn_str = (
            r'Driver={};'
            r'DBQ={}'.format(self.driver_name, filepath)
        )
        self.conn_str = conn_str


    def start_connection(self):
        #Create a connection and cursor object.
        self.is_connected = True
        if self.is_connected:
            connect = pyodbc.connect(self.conn_str, autocommit=self.autosave)
            self.connection = connect
            self.cursor = connect.cursor()
            self.is_cursor_active = True
        else:
            return None


    # def fetch_query(self, q_str):
    #     if isinstance(q_str, str):
    #         try:
    #             connection = start_connection()
    #             self.cursor = connection.cursor()
    #             data = self.cursor.execute(q_str)
    #         except pyodbc.Error as err:
    #             print(err)
    #             return None
    #         else:
    #             return data.fetchall()
    #     else:
    #         print('Query String should be string!')
    #         return None


    def col_name_validator(self, table_name='USERINFO'):
        #Get all the list of columns in the specified table_name
        table = list(self.cursor.tables(table=table_name).fetchone())
        columns = self.cursor.columns(table=table_name)
        return [column.column_name.lower() for column in columns]


    def get_userinfo(self, table_name='USERINFO'):
        # Take the user id and badgenumber in table_name.
        if self.is_cursor_active:
            data = self.col_name_validator()
            if 'userid' and 'badgenumber' in data:
                return self.cursor.execute(r"select {}, {} from {}".format(self.user_id, self.badgenum, table_name)).fetchall()
            else:
                return None
        else:
            return None


    # def select_targets(self, column_name, table_name):
    # Accepts tuple(col_name and table_name)
    def select_targets(self, row_col_target):
        if isinstance(row_col_target, tuple):
            try:
                column_name, table_name = row_col_target
            except IndexError as err:
                print('Column or Table name was not given!')
                return None
            else:
                #Select what column
                if not column_name and table_name:
                    print('You must specify a column and target table!')
                    return None
                else:
                    return column_name, table_name
        else:
            print('Target given was not tuple')
            return None


    def update_database(self, targets):
        if not isinstance(targets, tuple):
            print('Column and table target must not be empty!')
            return None
        else:
            col, table_name = targets
            userinfo = self.get_userinfo(table_name='USERINFO')
            datas = self.cursor.execute(r"select {} from {}".format(col, table_name)).fetchall()

            for users in datas:
                for user_id in userinfo:
                    if user_id[0] == users[0]:
                        self.cursor.execute("update {} set USERID=? where USERID=?".format(table_name,  user_id[1], users[0]))
                    else:
                        continue
            print('Done.')
            print('Saving...')
            if not self.autosave:
                self.connection.commit()

            else:
                print('Finished.')
                print('Conversion completed')


# c = Converter()
# c.set_driver_name('Microsoft Access Driver (*.mdb, *.accdb)')
# c.set_dbq_path('C:\\Users\\Admin3\\Desktop\\Master Folder\\Work\\granding.mdb')
# c.build_connection()
# c.start_connection()
# c.get_userinfo(table_name='USERINFO')
# targets = c.select_targets('USERID', 'CHECKINOUT')
# c.update_database(targets)
