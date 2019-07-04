import os
import logging
import psycopg2
import pymysql
import datetime
from happig_log import Log
import threading
from rw_lock import RWLock


class Database(object):
    def __init__(self, save_path, MYSQL_setting, POSTGRE_setting):
        self.MYSQL_setting = MYSQL_setting
        self.POSTGRE_setting = POSTGRE_setting
        self.log = Log(save_path)

        self.PostGre_connection = self.try_connect_PostGre()
        self.MySQL_connection = self.try_connect_MySQL()
        self.rw_lock = RWLock()

    def try_connect_PostGre(self):
        try:
            connection = psycopg2.connect(
                database=self.POSTGRE_setting["database"], user=self.POSTGRE_setting["user"],
                password=self.POSTGRE_setting["password"], host=self.POSTGRE_setting["host"],
                port=int(self.POSTGRE_setting["port"]), options='-c statement_timeout=5s')
            logging.info(connection)
            self.PostGre_connected = True
        except Exception as e:
            print(e)
            logging.error(e)
            self.PostGre_connected = False
            return None
        return connection

    def try_connect_MySQL(self):
        try:
            connection = pymysql.connect(
                database=self.MYSQL_setting["database"], user=self.MYSQL_setting["user"],
                password=self.MYSQL_setting["password"], host=self.MYSQL_setting["host"],
                port=int(self.MYSQL_setting["port"]))
            logging.info(connection)
        except Exception as e:
            print(e)
            logging.error(e)
            return None
        return connection

    def insert_tag(self, shed_id, sensor_id, tag, read_time):
        sql = "INSERT INTO rfid_records (shed_id, sensor_id, tag_id, read_time) VALUES (%s, %s, %s, %s)"
        val = (shed_id, sensor_id, tag, read_time)

        # 若 PostGre 之前斷線則重連
        prev_connection = self.PostGre_connected
        if not prev_connection:
            self.PostGre_connection = self.try_connect_PostGre()

        # 寫入 PostGre 或 MySQL
        self.PostGre_connected = self.try_do_insert_tag(
            self.PostGre_connection, 'PostGre', sql, val)
        if not self.PostGre_connected:
            self.try_do_insert_tag(self.MySQL_connection, 'MySQL', sql, val)

        # 若重新連上 PostGre，進行 data recovery
        if prev_connection == False and self.PostGre_connected == True:
            try:
                t = threading.Thread(target=self.data_recovery)
                t.start()
            except Exception as e:
                print(e)
                logging.error(e)

        prev_connection = self.PostGre_connected

    def try_do_insert_tag(self, connection, connection_name, sql, val):
        ok = True
        self.rw_lock.writer_acquire()
        try:
            cursor = connection.cursor()
            cursor.execute(sql, val)
            connection.commit()
        except Exception as e:
            print(e)
            logging.error(e)
            ok = False
        self.rw_lock.writer_release()
        if ok:
            msg = '%s#sonsor_id %s : ' % (
                connection_name, val[1])  # val[1] = sonsor_id
            if cursor.rowcount > 0:
                msg += str(cursor.rowcount) + " tag record inserted."
            else:
                msg += 'insert tag failed !'
            print(msg)
            logging.info(msg)
        return ok

    def data_recovery(self):
        print('data recovering...')
        sql = "SELECT COUNT(*) FROM rfid_records"
        try:
            MySQL_cursor = self.MySQL_connection.cursor()
            MySQL_cursor.execute(sql)
            MySQL_count = MySQL_cursor.fetchone()
            if MySQL_count[0] <= 0:
                print("no data for recovering")
                logging.info("no data for recovering")
                return
            sql = "SELECT * FROM rfid_records"
            MySQL_cursor = self.MySQL_connection.cursor()
            MySQL_cursor.execute(sql)
            MySQL_result = MySQL_cursor.fetchall()
            PostGre_cursor = self.PostGre_connection.cursor()

            for x in MySQL_result:
                sql = "INSERT INTO rfid_records (shed_id, sensor_id, tag_id, read_time) VALUES (%s, %s, %s, %s)"
                val = (x[1], x[2], x[3], x[4])
                PostGre_cursor.execute(sql, val)
                self.PostGre_connection.commit()

                if PostGre_cursor.rowcount > 0:
                    sql = "DELETE FROM rfid_records WHERE shed_id = %s AND sensor_id = %s AND tag_id = %s AND read_time = %s"
                    MySQL_cursor = self.MySQL_connection.cursor()
                    MySQL_cursor.execute(sql, val)
                    self.MySQL_connection.commit()
            print("data recovery finished")
            logging.info("data recovery finished")
        except Exception as e:
            print("data recovery failed")
            logging.info("data recovery failed")
            print(e)
            logging.error(e)

    def read_connection(self, connection):
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM rfid_records")
        connection.commit()
        myresult = cursor.fetchall()
        ps = myresult
        if len(ps) > 3:
            ps = ps[-3:]
        for x in ps:
            print(x)

    def readall(self):
        try:
            print('-'*50)
            print('PostGre')
            self.read_connection(self.PostGre_connection)
            print('-'*50)
            print('MySQL')
            self.read_connection(self.MySQL_connection)
            print('-'*50)
        except Exception as e:
            print(e)
            logging.error(e)

    def can_connect(self, host):
        return os.system("ping -W 1 -c 1 " + host) == 0
