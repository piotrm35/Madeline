"""
/***************************************************************************
  DB_connection_PostgreSQL.py

  PostgreSQL data base connection and manipulation.
  --------------------------------------
  version : 1.2
  Copyright: (C) 2018 by Piotr Michałowski
  Email: piotrm35@hotmail.com
/***************************************************************************
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License version 2 as published
 * by the Free Software Foundation.
 *
 ***************************************************************************/
"""
#------------------------------------------------------------------------------------------------------
# pip install psycopg2

import sys, traceback
import psycopg2

#========================================================================================================


class DB_connection_PostgreSQL:


    def __init__(self, DB_user_str, DB_password_str, DB_host_str, DB_port_str, DB_name_str):
        self.DB_user_str = DB_user_str
        self.DB_password_str = DB_password_str
        self.DB_host_str = DB_host_str
        self.DB_port_str = DB_port_str
        self.DB_name_str = DB_name_str
        self.DB_connection_status = False
        self.my_print("DB_connection_PostgreSQL.__init__ - OK")
        self.start_DB_connection(self.DB_user_str, self.DB_password_str, self.DB_host_str, self.DB_port_str, self.DB_name_str)


    def __del__(self):
        self.Stop_DB_connection()
        self.log_file_close()


    #----------------------------------------------------------------------------------------------------
    # public:


    def Restart_DB_connection(self):
        self.Stop_DB_connection()
        self.start_DB_connection(self.DB_user_str, self.DB_password_str, self.DB_host_str, self.DB_port_str, self.DB_name_str)


    def Stop_DB_connection(self):
        if self.DB_connection_status:
            try:
                self.cursor.close()
                self.cnx.close()
                self.my_print("Stop_DB_connection - OK")
            except:
                self.my_print("Stop_DB_connection - EXCEPTION: " + self.get_current_system_EXCEPTION_info())
            finally:
                self.DB_connection_status = False
        self.log_file_close()


    def Send_SQL_to_DB(self, sql):
        if self.DB_connection_status:
            try:
                self.cursor.execute(sql)
                self.cnx.commit()
                if sql.upper().startswith('SELECT'):
                    return self.cursor.fetchall()
                return 'OK'
            except:
                self.DB_connection_failure_counter += 1
                self.my_print("Send_SQL_to_DB - EXCEPTION(" + str(self.DB_connection_failure_counter) + "): ")
                self.my_print('sql = ' + str(sql))
                self.my_print(self.get_current_system_EXCEPTION_info())
        else:
            self.my_print("Send_SQL_to_DB - ERROR: self.DB_connection_status = False")
        return 'ERROR'


    #----------------------------------------------------------------------------------------------------
    # private:


    def start_DB_connection(self, DB_user_str, DB_password_str, DB_host_str, DB_port_str, DB_name_str):
        try:
            self.cnx = psycopg2.connect('user=' + DB_user_str + ' password=' + DB_password_str + ' host=' + DB_host_str + ' port=' + DB_port_str + ' dbname=' + DB_name_str)
            self.cursor = self.cnx.cursor()
            self.DB_connection_failure_counter = 0
            self.DB_connection_status = True
            self.my_print("Start_DB_connection - OK")
        except Exception as e:
            self.DB_connection_status = False
            raise Exception("Start_DB_connection - EXCEPTION: " + self.get_current_system_EXCEPTION_info())


    log_file = None
    

    def my_print(self, tx):
        print(tx)
##        if not self.log_file:
##            self.log_file = open('DB_connection_PostgreSQL_log_file.txt', 'w')
##        self.log_file.write(tx + '\n')


    def log_file_close(self):
        if self.log_file:
            self.log_file.close()
            print('DB_connection_PostgreSQL.log_file_close - OK')


    def get_current_system_EXCEPTION_info(self):
        ex_type, ex_value, ex_traceback = sys.exc_info()
        trace_back = traceback.extract_tb(ex_traceback)
        stack_trace = list()
        for trace in trace_back:
            stack_trace.append("File : %s , Line : %d, Func.Name : %s, Message : %s" % (trace[0], trace[1], trace[2], trace[3]))
        ex_info = ''
        ex_info += "\nException type : %s " % ex_type.__name__ + '\n'
        ex_info += "Exception message : %s" %ex_value + '\n'
        ex_info += "Stack trace : %s" %stack_trace
        return ex_info


#========================================================================================================
