"""
/***************************************************************************
  Madeline_x.x.py

  Generator dokumentów oparty na szablonach i konfiguratorze.
  --------------------------------------
  Copyright: (C) 2022 by Piotr Michałowski
  Email: piotrm35@hotmail.com
/***************************************************************************
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License version 2 as published
 * by the Free Software Foundation.
 *
 ***************************************************************************/
"""
SCRIPT_TITLE = 'Madeline'
SCRIPT_VERSION = '1.2'
GENERAL_INFO = u"""
author: Piotr Michałowski, Olsztyn, woj. W-M, Poland
piotrm35@hotmail.com
work begin: 03.10.2022
"""

import os, sys, datetime
import codecs
from PyQt5 import QtCore, QtWidgets, uic
import pyautogui
from docx import Document                   # sudo pip3 install python-docx
from lib.My_QLabel import My_QLabel
from lib.DB_connection_PostgreSQL import DB_connection_PostgreSQL
from lib.Text_filter import Text_filter

# setup begin --------------------------------------
DATABASE_USER = 'user_1'
#---------------------------------------------------
DOC_DATABASE_HOST = '127.0.0.1'
DOC_DATABASE_PORT = '5432'
DOC_DATABASE_NAME = 'documents'
#---------------------------------------------------
GGN_INFO_DATABASE_HOST = '127.0.0.1'
GGN_INFO_DATABASE_PORT = '5432'
GGN_INFO_DATABASE_NAME = 'ewid'
# setup end ----------------------------------------


#====================================================================================================================


class Madeline(QtWidgets.QMainWindow):


    def __init__(self):
        super(Madeline, self).__init__()
        self.DECYZJE_FOLDER_PATH = '.\\Decyzje'
        self.base_path = os.sep.join(os.path.realpath(__file__).split(os.sep)[0:-1])
        print("__init__, self.base_path = " + str(self.base_path))
        uic.loadUi(os.path.join(self.base_path, 'ui', 'Madeline.ui'), self)
        self.setWindowTitle(SCRIPT_TITLE + ' v. ' + SCRIPT_VERSION)
        self.Nr_dokumentu_lineEdit.textChanged.connect(self.Nr_dokumentu_lineEdit_textChanged)
        self.Skladajacy_textEdit.textChanged.connect(self.Skladajacy_textEdit_textChanged)
        self.Zapisz_skladajacy_pushButton.clicked.connect(self.Zapisz_skladajacy_pushButton_clicked)
        self.Dzialajacy_checkBox.stateChanged.connect(self.Dzialajacy_checkBox_stateChanged)
        self.Inwestor_textEdit.textChanged.connect(self.Inwestor_textEdit_textChanged)
        self.Zapisz_inwestor_pushButton.clicked.connect(self.Zapisz_inwestor_pushButton_clicked)
        self.Numery_dzialek_textEdit.textChanged.connect(self.Numery_dzialek_textEdit_textChanged)
        self.Analizuj_pushButton.clicked.connect(self.Analizuj_pushButton_clicked)
        self.Przedmiot_textEdit.textChanged.connect(self.Przedmiot_textEdit_textChanged)
        self.Zapisz_przedmiot_pushButton.clicked.connect(self.Zapisz_przedmiot_pushButton_clicked)
        self.Csv_file_pushButton.clicked.connect(self.Csv_file_pushButton_clicked)
        self.Automat_pushButton.clicked.connect(self.Automat_pushButton_clicked)
        self.Generate_pushButton.clicked.connect(self.Generate_pushButton_clicked)
        self.Clear_pushButton.clicked.connect(self.Clear_pushButton_clicked)
        self.Prompt_pushButton.clicked.connect(self.Prompt_pushButton_clicked)
        self.Info_pushButton.clicked.connect(self.Info_pushButton_clicked)
        password = pyautogui.password(text='for ligin ' + DATABASE_USER + ':', title='DB password', default='', mask='*')
        self.DOC_dB_connection_PostgreSQL = DB_connection_PostgreSQL(DATABASE_USER, password, DOC_DATABASE_HOST, DOC_DATABASE_PORT, DOC_DATABASE_NAME)
        self.GGN_INFO_dB_connection_PostgreSQL = DB_connection_PostgreSQL(DATABASE_USER, password, GGN_INFO_DATABASE_HOST, GGN_INFO_DATABASE_PORT, GGN_INFO_DATABASE_NAME)
        self.text_filter = Text_filter()
        self.DOC_TEMPLATES_MAP = {
                'Decyzja - TEMPLATE.docx': self.Decyzja_radioButton,
                'Zgoda na przebudowę - TEMPLATE.docx': self.Zgoda_na_przebudowe_radioButton,
                'Zezwolenie - TEMPLATE.docx': self.Zezwolenie_dr_wewn_radioButton,
                'Uzgodnienie - TEMPLATE.docx': self.Uzgodnienie_radioButton,
                'Uzgodnienie + SŁUŻEBNOŚĆ - TEMPLATE.docx': self.Uzgodnienie_i_sluz_przes_radioButton,
                'Uzgodnienie + PRZEKAZANIE - TEMPLATE.docx': self.Uzgodnienie_i_przekazanie_radioButton,
                'Uzgodnienie i opinia - TEMPLATE.docx': self.Uzgodnienie_i_opinia_radioButton,
                'Uzgodnienie i opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx': self.Uzgodnienie_i_opinia_i_sluz_przes_radioButton,
                'Uzgodnienie i opinia  + PRZEKAZANIE - TEMPLATE.docx': self.Uzgodnienie_i_opinia_i_przekazanie_radioButton,
                'Opinia - TEMPLATE.docx': self.Opinia_radioButton,
                'Opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx': self.Opinia_i_sluz_przes_radioButton,
                'Opinia + PRZEKAZANIE - TEMPLATE.docx': self.Opinia_i_przekazanie_radioButton,
                'Pismo puste - TEMPLATE.docx': self.Pismo_puste_radioButton
            }
        self.DOC_TEMPLATES_NO_SENDER_LIST = [
                'Uzgodnienie - TEMPLATE.docx',
                'Uzgodnienie + SŁUŻEBNOŚĆ - TEMPLATE.docx',
                'Uzgodnienie i opinia - TEMPLATE.docx',
                'Uzgodnienie i opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx',
                'Opinia - TEMPLATE.docx',
                'Opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx'
            ]
        self.RODZAJ_MAP = {
                'lokalizację': self.Lokalizacja_radioButton,
                'budowę': self.Budowa_radioButton,
                'przebudowę': self.Przebudowa_radioButton,
                'remont': self.Remont_radioButton
            }
        self.parcel_select_sql_result = None
        self.automat_map = None
        self.numer_dokumentu_w_sprawie = 0
        
        
    def closeEvent(self, event):        # overriding the method
        self.Nr_dokumentu_lineEdit.textChanged.disconnect(self.Nr_dokumentu_lineEdit_textChanged)
        self.Skladajacy_textEdit.textChanged.disconnect(self.Skladajacy_textEdit_textChanged)
        self.Zapisz_skladajacy_pushButton.clicked.disconnect(self.Zapisz_skladajacy_pushButton_clicked)
        self.Dzialajacy_checkBox.stateChanged.disconnect(self.Dzialajacy_checkBox_stateChanged)
        self.Inwestor_textEdit.textChanged.disconnect(self.Inwestor_textEdit_textChanged)
        self.Zapisz_inwestor_pushButton.clicked.disconnect(self.Zapisz_inwestor_pushButton_clicked)
        self.Numery_dzialek_textEdit.textChanged.disconnect(self.Numery_dzialek_textEdit_textChanged)
        self.Analizuj_pushButton.clicked.disconnect(self.Analizuj_pushButton_clicked)
        self.Przedmiot_textEdit.textChanged.disconnect(self.Przedmiot_textEdit_textChanged)
        self.Zapisz_przedmiot_pushButton.clicked.disconnect(self.Zapisz_przedmiot_pushButton_clicked)
        self.Csv_file_pushButton.clicked.disconnect(self.Csv_file_pushButton_clicked)
        self.Automat_pushButton.clicked.disconnect(self.Automat_pushButton_clicked)
        self.Generate_pushButton.clicked.disconnect(self.Generate_pushButton_clicked)
        self.Clear_pushButton.clicked.disconnect(self.Clear_pushButton_clicked)
        self.Prompt_pushButton.clicked.disconnect(self.Prompt_pushButton_clicked)
        self.Info_pushButton.clicked.disconnect(self.Info_pushButton_clicked)
        self.DOC_dB_connection_PostgreSQL.Stop_DB_connection()
        self.GGN_INFO_dB_connection_PostgreSQL.Stop_DB_connection()
        event.accept()


    #----------------------------------------------------------------------------------------------------------------
    # input widget methods:


    def Nr_dokumentu_lineEdit_textChanged(self):
        self.numer_dokumentu_w_sprawie = 0


    currently_edited_widget = None


    def Skladajacy_textEdit_textChanged(self):
        if not self.prompt_used_flag:
            self.currently_edited_widget = self.Skladajacy_textEdit
            tx = self.Skladajacy_textEdit.toPlainText().strip()
            if len(tx) > 0:
                self.Zapisz_skladajacy_pushButton.setEnabled(True)
                self.Prompt_pushButton.setEnabled(True)
            else:
                self.Zapisz_skladajacy_pushButton.setEnabled(False)
                self.Prompt_pushButton.setEnabled(False)
        else:
            self.Zapisz_skladajacy_pushButton.setEnabled(False)
            self.prompt_used_flag = False


    def Zapisz_skladajacy_pushButton_clicked(self):
        tx = self.Skladajacy_textEdit.toPlainText().strip()
        tx = self.clear_text(tx)
        if len(tx) > 0:
            if not self.is_present_in_adresaci_TAB(tx):
                INSERT_SQL = "INSERT INTO adresaci (adres) VALUES('" + tx + "');"
                print("Zapisz_skladajacy_pushButton_clicked, INSERT_SQL = " + str(INSERT_SQL))
                result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(INSERT_SQL)
                if result == 'ERROR':
                    self.DOC_dB_connection_PostgreSQL.Restart_DB_connection()
                    result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(INSERT_SQL)
                print("Zapisz_skladajacy_pushButton_clicked, result = " + str(result))
            else:
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'TEKST JEST JUŻ W BAZIE DANYCH')
            self.Zapisz_skladajacy_pushButton.setEnabled(False)
        

    def Dzialajacy_checkBox_stateChanged(self):
        self.Inwestor_textEdit.setEnabled(self.Dzialajacy_checkBox.isChecked())
        self.Zapisz_inwestor_pushButton.setEnabled(self.Dzialajacy_checkBox.isChecked())


    def Inwestor_textEdit_textChanged(self):
        if not self.prompt_used_flag:
            self.currently_edited_widget = self.Inwestor_textEdit
            tx = self.Inwestor_textEdit.toPlainText().strip()
            if len(tx) > 0:
                self.Zapisz_inwestor_pushButton.setEnabled(True)
                self.Prompt_pushButton.setEnabled(True)
            else:
                self.Zapisz_inwestor_pushButton.setEnabled(False)
                self.Prompt_pushButton.setEnabled(False)
        else:
            self.Zapisz_inwestor_pushButton.setEnabled(False)
            self.prompt_used_flag = False


    def Zapisz_inwestor_pushButton_clicked(self):
        tx = self.Inwestor_textEdit.toPlainText().strip()
        tx = self.clear_text(tx)
        if len(tx) > 0:
            if not self.is_present_in_adresaci_TAB(tx):
                INSERT_SQL = "INSERT INTO adresaci (adres) VALUES('" + tx + "');"
                print("Zapisz_inwestor_pushButton_clicked, INSERT_SQL = " + str(INSERT_SQL))
                result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(INSERT_SQL)
                if result == 'ERROR':
                    self.DOC_dB_connection_PostgreSQL.Restart_DB_connection()
                    result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(INSERT_SQL)
                print("Zapisz_inwestor_pushButton_clicked, result = " + str(result))
            else:
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'TEKST JEST JUŻ W BAZIE DANYCH')
            self.Zapisz_inwestor_pushButton.setEnabled(False)


    def Numery_dzialek_textEdit_textChanged(self):
        tx = self.Numery_dzialek_textEdit.toPlainText().strip()
        if len(tx) > 0:
            self.Analizuj_pushButton.setEnabled(True)
        else:
            self.Analizuj_pushButton.setEnabled(False)


    def Analizuj_pushButton_clicked(self):
        print('\nAnalizuj_pushButton_clicked START ##################################################')
        self.text_filter.set_text(self.Numery_dzialek_textEdit.toPlainText().strip())
        parcel_list = self.text_filter.get_parcel_list()
        print('parcel_list = ' + str(parcel_list))
        if parcel_list:
            PARCEL_SELECT_SQL = "SELECT id_umk, jednostka_nzw, rodzaj_wld_opis, opis_gr, uzytek_gr, numer_drogi, nazwa_ulicy FROM ggn_info WHERE id_umk = '"
            sql = PARCEL_SELECT_SQL + "' OR id_umk = '".join(parcel_list) + "';"
            self.parcel_select_sql_result = self.GGN_INFO_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
            if self.parcel_select_sql_result == 'ERROR':
                self.GGN_INFO_dB_connection_PostgreSQL.Restart_DB_connection()
                self.parcel_select_sql_result = self.GGN_INFO_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
                if self.parcel_select_sql_result == 'ERROR':
                    print("Analizuj_pushButton_clicked PROBLEM: self.parcel_select_sql_result == 'ERROR'")
                    QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, "self.parcel_select_sql_result == 'ERROR'")
                    print('Analizuj_pushButton_clicked STOP ###################################################\n')
                    return
            if len(parcel_list) != len(self.parcel_select_sql_result):
                tmp_działki_list = parcel_list.copy()
                for res in self.parcel_select_sql_result:
                    del(tmp_działki_list[tmp_działki_list.index(res[0])])
                print('Analizuj_pushButton_clicked PROBLEM: len(parcel_list) != len(self.parcel_select_sql_result), BRAK DANYCH DLA: ' + str(tmp_działki_list))
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'len(parcel_list) != len(self.parcel_select_sql_result), BRAK DANYCH DLA: ' + str(tmp_działki_list))
            print('\n-----------------------------------------------------------------------------------')
            print('self.parcel_select_sql_result = ' + str(self.parcel_select_sql_result))
            print('\n')
            if not self.parcel_select_sql_result:
                print('Analizuj_pushButton_clicked PROBLEM: BRAK DANYCH Z BAZY DANYCH DZIAŁEK EWIDENCYJNYCH')
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BRAK DANYCH Z BAZY DANYCH DZIAŁEK EWIDENCYJNYCH')
                print('Analizuj_pushButton_clicked STOP ###################################################\n')
                return
            conv = lambda i : i or 'None'
            self.automat_map = self.get_initial_automat_map()
            skladajacy = self.Skladajacy_textEdit.toPlainText().strip()
            for res in self.parcel_select_sql_result:
                if res[2] == 'trwały zarząd':        # rodzaj_wld_opis
                    if res[5] is not None:          # numer_drogi
                        rodzaj = None
                        for key, value in self.RODZAJ_MAP.items():
                            if value.isChecked():
                                rodzaj = key
                                break
                        if not rodzaj or rodzaj == 'lokalizację':
                            print('Analizuj_pushButton_clicked PROBLEM: WYBIERZ RODZAJ (budowa, przebudowa lub remont)')
                            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'WYBIERZ RODZAJ (budowa, przebudowa lub remont)')
                            print('Analizuj_pushButton_clicked STOP ###################################################\n')
                            return
                        if rodzaj == 'budowę':
                            self.automat_map['Decyzja - TEMPLATE.docx'].append(res[0])
                        elif rodzaj == 'przebudowę' or rodzaj == 'remont':
                            self.automat_map['Zgoda na przebudowę - TEMPLATE.docx'].append(res[0])
                        else:
                            print('Analizuj_pushButton_clicked PROBLEM: rodzaj(' + res[0] + ') = ' + str(rodzaj) + ' -> POMINIĘTO w automat_map')
                            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'rodzaj(' + res[0] + ') = ' + str(rodzaj) + ' -> POMINIĘTO w automat_map')
                    else:
                        self.automat_map['Zezwolenie - TEMPLATE.docx'].append(res[0])
                elif res[2] == 'administracja':      # rodzaj_wld_opis
                    if res[4] == 'dr':              # uzytek_gr
                        if not skladajacy or skladajacy.upper() == 'GGN':
                            self.automat_map['Uzgodnienie + SŁUŻEBNOŚĆ - TEMPLATE.docx'].append(res[0])
                        else:
                            self.automat_map['Uzgodnienie + PRZEKAZANIE - TEMPLATE.docx'].append(res[0])
                    else:
                        if not skladajacy or skladajacy.upper() == 'GGN':
                            self.automat_map['Opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx'].append(res[0])
                        else:
                            self.automat_map['Opinia + PRZEKAZANIE - TEMPLATE.docx'].append(res[0])
                else:
                    print('Analizuj_pushButton_clicked, rodzaj_wld_opis(' + res[0] + ') = ' + str(res[2]) + ' -> POMINIĘTO w automat_map')
                    QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'rodzaj_wld_opis(' + res[0] + ') = ' + str(res[2]) + ' -> POMINIĘTO w automat_map')
            if self.automat_map['Uzgodnienie + SŁUŻEBNOŚĆ - TEMPLATE.docx'] and self.automat_map['Opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx']:
                self.automat_map['Uzgodnienie i opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx'][0] += self.automat_map['Uzgodnienie + SŁUŻEBNOŚĆ - TEMPLATE.docx']
                self.automat_map['Uzgodnienie i opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx'][1] += self.automat_map['Opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx']
                self.automat_map['Uzgodnienie + SŁUŻEBNOŚĆ - TEMPLATE.docx'] = []
                self.automat_map['Opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx'] = []
            if self.automat_map['Uzgodnienie + PRZEKAZANIE - TEMPLATE.docx'] and self.automat_map['Opinia + PRZEKAZANIE - TEMPLATE.docx']:
                self.automat_map['Uzgodnienie i opinia  + PRZEKAZANIE - TEMPLATE.docx'][0] += self.automat_map['Uzgodnienie + PRZEKAZANIE - TEMPLATE.docx']
                self.automat_map['Uzgodnienie i opinia  + PRZEKAZANIE - TEMPLATE.docx'][1] += self.automat_map['Opinia + PRZEKAZANIE - TEMPLATE.docx']
                self.automat_map['Uzgodnienie + PRZEKAZANIE - TEMPLATE.docx'] = []
                self.automat_map['Opinia + PRZEKAZANIE - TEMPLATE.docx'] = []
            self.automat_map = self.remove_empty_elements_from_automat_map(self.automat_map)
            print('self.automat_map = ' + str(self.automat_map))
            if len(self.automat_map) == 0:
                self.Automat_pushButton.setEnabled(False)
                print('Analizuj_pushButton_clicked -> brak danych po analizie')
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'brak danych po analizie')
                print('Analizuj_pushButton_clicked STOP ###################################################\n')
                return
            elif len(self.automat_map) == 1:
                self.Automat_pushButton.setEnabled(False)
            else:
                self.Automat_pushButton.setEnabled(True)
            print('\n===================================================================================\n')
            for key, value in self.automat_map.items():
                print(key)
                if key.startswith('Uzgodnienie i opinia'):
                    print('UZGODNIENIE:')
                    for p in value[0]:
                        for res in self.parcel_select_sql_result:
                            if res[0] == p:
                                print(', '.join([conv(i) for i in res]))
                                break
                    print('OPINIA:')
                    for p in value[1]:
                        for res in self.parcel_select_sql_result:
                            if res[0] == p:
                                print(', '.join([conv(i) for i in res]))
                                break
                else:
                    for p in value:
                        for res in self.parcel_select_sql_result:
                            if res[0] == p:
                                print(', '.join([conv(i) for i in res]))
                                break
            print('\n===================================================================================\n')
            self.numer_dokumentu_w_sprawie = 0
            self.ustaw_dzialki_ulice_i_szablon(self.parcel_select_sql_result, self.automat_map, 0)
        else:
            print('Analizuj_pushButton_clicked PROBLEM: parcel_list = ' + str(parcel_list))
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BRAK DANYCH DZIAŁEK EWIDENCYJNYCH PO FILTROWANIU Z TEKSTU')
        print('Analizuj_pushButton_clicked STOP ###################################################\n')


    def ustaw_dzialki_ulice_i_szablon(self, parcel_select_sql_result, automat_map, idx):
        if not parcel_select_sql_result:
            print('ustaw_dzialki_ulice_i_szablon PROBLEM: parcel_select_sql_result = ' + str(parcel_select_sql_result))
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'PROBLEM: parcel_select_sql_result = ' + str(parcel_select_sql_result))
            return False
        if not automat_map:
            print('ustaw_dzialki_ulice_i_szablon PROBLEM: automat_map = ' + str(automat_map))
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'PROBLEM: automat_map = ' + str(automat_map))
            return False
        if idx < 0 or len(automat_map) <= idx:
            print('automat_map = ' + str(automat_map))
            print('ustaw_dzialki_ulice_i_szablon PROBLEM: idx = ' + str(idx) + ' -> poza zakresem')
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'PROBLEM: idx = ' + str(idx) + ' -> poza zakresem')
            return False
        automat_map_keys = list(automat_map.keys())
        print('automat_map_keys[' + str(idx) + '] = ' + str(automat_map_keys[idx]))
        self.DOC_TEMPLATES_MAP[automat_map_keys[idx]].setChecked(True)
        nazwa_ulicy_list = []
        tmp_działki_list = []
        if automat_map_keys[idx].startswith('Uzgodnienie i opinia'):
            tx = 'UZGODNIENIE:\n'
            tx += ', '.join(automat_map[automat_map_keys[idx]][0])
            tx += '\nOPINIA:\n'
            tx += ', '.join(automat_map[automat_map_keys[idx]][1])
            self.Numery_dzialek_textEdit.setText(tx)
            tmp_działki_list += automat_map[automat_map_keys[idx]][0]
            tmp_działki_list += automat_map[automat_map_keys[idx]][1]
        else:
            self.Numery_dzialek_textEdit.setText(', '.join(automat_map[automat_map_keys[idx]]))
            tmp_działki_list += automat_map[automat_map_keys[idx]]
        print('tmp_działki_list = ' + str(tmp_działki_list))
        for p in tmp_działki_list:
            for res in parcel_select_sql_result:
                if res[0] == p:
                    nazwa_ulicy = res[-1]
                    if nazwa_ulicy:
                        nazwa_ulicy_list.append(nazwa_ulicy.strip())
                    break
        nazwa_ulicy_list = list(set(nazwa_ulicy_list))
        print('nazwa_ulicy_list = ' + str(nazwa_ulicy_list))
        if nazwa_ulicy_list:
            self.Ulice_textEdit.setText(', '.join(nazwa_ulicy_list))
        else:
            self.Ulice_textEdit.setText('.......................')
        return True
                

    def Przedmiot_textEdit_textChanged(self):
        if not self.prompt_used_flag:
            self.currently_edited_widget = self.Przedmiot_textEdit
            tx = self.Przedmiot_textEdit.toPlainText().strip()
            if len(tx) > 0:
                self.Zapisz_przedmiot_pushButton.setEnabled(True)
                self.Prompt_pushButton.setEnabled(True)
            else:
                self.Zapisz_przedmiot_pushButton.setEnabled(False)
                self.Prompt_pushButton.setEnabled(False)
        else:
            self.Zapisz_przedmiot_pushButton.setEnabled(False)
            self.prompt_used_flag = False


    def Zapisz_przedmiot_pushButton_clicked(self):
        tx = self.Przedmiot_textEdit.toPlainText().strip()
        tx = tx.replace('\n', ' ')
        tx = tx.replace(' ,', ', ')
        tx = self.replace_loop(tx, '  ', ' ')
        tx = tx.strip()
        if len(tx) > 0:
            if not self.is_present_in_przedmioty_TAB(tx):
                INSERT_SQL = "INSERT INTO przedmioty (przedmiot) VALUES('" + tx + "');"
                print("Zapisz_przedmiot_pushButton_clicked, INSERT_SQL = " + str(INSERT_SQL))
                result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(INSERT_SQL)
                if result == 'ERROR':
                    self.DOC_dB_connection_PostgreSQL.Restart_DB_connection()
                    result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(INSERT_SQL)
                print("Zapisz_przedmiot_pushButton_clicked, result = " + str(result))
            else:
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'TEKST JEST JUŻ W BAZIE DANYCH')
            self.Zapisz_przedmiot_pushButton.setEnabled(False)


    def Csv_file_pushButton_clicked(self):
        if self.parcel_select_sql_result:
            nr_sprawy = self.Nr_dokumentu_lineEdit.text().strip()
            nr_sprawy_list = nr_sprawy.split('.')
            if len(nr_sprawy_list) != 4:
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNY NUMER SPRAWY')
                return
            file_path = os.path.join(self.DECYZJE_FOLDER_PATH, nr_sprawy_list[2] + ' - działki tab.csv')
            f = codecs.open(file_path, 'w', 'utf-8')
            for res in self.parcel_select_sql_result:
                conv = lambda i : i or 'None'
                f.write(';'.join([conv(i) for i in res]) + '\n')
            f.close()
            print('Csv_file_pushButton_clicked, dane zapisano do pliku: ' + file_path)
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'dane zapisano do pliku: ' + file_path)
        else:
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BRAK DANYCH')


    def Automat_pushButton_clicked(self):
        print('Automat_pushButton_clicked')
        self.numer_dokumentu_w_sprawie = 0
        for idx in range(len(self.automat_map)):
            if not self.ustaw_dzialki_ulice_i_szablon(self.parcel_select_sql_result, self.automat_map, idx):
                break
            if not self.Generate_pushButton_clicked():
                break


    def Generate_pushButton_clicked(self):
        template_file_name = None
        for key, value in self.DOC_TEMPLATES_MAP.items():
            if value.isChecked():
                template_file_name = key
                break
        print("Generate_pushButton_clicked, template_file_name = " + str(template_file_name))
        if not template_file_name:
            print("Generate_pushButton_clicked, PROBLEM: WYBIERZ SZABLON DOKUMENTU")
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'WYBIERZ SZABLON DOKUMENTU')
            return False
        doc = Document(os.path.join('templates', template_file_name))
        rodzaj = None
        for key, value in self.RODZAJ_MAP.items():
            if value.isChecked():
                rodzaj = key
                break
        if not rodzaj:
            print("Generate_pushButton_clicked, PROBLEM: WYBIERZ RODZAJ")
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'WYBIERZ RODZAJ')
            return False
        self.replace_text_DOCX(doc, '[RODZAJ]', rodzaj)
        self.replace_text_DOCX(doc, '+[DATA_DZISIEJSZA]', self.get_time_str())
        if template_file_name not in self.DOC_TEMPLATES_NO_SENDER_LIST:
            adres_skladajacego = self.Skladajacy_textEdit.toPlainText().strip()
            if not adres_skladajacego:
                print("Generate_pushButton_clicked, PROBLEM: BŁĘDNY ADRES SKŁADAJĄCEGO")
                QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNY ADRES SKŁADAJĄCEGO')
                return False
            if self.Dzialajacy_checkBox.isChecked():
                adres_inwestora = self.Inwestor_textEdit.toPlainText().strip()
                if not adres_inwestora:
                    print("Generate_pushButton_clicked, PROBLEM: BŁĘDNY ADRES INWESTORA")
                    QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNY ADRES INWESTORA')
                    return False
                adres_strony = adres_skladajacego + '\ndziałający w imieniu i na rzecz\n' + adres_inwestora
            else:
                adres_strony = adres_skladajacego
            self.replace_text_DOCX(doc, '[ADRES_STRONY]', adres_strony)
            self.replace_text_DOCX(doc, '[ADRES_STRONY_W_LINII]', self.clear_text(adres_skladajacego))
        numery_dzialek = self.Numery_dzialek_textEdit.toPlainText().strip()
        if not numery_dzialek:
            print("Generate_pushButton_clicked, PROBLEM: BŁĘDNY NUMER DZIAŁKI")
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNY NUMER DZIAŁKI')
            return False
        if template_file_name.startswith('Uzgodnienie i opinia'):
            if numery_dzialek.startswith('UZGODNIENIE:') and 'OPINIA:' in numery_dzialek:
                numery_dzialek = numery_dzialek.replace('UZGODNIENIE:', '')
                numery_dzialek_list = numery_dzialek.split('OPINIA:')
                if len(numery_dzialek_list) == 2:
                    numery_dzialek = self.clear_text(numery_dzialek_list[0])
                    numery_dzialek_2 = self.clear_text(numery_dzialek_list[1])
                    self.replace_text_DOCX(doc, '[NUMERY_DZIAŁEK_2]', numery_dzialek_2)
        self.replace_text_DOCX(doc, '[NUMERY_DZIAŁEK]', numery_dzialek)
        nazwa_ulicy = self.Ulice_textEdit.toPlainText().strip()
        if not nazwa_ulicy:
            print("Generate_pushButton_clicked, PROBLEM: BŁĘDNY NAZWA ULICY")
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNA NAZWA ULICY')
            return False
        self.replace_text_DOCX(doc, '[NAZWA_ULICY]', nazwa_ulicy)
        przedmiot = self.Przedmiot_textEdit.toPlainText().strip()
        if not przedmiot:
            print("Generate_pushButton_clicked, PROBLEM: BŁĘDNY NAZWA PRZEDMIOTU")
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNA NAZWA PRZEDMIOTU')
            return False
        self.replace_text_DOCX(doc, '[PRZEDMIOT]', przedmiot)	
        nr_sprawy = self.Nr_dokumentu_lineEdit.text().strip()
        nr_sprawy_list = nr_sprawy.split('.')
        if len(nr_sprawy_list) != 4:
            print("Generate_pushButton_clicked, PROBLEM: BŁĘDNY NUMER SPRAWY")
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNY NUMER SPRAWY')
            return False
        if self.numer_dokumentu_w_sprawie > 0:
            nr_sprawy = nr_sprawy_list[0] + '.' + nr_sprawy_list[1] + '.' + nr_sprawy_list[2] + '-' + str(self.numer_dokumentu_w_sprawie) + '.' + nr_sprawy_list[3]
        self.replace_text_DOCX(doc, '[NUMER_SPRAWY]', nr_sprawy)
        self.numer_dokumentu_w_sprawie += 1
        self.save_document(nr_sprawy, nazwa_ulicy, self.DECYZJE_FOLDER_PATH, template_file_name, doc)
        return True
    

    def Clear_pushButton_clicked(self):
        self.Nr_dokumentu_lineEdit.setText("TE.4061.")
        self.Skladajacy_textEdit.setText("")
        self.Inwestor_textEdit.setText("")
        self.Numery_dzialek_textEdit.setText("")
        self.Ulice_textEdit.setText("")
        self.Przedmiot_textEdit.setText("")
        self.Inwestor_textEdit.setEnabled(False)
        self.Zapisz_skladajacy_pushButton.setEnabled(False)
        self.Zapisz_inwestor_pushButton.setEnabled(False)
        self.Analizuj_pushButton.setEnabled(False)
        self.Zapisz_przedmiot_pushButton.setEnabled(False)
        self.Prompt_pushButton.setEnabled(False)
        self.Automat_pushButton.setEnabled(False)
        self.Dzialajacy_checkBox.setChecked(False)
        self.Lokalizacja_radioButton.setChecked(True)
        self.remove_widgets_from_gridLayout()
        self.prompt_used_flag = False
        self.parcel_select_sql_result = None
        self.automat_map = None
        self.numer_dokumentu_w_sprawie = 0
    

    def Prompt_pushButton_clicked(self):
        if self.currently_edited_widget:
            self.remove_widgets_from_gridLayout()
            if self.currently_edited_widget == self.Przedmiot_textEdit:
                SQL = "SELECT przedmiot FROM przedmioty WHERE przedmiot ~* '" + self.currently_edited_widget.toPlainText().strip() + "';"
            else:       # self.Skladajacy_textEdit lub self.Inwestor_textEdit
                SQL = "SELECT adres FROM adresaci WHERE adres ~* '" + self.currently_edited_widget.toPlainText().strip() + "';"
            self.fill_gridLayout(SQL)
        else:
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BRAK DANYCH')


    def Info_pushButton_clicked(self):
        QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, SCRIPT_TITLE + ' v. ' + SCRIPT_VERSION + '\n\n' + GENERAL_INFO)


    #----------------------------------------------------------------------------------------------------------------
    # gridLayout methods:


    My_QLabel_WIDTH = 200
    My_QLabel_HEIGHT = 50


    def fill_gridLayout(self, sql):
        result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
        if result == 'ERROR':
            self.DOC_dB_connection_PostgreSQL.Restart_DB_connection()
            result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
        print("fill_gridLayout, result = " + str(result))
        row = 0
        for record in result:
            self.gridLayout.addWidget(self.get_QLabel(self.restore_text(str(record[0]))), row, 0, QtCore.Qt.AlignTop)
            self.gridLayout.setRowStretch(row, 1)
            row += 1
            if row >= 10:
                break
        self.gridLayout.addWidget(QtWidgets.QLabel(), row, 0, QtCore.Qt.AlignTop)   # dopełnienie od dołu, żeby "przepchnąć" etykiety do góry
        self.gridLayout.setRowStretch(row, 5)


    prompt_used_flag = False
    

    def set_selected_text_QLabel(self, text):
        if self.currently_edited_widget:
            self.prompt_used_flag = True
            self.currently_edited_widget.setText(text)
        else:
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BRAK DANYCH')


    def get_QLabel(self, text):
        _tmp_label = My_QLabel(self, text)
        _tmp_label.setMinimumSize(self.My_QLabel_WIDTH, self.My_QLabel_HEIGHT)
        _tmp_label.setSizePolicy(QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Expanding))
        _tmp_label.setStyleSheet("border: 1px solid black;")
        _tmp_label.setText(text)
        return _tmp_label


    def remove_widgets_from_gridLayout(self):
        for i in reversed(range(self.gridLayout.count())):
            _widget_to_remove = self.gridLayout.itemAt(i).widget()
            self.gridLayout.removeWidget(_widget_to_remove)
            _widget_to_remove.setParent(None)
            del _widget_to_remove


    #----------------------------------------------------------------------------------------------------------------
    # auxiliary methods:
    

    def clear_text(self, text):
        if text.startswith('\n'):
            text = text[1:]
        if text.startswith(','):
            text = text[1:]
        text = text.replace('\n', ', ')
        text = text.replace(' ,', ', ')
        text = self.replace_loop(text, '  ', ' ')
        return text.strip()


    def restore_text(self, text):
        tx = text.replace(', ', '\n')
        return tx


    def replace_loop(self, text, old_tx, new_tx):
        while True:
            n = len(text)
            text = text.replace(old_tx, new_tx)
            if n == len(text):
                break
        return text


    def replace_text_DOCX(self, doc, old_tx, new_tx):
        found = False
        for p in doc.paragraphs:
            if old_tx in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if old_tx in inline[i].text:
                        inline[i].text = inline[i].text.replace(old_tx, str(new_tx))
                        found = True
        if found:
            print('replace_text_DOCX: ' + str(old_tx) + ' -> ' + str(new_tx))
        else:
            print('replace_text_DOCX: ' + old_tx + ' NOT FOUND')


    def get_time_str(self):
        return str(datetime.datetime.now().strftime('%d.%m.%Y'))


    def save_document(self, nr_sprawy, nazwa_ulicy, result_dir_path, template_file_name, doc):
        nr_sprawy_list = nr_sprawy.split('.')
        if len(nr_sprawy_list) != 4:
            QtWidgets.QMessageBox.information(self, SCRIPT_TITLE, 'BŁĘDNY NUMER SPRAWY')
            return
        if len(nazwa_ulicy) > 20:
            nazwa_ulicy = nazwa_ulicy[0:20]
        nazwa_pliku_wynikowego = nr_sprawy_list[2] + ' - ' + template_file_name.replace('TEMPLATE', nazwa_ulicy)
        result_file_path = os.path.join(result_dir_path, nazwa_pliku_wynikowego)
        doc.save(result_file_path)
        print('save_document, zapisano do: ' + result_file_path)


    def is_present_in_adresaci_TAB(self, tx):
        sql = "SELECT adres FROM adresaci WHERE adres = '" + tx + "';"
        result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
        if result == 'ERROR':
            self.DOC_dB_connection_PostgreSQL.Restart_DB_connection()
            result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
        return result


    def is_present_in_przedmioty_TAB(self, tx):
        sql = "SELECT przedmiot FROM przedmioty WHERE przedmiot = '" + tx + "';"
        result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
        if result == 'ERROR':
            self.DOC_dB_connection_PostgreSQL.Restart_DB_connection()
            result = self.DOC_dB_connection_PostgreSQL.Send_SQL_to_DB(sql)
        return result


    def get_initial_automat_map(self):
        return {
                'Decyzja - TEMPLATE.docx': [],
                'Zgoda na przebudowę - TEMPLATE.docx': [],
                'Zezwolenie - TEMPLATE.docx': [],
                'Uzgodnienie + SŁUŻEBNOŚĆ - TEMPLATE.docx': [],
                'Uzgodnienie + PRZEKAZANIE - TEMPLATE.docx': [],
                'Uzgodnienie i opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx': [[], []],
                'Uzgodnienie i opinia  + PRZEKAZANIE - TEMPLATE.docx': [[], []],
                'Opinia + SŁUŻEBNOŚĆ - TEMPLATE.docx': [],
                'Opinia + PRZEKAZANIE - TEMPLATE.docx': []
            }


    def remove_empty_elements_from_automat_map(self, automat_map):
        if automat_map:
            for k in automat_map.copy().keys():
                if len(automat_map[k]) <= 1:
                    if not automat_map[k]:
                        del(automat_map[k])
                else:
                    if not automat_map[k][0] and not automat_map[k][1]:
                        del(automat_map[k])
            if automat_map:
                return automat_map
        return None


#====================================================================================================================


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    madeline = Madeline()
    madeline.show()
    sys.exit(app.exec_())



