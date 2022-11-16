"""
/***************************************************************************
  Text_filter.py

  Parcel id and other info extraction.
  --------------------------------------
  version : 4.7
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
# pip install textract


import re
import os, time
from datetime import datetime
import codecs
import textract
import win32com.client


#======================================================================================================


class Text_filter:


    text = None


    def set_text(self, tx):
        if tx:
            tx = tx.replace("'", "")    # usuwanie "dziwnych" znaków
        self.text = tx


    def set_text_from_file(self, file_path):
        file_path = os.path.abspath(file_path)
        if str(file_path).lower().endswith('.txt'):
            input_file =  codecs.open(file_path, 'r', 'utf-8')
            tx = input_file.read()
            input_file.close()
        elif str(file_path).lower().endswith('.docx') or str(file_path).lower().endswith('.odt'):
            tx = textract.process(file_path)
            tx = tx.decode('utf-8')
        elif str(file_path).lower().endswith('.doc'):
            wordapp = win32com.client.Dispatch("Word.Application")
            wordapp.Documents.Open(file_path)
            doc = wordapp.ActiveDocument
            tx = ''
            for paragraph in doc.Paragraphs:
                tx += paragraph.Range.Text + '\n'
            doc.Close()
            wordapp.Quit()
        else:
            tx = None
        self.set_text(tx)
    

    def get_parcel_list(self):
        if self.text:
            word_list = self.get_word_list(self.text)
##            print('word_list: ' + str(word_list))
            parcel_list = []
            tmp_nr_list = []
            tmp_obr = None
            mode = None
            blanks = 0
            for word in word_list:
##                print('\n')
##                print('word: ' + str(word))
##                print('mode: ' + str(mode))
##                print('tmp_obr: ' + str(tmp_obr))
##                print('tmp_nr_list: ' + str(tmp_nr_list))
##                print('blanks: ' + str(blanks))
                if mode == 'NR' and not tmp_nr_list and self.is_parcel_id(word):
                    parcel_list.append(self.remove_leading_zeros_from_string(word))
                    tmp_nr_list = []
                    tmp_obr = None
                    blanks = 0
                elif word == 'OBR':
                    mode = 'OBR'
                    if tmp_obr and tmp_nr_list:
                        for tx2 in tmp_nr_list:
                            parcel_list.append(tmp_obr + '-' + tx2)
                        tmp_nr_list = []
                        tmp_obr = None
                        blanks = 0
                elif word == 'NR':
                    mode = 'NR'
                    if tmp_obr and tmp_nr_list:
                        for tx2 in tmp_nr_list:
                            parcel_list.append(tmp_obr + '-' + tx2)
                        tmp_obr = None
                        blanks = 0
                    tmp_nr_list = []
                elif word == 'ZAL':
                    mode = 'ZAL'
                    tmp_nr_list = []
                    tmp_obr = None
                elif self.is_parcel_no(word):
                    if mode == 'OBR' and word.isdigit():
                        tmp_obr = self.remove_leading_zeros_from_string(word)
                        if tmp_nr_list:
                            for tx2 in tmp_nr_list:
                                parcel_list.append(tmp_obr + '-' + tx2)
                            tmp_nr_list = []
                            tmp_obr = None
                            mode = None
                            blanks = 0
                    elif mode == 'NR':
                        tmp_nr_list.append(word)
                    else:
                        mode = 'NR'
                        tmp_nr_list.append(word)
                elif self.is_parcel_id(word) and mode != 'ZAL':
                    parcel_list.append(self.remove_leading_zeros_from_string(word))
                else:
                    blanks += 1
                    if blanks >= 2:
                        if tmp_obr and tmp_nr_list:
                            for tx2 in tmp_nr_list:
                                parcel_list.append(tmp_obr + '-' + tx2)
                        tmp_nr_list = []
                        tmp_obr = None
                        mode = None
                        blanks = 0
            if tmp_obr and tmp_nr_list:
                for tx2 in tmp_nr_list:
                    parcel_list.append(tmp_obr + '-' + tx2)
            return list(dict.fromkeys(parcel_list))    # usuwa duplikaty
        else:
            print('get_parcel_list: self.text IS EMPTY')
        return None


    def get_file_last_modification_time(self, file_path):   # output in format: '2001-02-16 20:38:40'
        date_time = datetime.fromtimestamp(os.path.getmtime(file_path))
        return date_time.strftime("%Y-%m-%d %H:%M:%S")


    def get_date_string_from_text(self):      # output in format: 'dd.mm.YYYY'
        if self.text:
            tx = self.text
            pattern_object = re.compile(r"Olsztyn,? (dnia )?(\d.+)", re.IGNORECASE)
            res = pattern_object.search(tx)
            if res:
##                print('res(1): ' + str(res))
                date = res.group(2)
                return self.convert_date(date)
            else:
                pattern_object = re.compile(r"Olsztyn,? (\d{4}\.\d{2}\._+)", re.IGNORECASE)
                res = pattern_object.search(tx)
                if res:
##                    print('res(2): ' + str(res))
                    date = res.group(1)
                    return self.convert_date(date)
                else:
                    pattern_object = re.compile(r"z dnia (.+)", re.IGNORECASE)
                    res = pattern_object.search(tx)
                    if res:
##                        print('res(3): ' + str(res))
                        date = res.group(1)
                        return self.convert_date(date)
        else:
            print('get_date_string_from_text: self.text IS EMPTY')
        return None


    def get_timestamp_formatted_date_string(self, date_str):   # output in format: 'YYYY-mm-dd'
        try:
            date_str_list = date_str.split('.')
            if date_str_list and len(date_str_list) == 3:
                return date_str_list[2] + '-' + date_str_list[1] + '-' + date_str_list[0]
        except:
##            print('get_timestamp_formatted_date_string ERROR for date_str = ' + str(date_str))
            pass
        return None


    def get_time_from_string(self, date_str):   # date_str in format: 'dd.mm.YYYY'
        try:
            return time.mktime(datetime.strptime(date_str, "%d.%m.%Y").timetuple())
        except:
##            print('get_time_from_string ERROR for date_str = ' + str(date_str))
            pass
        return None


    def get_doc_id(self):
        if self.text:
            for doc_id_key_re in self.DOC_ID_KEY_RE_TUPLE:
                pattern_object = re.compile(doc_id_key_re, re.IGNORECASE)
                res = pattern_object.search(self.text)
                if res:
                    doc_id = res.group(1).strip()
                    if len(doc_id) < 30:
                        return doc_id
                    else:
                        doc_id = self.replace_loop(doc_id, '\t\t', '\t')
                        doc_id = doc_id.replace('\t', ' ')
                        doc_id = self.replace_loop(doc_id, '  ', ' ')
                        return doc_id.split(' ')[0]
        else:
            print('get_doc_id: self.text IS EMPTY')
        return None
        

    #--------------------------------------------------------------------------------------------------
    # auxiliary functions:
    

    DOC_ID_KEY_RE_TUPLE = (
        r'Decyzja nr:?(.+)',
        r'Postanowienie nr:?(.+)',
        r'Znak:?(.+)',
        r'Nr sprawy:?(.+)',
        r'Znak sprawy:?(.+)',
        r'Umowa dierżawy nr:?(.+)',
        r'Umowa użyczenia nr:?(.+)',
        r'Umowa nr:?(.+)',
        r'POROZUMIENIE nr:? (.+)',
        r'(TE[-\.].+)'
        )


    MONTH_DICT = {
            'stycznia':'01',
            'styczeń':'01',
            'lutego':'02',
            'luty':'02',
            'marca':'03',
            'marzec':'03',
            'kwietnia':'04',
            'kwiecień':'04',
            'maja':'05',
            'maj':'05',
            'czerwca':'06',
            'czerwiec':'06',
            'lipca':'07',
            'lipiec':'07',
            'sierpnia':'08',
            'sierpień':'08',
            'września':'09',
            'wrzesień':'09',
            'października':'10',
            'październik':'10',
            'listopada':'11',
            'listopad':'11',
            'grudnia':'12',
            'grudzień':'12'
        }


    def convert_date(self, date):   # to format: 'dd.mm.YYYY'
##        print('date = ' + date)
        date = date.replace('roku', '')
        date = date.replace('rok', '')
        date = date.replace('r.', '')
        if date.endswith('r'):
            date = date[0:-1]
        date = self.replace_loop(date, '__', '_')
        if '_._' in date or '_-_' in date:
            return None
        date = date.replace('_', '15')
        date = date.strip()
        date_list = None
        if '.' in date:
            date_list = date.split('.')
        elif '-' in date:
            date_list = date.split('-')
        elif ' ' in date:
            date_list = date.split(' ')
##        print('date_list = ' + str(date_list))
        if date_list:
            if len(date_list) == 3 and date_list[0].isdigit() and date_list[2].isdigit():
                if not date_list[1].isdigit():
                    try:
                        date_list[1] = self.MONTH_DICT[date_list[1].strip()]
                    except:
                        date_list[1] = ''
                for i in range(3):
                    if len(date_list[i]) == 1 and date_list[i].isdigit():
                        date_list[i] = '0' + date_list[i]
                if len(date_list[0]) == 2 and len(date_list[1]) == 2 and len(date_list[2]) == 4:
                    return date_list[0] + '.' + date_list[1] + '.' + date_list[2]
                if len(date_list[0]) == 4 and len(date_list[1]) == 2 and len(date_list[2]) == 2:
                    return date_list[2] + '.' + date_list[1] + '.' + date_list[0]
        return None
        


    OBR_KEY_WORDS_TUPLE = ('w obrębie ewidencyjnym miasta olsztyn', 'położona w obrębie nr', 'położona w obrębie', 'położone w obrębie nr', 'położone w obrębie', 'w obrębie nr', 'w obrębie', 'nr obręb nr', 'obręb nr', 'obrębu', 'obrębie', 'obręb', 'obr')
    DZ_KEY_WORDS_TUPLE = ('nieruchomością', 'nieruchomościami', 'nieruchomościach', 'nieruchomości', 'działce', 'działkę', 'działkach', 'działkami', 'działka', 'działki', 'działek', 'dz')
    NR_KEY_WORDS_TUPLE = ('numerem', 'numery', 'numerach', 'numer', 'nr')
    ZAL_KEY_WORDS_TUPLE = ('załączników', 'załącznikiem', 'załącznika', 'załącznik', 'kv-zał nr', 'zał nr', 'zał')


    def get_word_list(self, tx):
##        print('tx(1) = ' + tx)
        tx = tx.lower()
        tx = self.replace_loop(tx, '  ', ' ')
        tx = tx.replace('\n', ' ')
        tx = tx.replace('\r', ' ')
        tx = tx.replace('\t', ' ')
        tx = tx.replace('≤', ' ')
        tx = tx.replace('<', ' ')
        tx = tx.replace('>', ' ')
        tx = tx.replace('[', ' ')
        tx = tx.replace(']', ' ')
        tx = tx.replace('{', ' ')
        tx = tx.replace('}', ' ')
        tx = tx.replace('(', ' ')
        tx = tx.replace(')', ' ')
        tx = tx.replace('+', ' ')
        tx = tx.replace('=', ' ')
        tx = tx.replace('_', ' ')
        tx = tx.replace('?', ' ')
        tx = tx.replace('§', ' ')
        tx = tx.replace('.', ' ')
        tx = tx.replace(',', ' ')
        tx = tx.replace(':', ' ')
        tx = tx.replace(';', ' ')
        tx = tx.replace("'", "")
        tx = tx.replace('"', '')
        tx = tx.replace('”', '')
        tx = tx.replace('„', '')
        tx = tx.replace('–', '-')   # inny ninus!
        fd_list = re.findall(r"\d\s?-\s?\d", tx)
        for fd in fd_list:
            tx = tx.replace(fd, fd.replace(' ', ''))
        tx = tx.replace(' lub ', ' ')
        tx = tx.replace(' albo ', ' ')
        tx = tx.replace(' i ', ' ')
        tx = tx.replace(' oraz ', ' ')
        tx = tx.replace(' a także ', ' ')
        tx = tx.replace(' ponadto ', ' ')
        tx = tx.replace(' ponad to ', ' ')
        tx = tx.replace('\xa0', '')
        tx = ' ' + tx + ' '
        tx = self.replace_loop(tx, '  ', ' ')
##        print('tx(2) = ' + tx)
        for word in self.ZAL_KEY_WORDS_TUPLE:
            tx = tx.replace(' ' + word + ' ', ' ZAL ')
        for word in self.OBR_KEY_WORDS_TUPLE:
##            print('word(OBR): ' + word)
            tx = tx.replace(' ' + word + ' ', ' OBR ')
##            print('tx: ' + tx)
        for word in self.DZ_KEY_WORDS_TUPLE:
            tx = tx.replace(' ' + word + ' ', ' DZ ')
        for word in self.NR_KEY_WORDS_TUPLE:
            tx = tx.replace(' ' + word + ' ', ' NR ')
        tx = tx.strip()
        tx = self.replace_loop(tx, '--', '-')
        word_list = []
        for word in tx.split(' '):
            if not word.islower() and word != '–':
                word_list.append(word)
            else:
                word_list.append(' ')
        tmp_tx = ';'.join(word_list)
        tmp_tx = tmp_tx.replace('OBR;NR', 'OBR')
        tmp_tx = tmp_tx.replace('OBR; ;NR', 'OBR')
        
        tmp_tx = tmp_tx.replace('DZ;NR', 'NR')
        tmp_tx = tmp_tx.replace('DZ; ;NR', 'NR')
        tmp_tx = tmp_tx.replace('NR;DZ', 'NR')
        tmp_tx = tmp_tx.replace('NR; ;DZ', 'NR')
        tmp_tx = tmp_tx.replace('DZ', 'NR')
        tmp_tx = self.replace_loop(tmp_tx, 'NR;NR', 'NR')
        tmp_tx = self.replace_loop(tmp_tx, 'NR; ;NR', 'NR')
        tmp_tx = tmp_tx.replace('ZAL;NR', 'ZAL')
        tmp_tx = tmp_tx.replace('ZAL; ;NR', 'ZAL')
        tmp_tx = tmp_tx.replace('NR;ZAL', 'ZAL')
        tmp_tx = tmp_tx.replace('NR; ;ZAL', 'ZAL')
        return tmp_tx.split(';')


    def is_parcel_no(self, tx):
        if tx.isdigit():
            return True
        else:
            tx2_list = tx.split('/')
            if len(tx2_list) == 2 and tx2_list[0].isdigit() and tx2_list[1].isdigit():
                return True
        return False


    def is_parcel_id(self, tx):
        tx_list = tx.split('-')
        if len(tx_list) == 2 and tx_list[0].isdigit():
            return self.is_parcel_no(tx_list[1])
        return False


    def replace_loop(self, text, old_tx, new_tx):
        while True:
            n = len(text)
            text = text.replace(old_tx, new_tx)
            if n == len(text):
                break
        return text


    def remove_leading_zeros_from_string(self, tx):
        while tx.startswith('0'):
            tx = tx[1:]
        return tx
        

#======================================================================================================
# Test:


def test():
    INPUT_FILE_PATH = os.path.join('doc_for_TEST', 'TXT', 'działki.txt')
    print('INPUT_FILE_PATH: ' + INPUT_FILE_PATH)
    text_filter = Text_filter()
    start_time = time.time()
    text_filter.set_text_from_file(INPUT_FILE_PATH)
    load_time = time.time() - start_time
    parcel_list = text_filter.get_parcel_list()
    print('parcel_list: ' + str(parcel_list))
    date = text_filter.get_date_string_from_text()
    print('date: ' + str(date))
    timestamp_formatted_date = text_filter.get_timestamp_formatted_date_string(date)
    print('timestamp_formatted_date: ' + str(timestamp_formatted_date))
    time_from_date = text_filter.get_time_from_string(date)
    print('time_from_date: ' + str(time_from_date))
    file_last_modification_time = text_filter.get_file_last_modification_time(INPUT_FILE_PATH)
    print('file_last_modification_time: ' + str(file_last_modification_time))
    doc_id = text_filter.get_doc_id()
    print('doc_id: ' + str(doc_id))
    print('load_time = ' + str(load_time))
    print('\n')
##    print('text_filter.text: ' + text_filter.text)

if __name__ == "__main__":
    test()

