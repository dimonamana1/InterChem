# -*- coding: utf-8 -*-
import csv
import os
import re
import tempfile
import zipfile
from difflib import SequenceMatcher
import openpyxl
import xlrd
from xlsxwriter import Workbook


def fix_xlsx(in_file):
    zin = zipfile.ZipFile(in_file, 'r')
    if 'xl/SharedStrings.xml' in zin.namelist():
        tmpfd, tmp = tempfile.mkstemp(dir=os.path.dirname(in_file))
        os.close(tmpfd)

        with zipfile.ZipFile(tmp, 'w') as zout:
            for item in zin.infolist():
                if item.filename == 'xl/SharedStrings.xml':
                    zout.writestr('xl/sharedStrings.xml', zin.read(item.filename))
                else:
                    zout.writestr(item, zin.read(item.filename))

        zin.close()
        os.remove(in_file)
        os.rename(tmp, in_file)


class Matcher:
    __rfilepath = "Препараты.csv"  # drugs filepath
    __rfilepath2 = "allbd.csv"  # garbage filepath
    __dictdb = {}  # drugs dict
    __dictallbd = {}  # garbage dict
    __directory = "C:/Users/4r4r5/Desktop/reports/"  # directory to work with
    __reportsdirectory = 'C:/Users/4r4r5/Desktop/reports_result/'  # directory to save reports
    __files = os.listdir(__directory)

    @staticmethod
    def __csv_writer(path, data):
        with open(path, "a", newline='') as csv_file:
            writer = csv.writer(csv_file, delimiter=';')
            writer.writerow(data)

    @staticmethod
    def __is_digit(string):
        if string.isdigit():
            return True
        else:
            try:
                float(string)
                return True
            except ValueError:
                return False

    @staticmethod
    def __csv_from_excel(filename, datapath):  # read from datapath; write to filename
        try:
            if (datapath[-4:]) == "xlsx":
                try:
                    wb = openpyxl.load_workbook(datapath)
                except KeyError:
                    fix_xlsx(filename)
                    wb = openpyxl.load_workbook(datapath)
                filenames = []
                sheets = wb.sheetnames
                for sheet in sheets:
                    sh = wb[str(sheet)]
                    with open(filename[:-4] + sheet + ".csv", 'tw', newline='') as f:
                        filenames.append(filename[:-4] + sheet + ".csv")
                        wr = csv.writer(f, quoting=csv.QUOTE_ALL, delimiter=";")
                        for row in sh.rows:
                            temp = []
                            for cell in row:
                                temp.append(cell.value)
                            wr.writerow(temp)
                        f.close()
                return filenames
            elif (datapath[-4:]) == ".xls":
                wb = xlrd.open_workbook(datapath)
                sheetlist = wb.sheet_names()
                filenames = []
                for sheet in sheetlist:
                    sh = wb.sheet_by_name(sheet)
                    with open(filename[:-4] + sheet + ".csv", 'tw', newline='') as f:
                        filenames.append(filename[:-4] + sheet + ".csv")
                        wr = csv.writer(f, quoting=csv.QUOTE_ALL, delimiter=";")
                        for rownum in range(sh.nrows):
                            wr.writerow(sh.row_values(rownum))
                        f.close()
                return filenames
        except UnicodeDecodeError:
            print(datapath)

    @staticmethod
    def setrfilepath(filepath, filepath2):
        Matcher.__rfilepath = filepath
        Matcher.__rfilepath2 = filepath2

    @staticmethod
    def setdirectory(directory):
        Matcher.__directory = directory

    @staticmethod
    def match(rpath, wpath):  # read form rpath and write to wpath with renamed drugs
        with open(rpath, "r", newline="") as file:
            reader = csv.reader(file, delimiter=';', quotechar='"')
            for row2 in reader:
                for i in range(len(row2)):
                    word = row2[i]
                    temp = re.sub(r"[#()./|i*®^ї'%і_\-$!~=\"]", "", word.lower())
                    curpos = temp.strip()
                    if Matcher.__is_digit(curpos) or curpos == "" or (curpos in Matcher.__dictallbd):
                        continue
                    elif curpos in Matcher.__dictdb:
                        row2[i] = Matcher.__dictdb[curpos]
                Matcher.__csv_writer(wpath, row2)

    @staticmethod
    def __exelwriter(csvfile):  # creates same named as csvfile xlsx and write data to it
        workbook = Workbook(Matcher.__reportsdirectory + csvfile[:-3] + '.xlsx')
        worksheet = workbook.add_worksheet()
        with open(csvfile, 'rt', newline="") as f:
            reader = csv.reader(f, delimiter=';')
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
        workbook.close()

    def rename_drugs(self):  # unites some methods:
        with open(self.__rfilepath, "r", newline="") as file:  # create drugs dict
            reader = csv.reader(file, delimiter=';', skipinitialspace=True)
            for row1 in reader:
                self.__dictdb[row1[0]] = row1[1]
        with open(self.__rfilepath2, "r", newline="") as file:  # create garbage dict
            reader = csv.reader(file, delimiter=';', skipinitialspace=True)
            for row2 in reader:
                self.__dictallbd[row2[0]] = row2[0]
        for file in self.__files:
            # print(file[-4:])
            filenames = Matcher.__csv_from_excel(file,
                                                 self.__directory + file)  # for files from directory creates csv copy
            for filenm in filenames:  # for files from created list creates result files with renamed drugs
                with open(filenm[:-4] + "res" + ".csv", 'tw', newline=''):  # and writes result to xslx
                    w = filenm[:-4] + "res" + ".csv"
                    self.match(filenm, w)
                    self.__exelwriter(w)
            for filenm in filenames:  # deletes temporary files
                w = filenm[:-4] + "res" + ".csv"
                os.remove(filenm)
                os.remove(w)

    def createbd(self, rfilepath1):  # add drugs from filepath1 to drugs dict
        with open(self.__rfilepath, "r", newline="") as file:
            reader = csv.reader(file, delimiter=';', skipinitialspace=True)
            for row1 in reader:
                self.__dictdb[row1[0]] = row1[1]
        with open(rfilepath1, "r", newline="") as file:
            reader = csv.reader(file, delimiter=';', skipinitialspace=True)
            for row2 in reader:
                temp = re.sub(r"[#()./|i*®^ї'%і_\-$!~=\"]", "", row2[0].lower())
                curpos = temp.strip()
                if curpos in self.__dictdb:
                    continue
                else:
                    mlist = []
                    for i in self.__dictdb:
                        mlist.append([SequenceMatcher(None, curpos, i).ratio(), i])
                    mlist.sort(reverse=True)
                    print(curpos, " = ", self.__dictdb[mlist[0][1]], "?")
                    uinput1 = input()
                    if uinput1 == "y":
                        data = [curpos, self.__dictdb[mlist[0][1]]]
                        self.__csv_writer(self.__rfilepath, data)
                        self.__dictdb[curpos] = data[1]
                        print("ok")
                    else:
                        dictemp = {"1": self.__dictdb[mlist[1][1]],
                                   "2": self.__dictdb[mlist[2][1]],
                                   "3": self.__dictdb[mlist[3][1]]}
                        print("choose one: \n", self.__dictdb[mlist[1][1]], '\n', self.__dictdb[mlist[2][1]], '\n',
                              self.__dictdb[mlist[3][1]])
                        uinput2 = input()
                        if uinput2 == "1" or uinput2 == "2" or uinput2 == "3":
                            data = [curpos, dictemp.get(uinput2)]
                            self.__csv_writer(self.__rfilepath, data)
                            self.__dictdb[curpos] = data[1]
                            print("ok")
                        else:
                            print("insert your own:")
                            uinput3 = input()
                            data = [curpos, uinput3]
                            self.__csv_writer(self.__rfilepath, data)
                            self.__dictdb[curpos] = data[1]
                            print("ok")
