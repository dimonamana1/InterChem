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
from openpyxl import Workbook


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
    def __is_garbage(string):
        digit = 0
        non_digit = 0
        for i in string:
            if i.isdigit():
                digit += 1
            else:
                non_digit += 1
        if digit / (digit + non_digit) > 0.4:
            return True
        else:
            return False

    @staticmethod
    def __response_handler(curpos, mlist):
        try:
            while len(mlist) < 4:
                mlist.append([0, "Error"])
            print(curpos, " = ", Matcher.__dictdb[mlist[0][1]], "?")
            uinput1 = input()
            if uinput1 == "y":
                data = [curpos, Matcher.__dictdb[mlist[0][1]]]
                Matcher.__csv_writer(Matcher.__rfilepath, data)
                Matcher.__dictdb[curpos] = data[1]
                print("ok")
            elif uinput1 == "garb":
                Matcher.__csv_writer(Matcher.__rfilepath2, [curpos])
                Matcher.__dictallbd[curpos] = curpos
                print("added to garbage")
            else:
                dictemp = {"1": Matcher.__dictdb[mlist[1][1]],
                           "2": Matcher.__dictdb[mlist[2][1]],
                           "3": Matcher.__dictdb[mlist[3][1]]}
                print("choose one: \n", Matcher.__dictdb[mlist[1][1]], '\n', Matcher.__dictdb[mlist[2][1]], '\n',
                      Matcher.__dictdb[mlist[3][1]])
                uinput2 = input()
                if uinput2 == "1" or uinput2 == "2" or uinput2 == "3":
                    data = [curpos, dictemp.get(uinput2)]
                    Matcher.__csv_writer(Matcher.__rfilepath, data)
                    Matcher.__dictdb[curpos] = data[1]
                    print("ok")
                else:
                    print("insert your own:")
                    uinput3 = input()
                    data = [curpos, uinput3]
                    Matcher.__csv_writer(Matcher.__rfilepath, data)
                    Matcher.__dictdb[curpos] = data[1]
                    print("ok")
        except IndexError:
            print("list out of bonds")

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
    def setreportsdirectory(directory):
        Matcher.__reportsdirectory = directory

    @staticmethod
    def match(rpath, wpath):  # read form rpath and write to wpath with renamed drugs
        with open(rpath, "r", newline="") as file:
            reader = csv.reader(file, delimiter=';', quotechar='"')
            for row2 in reader:
                for i in range(len(row2)):
                    word = row2[i]
                    if Matcher.__is_digit(word):
                        word = word.replace('.', ',')
                        row2[i] = word
                    temp = re.sub(r"[#()./|i*®^ї'%і_\-$!~=\"]", "", word.lower())
                    curpos = temp.strip()
                    if curpos in Matcher.__dictdb:
                        row2[i] = Matcher.__dictdb[curpos]
                    elif Matcher.__is_digit(curpos) or curpos == "" or (
                            curpos in Matcher.__dictallbd) or Matcher.__is_garbage(curpos):
                        continue
                    else:
                        mlist = []
                        for k in Matcher.__dictdb:
                            a = SequenceMatcher(None, curpos, k).ratio()
                            if a > 0.6:
                                mlist.append([a, k])
                        if len(mlist) == 0:
                            Matcher.__csv_writer(Matcher.__rfilepath2, [curpos])
                            Matcher.__dictallbd[curpos] = curpos
                        if len(mlist) != 0:
                            mlist.sort(reverse=True)
                            Matcher.__response_handler(curpos, mlist)
                Matcher.__csv_writer(wpath, row2)

    @staticmethod
    def __exelwriter(csvfile):  # creates same named as csvfile xlsx and write data to it
        wb = Workbook(Matcher.__reportsdirectory + csvfile[:-3] + '.xlsx')
        ws = wb.create_sheet("kaban")
        with open(csvfile, 'rt', newline="") as f:
            reader = csv.reader(f, delimiter=';')
            for subarray in reader:
                ws.append(subarray)
        wb.save(Matcher.__reportsdirectory + csvfile[:-3] + '.xlsx')

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
                    Matcher.__response_handler(curpos, mlist)
