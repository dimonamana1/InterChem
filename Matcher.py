import csv
from difflib import SequenceMatcher


class Matcher:
    rfilepath = ""
    rfilepath1 = ""
    wfilepath = ""
    __dictdb = {}

    def __init__(self, rfilepath, rfilepath1, wfilepath):
        self.rfilepath = rfilepath
        self.rfilepath1 = rfilepath1
        self.wfilepath = wfilepath

    @staticmethod
    def __csv_writer(self, data):
        with open(self.wfilepath, "a", newline='') as csv_file:
            writer = csv.writer(csv_file, delimiter=';')
            writer.writerow(data)

    def match(self):
        with open(self.rfilepath, "r", newline="") as file:
            reader = csv.reader(file, delimiter=';')
            for row1 in reader:
                self.__dictdb[row1[1]] = row1[0]
        with open(self.rfilepath1, "r", newline="") as file:
            reader = csv.reader(file, delimiter=';')
            for row2 in reader:
                curpos = row2[0]
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
                        data = [self.__dictdb[mlist[0][1]], curpos]
                        self.__csv_writer(self, data)
                        print("ok")
                    else:
                        dictemp = {"1": self.__dictdb[mlist[1][1]],
                                   "2": self.__dictdb[mlist[2][1]],
                                   "3": self.__dictdb[mlist[3][1]]}
                        print("choose one: \n", self.__dictdb[mlist[1][1]], '\n', self.__dictdb[mlist[2][1]], '\n',
                              self.__dictdb[mlist[3][1]])
                        uinput2 = input()
                        if uinput2 == "1" or uinput2 == "2" or uinput2 == "3":
                            data = [dictemp.get(uinput2), curpos]
                            self.__csv_writer(self, data)
                            print("ok")
                        else:
                            print("insert your own:")
                            uinput3 = input()
                            data = [uinput3, curpos]
                            self.__csv_writer(self, data)
                            print("ok")
