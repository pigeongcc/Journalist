import sys
from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Cm
from docx.shared import Pt
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog


self.open_file()

def open_file(self):
    filename = QFileDialog.getOpenFileName(self, "Открыть файл Word", '/home')
    if filename[0]:
        self.FILENAME = filename[0]
        self.word = Document(self.FILENAME)
        self.calculations()


def calculations(self):
        wordTable = self.word.tables[0]
        quantityOfColumns = len(wordTable.rows[0].cells)
        quantityOfDisciplines = quantityOfColumns - 4
        quantityOfLines = len(wordTable.rows)
        quantityOfCadets = quantityOfLines - 2

        #self.averagesCadet = [None for i in range(quantityOfCadets)]
        #self.averagesDiscipline = [None for i in range(quantityOfDisciplines)]

        self.grades = [[None for i in range(1, quantityOfCadets)] for j in range(1+2, quantityOfDisciplines+2)]
        for i in range(1, quantityOfCadets):
            for j in range(1+2, quantityOfDisciplines+2):
                self.grades[i][j] = wordTable.rows[i].cells[j].text
                print(self.grades[i][j])
            print()
            
        

#        for row in wordTable.rows[3:(quantityOfColumns - 2)]:
            row[quantityOfColumns - 1] = self.averageCadet



#def averageC(person):
            sum = 0
            num = 0
            for j in range(3, self.N):
                sumToAdd, numToAdd = stringAnalyseC(person, j)
                sum += sumToAdd
                num += numToAdd
            try:
                return round(sum / num, 2)
            except ZeroDivisionError:
                return 0

#def stringAnalyseC(person, discipline):
            sum = 0
            num = 0
            line = self.grades[person][discipline].text()
            for digit in line:
                try:
                    sum += int(digit)
                    num += 1
                except:
                    continue
            return sum, num