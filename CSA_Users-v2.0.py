#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
#import csv
import os
import sys
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import xlwt
from xlwt import Workbook
import xlrd
from xlrd import open_workbook
import openpyxl


class Ui_Dialog(object):
    global path

    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(526, 300)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(95, 210, 341, 32))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Close)
        self.buttonBox.setCenterButtons(True)
        self.buttonBox.setObjectName("buttonBox")
        self.pushButton = QtWidgets.QPushButton(Dialog)
#        self.pushButton.setGeometry(QtCore.QRect(205, 50, 121, 28))
        self.pushButton.setGeometry(QtCore.QRect(153, 50, 340, 32))
 
        self.pushButton.setObjectName("pushButton")
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(13, 100, 500, 31))
        self.label.setFrameShape(QtWidgets.QFrame.Panel)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")

        self.label2 = QtWidgets.QLabel(Dialog)
        self.label2.setGeometry(QtCore.QRect(13, 135, 500, 31))
        self.label2.setFrameShape(QtWidgets.QFrame.Panel)
        self.label2.setAlignment(QtCore.Qt.AlignCenter)
        self.label2.setObjectName("label2")

        self.label3 = QtWidgets.QLabel(Dialog)
        self.label3.setGeometry(QtCore.QRect(13, 170, 500, 31))
        self.label3.setFrameShape(QtWidgets.QFrame.Panel)
        self.label3.setAlignment(QtCore.Qt.AlignCenter)
        self.label3.setObjectName("label2")

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.pushButton.setText(_translate("Dialog", " Click Here to Select a Data File "))
        self.pushButton.setStyleSheet("border: 1px solid blue; padding: 3px")
        self.pushButton.adjustSize()
        self.pushButton.setShortcut(_translate("Dialog", "Ctrl+O"))
        self.label.setText(_translate("Dialog", "<filepath>"))
        self.label2.setStyleSheet("border: 0px")
        self.label2.setText(_translate("Dialog", ""))
        self.label3.setStyleSheet("border: 0px")
        self.label3.setText(_translate("Dialog", ""))

        self.pushButton.clicked.connect(self.pushButton_handler)


    def pushButton_handler(self):
        self.open_dialog_box()
        
    def open_dialog_box(self):
        filename = QFileDialog.getOpenFileName(filter = ('*.csv; *.xls*'))
        # filename = QFileDialog.getOpenFileName(filter = ('*.csv'))
        path = filename[0]

        self.label.setText(path)
        
        print("path = ", path)

        OUTFILE = path
        
        OUTFILE_split = OUTFILE.rsplit(".", 1)
        print("OUTFILE_name = ", OUTFILE_split[0] + "_out." + OUTFILE_split[1]) #
        # print("Output file = ",OUTFILE) #
        
        
        if OUTFILE_split[1] == "csv":
            OUTFILE_name = OUTFILE_split[0]+"_out.csv"
            print("OUTFILE_name (CSV) = ", OUTFILE_name)

            """CSV Files"""
            ## Open & Read Data File
            df = pd.read_csv(OUTFILE, sep=(','), header=None, na_filter=False, iterator=False)
    
            ## Cleaning the dataframe (removing/renaming columns, first row data, etc.)
            df.columns
            df.drop([1,2,3,4], axis = 1, inplace = True)  #delete columns
            print("line 76", df)      
    
            # ## Remove blank rows
            nan_value = float("NaN")
            
            #Convert NaN values to empty string
            df.replace("", nan_value, inplace=True)
            
            df.dropna(subset = [0], inplace=True)
            
            # df.iloc[0] = ['Group Name: 41'] #add row only necessary if header not set to 'None'
            
            df.insert(1,'Group',0) #add column only necessary if 2nd column doesn't exist
            
            df.columns = ['Name', 'Group'] #rename columns
            
            # ## Creating the Check Column
            df.insert(2,'Check',0)
            
            # ## Extract Group Name into Group column
            df['Group'] = df['Name'].str.split(':').str[1]
    
    
            ## Assign numeric value to records that are null
            #null_check = all_data.Group.isnull()
            df.Check=df.Group.fillna(df.Check.replace({0:1}))
             
            
            ## Using the check to form iterator
            for i in range(len(df)) :
            
                if df.iloc[i, 2] != 1:
                    group_id = df.iloc[i, 2]
            
                if df.iloc[i, 2] == 1:
                    df.iloc[i, 1] = group_id
            
            ## Delete obsolete columns
            
            df.drop(['Check'], axis = 1, inplace = True)  #delete columns
    
            ## Delete Group Name Rows
            patternDel = "Group Name:*"
            filter = df['Name'].str.contains(patternDel)
            
            df = df[~filter]
            df = df.reset_index(drop = True)
            df.index = np.arange(1,len(df)+1)
            print(df.head(30))
    
            ## Saving CSV Output File
            path_out = OUTFILE_name        
            # path_out_full = path.rsplit("/",1)
            # path_out = path_out_full[0]
            print("path_out = ",path_out)
            df.to_csv(path_out)

            self.label2.setText("<font color = 'red'>'Your file is ready in the folder below. You may now close this window.")
    
            path_open = path_out.replace('/', '\\')
            ### / doesn't work; must be \ ###
    
            print("Path Open = ", path_open)
            self.label3.setText(path_open)
    
            ## Open File Automatically after run
            path_open = '"'+os.path.realpath(path_open)+'"'
            os.startfile(path_open)
            
        # print("file has been saved")


        elif OUTFILE_split[1] == "xlsx" or "xls":
            OUTFILE_ext = OUTFILE_split[1]
            print("file extension = ", OUTFILE_ext)
            # OUTFILE_split = OUTFILE.rsplit(".", 1)            
            
            OUTFILE_name = OUTFILE_split[0]+"_out."+OUTFILE_ext
            
            print("OUTFILE_name (Excel) = ", OUTFILE_name)        

            ## Open & Read Data File

            df = pd.read_excel(OUTFILE, header=None, index_col=False)    
            print(df.head(30))
            
            ## Cleaning the dataframe (removing/renaming columns, first row data, etc.)
            print("columns = ",df.columns)
            
            print("line 76", df.head(30))      
    
    
            # ## Remove blank rows
            # df.dropna(subset = [0], inplace=True)
            nan_value = float("NaN")
            
            # #Convert NaN values to empty string
            df.replace("", nan_value, inplace=True)
            df = df.reset_index(drop = True)
            
            # df.head(25)
            

            
            # df.iloc[0] = ['Group Name: 41'] #add row only necessary if header not set to 'None'
            
            df.insert(1,'Group', 0) #add column only necessary if 2nd column doesn't exist
            
            df.columns = ['Name', 'Group'] #rename columns
            
            print("columns = ", df.columns)
            print(df)
            
            # ## Creating the Check Column
            df.insert(2,'Check',0)
            
            # ## Extract Group Name into Group column
            df['Group'] = df['Name'].str.split(':').str[1]
            
            print("checkpoint Groups")
            print(df)

            # ## Remove blank rows
            # nan_value = float("NaN")
            # df = df.reset_index(drop = True)
    
            ## Assign numeric value to records that are null
            # null_check = all_data.Group.isnull()
            df.Check=df.Group.fillna(df.Check.replace({0:1}))
             
            print("Checkpoint NaN")
            print(df.head(30))
            
            
            ## Using the check to form iterator
            for i in range(len(df)) :
            
                if df.iloc[i, 2] != 1:
                    group_id = df.iloc[i, 2]

            
                if df.iloc[i, 2] == 1:
                    df.iloc[i, 1] = group_id
                    
            print(df.head(30))
            
            ## Delete obsolete columns
            
            df.drop(['Check'], axis = 1, inplace = True)  #delete columns
    
            ## Delete Group Name Rows
            patternDel = "Group Name:*"
            filter = df['Name'].str.contains(patternDel)

            
            df = df[~filter]
            df = df.reset_index(drop = True)
            df.index = np.arange(1,len(df)+1)
            print(df.head(30))
    
            ## Saving XLS* Output File
            path_out = OUTFILE_name        
            
            print("path_out = ",path_out)
            df.to_excel(path_out)

            self.label2.setText("<font color = 'red'>'Your file is ready in the folder below. You may now close this window.")
    
            path_open = path_out.replace('/', '\\')
            ### / doesn't work; must be \ ###
    
            print("Path Open = ", path_open)
            self.label3.setText(path_open)
    
            ## Open File Automatically after run
            path_open = '"'+os.path.realpath(path_open)+'"'
            print("Path Open is now = ", path_open)
            os.startfile(path_open)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())