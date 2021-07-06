from PIL import ImageGrab
import pandas as pd
import numpy as np
import xlwings as xw
import win32api
from xlwings.constants import DeleteShiftDirection,InsertShiftDirection

class Resize_Table:
    """
    create new attribute for xlwings to resize tables by adding the right number of rows
    """
    def __init__(self,frame):
        self.frame = frame

    def __getitem__(self, item):
        table,header = item
        tableshape_rows = table.shape[0]
        excel_shape = self.frame.shape
        if header:
            self.frame.sheet.range(self.frame.row - 1, self.frame.column).options(np.array).value = table.columns

        upper_left_table = (self.frame.row + 1, self.frame.column)
        upper_right_table = (self.frame.row + 1, self.frame.column + excel_shape[1] - 1)

        # if need to add lines
        if table.shape[0] - excel_shape[0] > 0:
            print('add lines')
            rows_to_add = table.shape[0] - excel_shape[0]
            self.frame.sheet.range(upper_left_table,
                                (upper_right_table[0] + rows_to_add - 1, upper_right_table[1])).api.Insert(
                InsertShiftDirection.xlShiftDown)

        # remove lines
        elif table.shape[0] - excel_shape[0] < 0:
            print('delete lines')
            rows_to_remove = excel_shape[0] - table.shape[0]
            self.frame.sheet.range(upper_left_table,
                                (upper_right_table[0] + rows_to_remove - 1, upper_right_table[1])).api.Delete(
                DeleteShiftDirection.xlShiftUp)
        self.frame.options(pd.DataFrame, header=False, index=False).value = table

        return

class savepng:
    """
    create attribute to save excel range as picture
    """
    def __init__(self,frame):
        self.frame = frame

    def __getitem__(self, item):
        path = item
        self.frame.options(header=True).api.CopyPicture()
        self.frame.sheet.api.Paste()
        pic = self.frame.sheet.pictures[0]
        pic.api.Copy()
        pic.delete()
        img = ImageGrab.grabclipboard()
        img.save(path)

class MsgBox:
    def __init__(self,frame):
        self.frame = frame

    def __getitem__(self, item):
        title,text = item
        win32api.MessageBox(self.frame.app.hwnd, text, title)


# add the attributes to the xlwings library
xw.main.Range.save = property(lambda frame: savepng(frame))
xw.main.Range.xresize = property(lambda frame: Resize_Table(frame))
xw.main.books.MsgBox = property(lambda frame: MsgBox(frame))