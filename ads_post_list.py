
import string
import datetime as dt

from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook


class AdsPostList():

    def __init__(self, wb):
        self.wb = wb
    
    def render(self):
        wb = self.wb
        cover_page = wb.create_sheet("Ads Post List") 