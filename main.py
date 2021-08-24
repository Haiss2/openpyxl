import string
import datetime as dt

from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook
from openpyxl.chart import PieChart, ProjectedPieChart, Reference
from openpyxl.chart.series import DataPoint

from cover_page import CoverPage 
from follower_analysis import FollowerAnalysis
from m_o_m import MoM

wb = Workbook()
std = wb.get_sheet_by_name('Sheet')
wb.remove_sheet(std)

CoverPage(wb).render()
FollowerAnalysis(wb).render()
MoM(wb).render()

wb.save("reboot_report.xlsx")
