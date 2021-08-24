import string
import datetime as dt

from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook
from openpyxl.chart import PieChart, ProjectedPieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.image import Image


class MoM():

    def __init__(self, wb):
        self.wb = wb
        self.thin = Side(border_style="thin", color="000000")
        
    def set_border(self, ws, cell_range):
        rows = ws[cell_range]
        for row in rows:
            row[0].border = Border(left=self.thin)
            row[-1].border = Border(right=self.thin)
        for c in rows[0]:
            c.border = Border(top=self.thin)
        for c in rows[-1]:
            c.border = Border(bottom=self.thin)
        rows[0][0].border = Border(top=self.thin, left=self.thin)
        rows[0][-1].border = Border(top=self.thin, right=self.thin)
        rows[-1][0].border = Border(bottom=self.thin, left=self.thin)
        rows[-1][-1].border = Border(bottom=self.thin, right=self.thin)

    def render_grid(self, ws, possition, img, unit_of_measure, title, value_of_6_months):
        self.set_border(ws, possition)
        img = Image('img/user.png')
        ws.add_image(img, 'A1')


    def render(self):
        wb = self.wb
        m_o_m = wb.create_sheet("MoM") 
        m_o_m.sheet_view.showGridLines = False
        
        
        for x in [*list(string.ascii_uppercase), 'AA', 'AB', 'AC', 'AD', 'AE']:
            m_o_m.column_dimensions[x].width = 10
        for y in range(1, 60):
            m_o_m.row_dimensions[y].height = 25

        self.render_grid(m_o_m, 'A3:G19', 'img', 'unit_of_measure', 'title', 'value_of_6_months')
            
        
        

    