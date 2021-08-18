import os
from xlrd import open_workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment

import util


util.running_prerequisite()
logger = util.get_logger(__file__)


class Procurement:

    excluded_type = ['SUS', 'SEC', 'SEH', 'SUP']

    def __init__(self, main_excel, parse_local, bom_path, pallet_path):

        self.main_excel = main_excel
        self.parse_local = parse_local
        self.bom_path = bom_path
        self.pallet_path = pallet_path

        self.wb = load_workbook(main_excel, read_only=False)

    # get data
    def read_main_excel(self):
        table = self._get_worksheet()

        for row_index, row in enumerate(table.iter_rows(min_row=2, min_col=1, max_col=10, values_only=True)):
            if row[0] is not None:
                yield row

    # get specific data
    def get_required_data(self, file_name):
        for row in self.yield_bom_content(file_name):
            if row[7] == '' or row[7] is None:
                raise ValueError('{}格式異常'.format(file_name))
            elif row[7] not in self.excluded_type:
                if str(row[7]) != 'SPC':
                    print('1', row[7], row[1], row[3], row[4], row[4],)
                elif str(row[7]) == 'SPC' and float(row[3].split('\\')[0][:-1]) >= 12.0:
                    print('2', row[7], row[1], row[3], row[4], row[4], )

    # decide format and yield content
    def yield_bom_content(self, file_name):
        path = self.bom_path
        prefix = file_name[0]
        completed_path = os.path.join(path, prefix, file_name)

        if os.path.isfile(completed_path + '.xls'):
            target_wb = open_workbook(completed_path + '.xls')
            target_sh = target_wb.sheets()[0]
            for row in range(2, target_sh.nrows):
                if type(target_sh.row_values(row)[0]) != str:
                    yield target_sh.row_values(row)
        elif os.path.isfile(completed_path + '.xlsx'):
            target_wb = load_workbook(completed_path + '.xlsx', read_only=False)
            target_sh = target_wb[target_wb.sheetnames[0]]
            for row in target_sh.iter_rows(min_row=3, min_col=1, max_col=12, values_only=True):
                if row[0] is not None:
                    yield row
        # xls或xlsx都找不到的話就只能:
        else:
            return False






    def _get_worksheet(self):
        return self.wb[self.wb.sheetnames[0]]

    def save(self):
        self.wb.save(self.file_path)