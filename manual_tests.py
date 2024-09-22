"""
Gnumeric-py: Reading and writing gnumeric files with python
Copyright (C) 2017 Michael Lipschultz

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>.
"""

import os

from gnumeric import workbook


def write_workbook_with_one_worksheet(out_dir):
    filename = 'one_worksheet.gnumeric'
    wb = workbook.Workbook()
    wb.create_sheet('asdf')
    wb.save(os.path.join(out_dir, filename))


def test_order_of_cells_in_worksheet_does_not_matter(out_dir):
    filename = 'cell_order_worksheet.gnumeric'
    wb = workbook.Workbook()
    ws = wb.create_sheet('CellOrder')

    cell = ws.cell(2, 2)
    cell.value = '3:C'
    cell = ws.cell(0, 2)
    cell.value = '1:C'
    cell = ws.cell(0, 0)
    cell.value = '1:A'
    cell = ws.cell(0, 1)
    cell.value = '1:B'
    cell = ws.cell(1, 2)
    cell.value = '2:C'
    cell = ws.cell(1, 1)
    cell.value = '2:B'

    wb.save(os.path.join(out_dir, filename))


def test_assigning_wrong_value_type_to_cell(out_dir):
    pass


def test_saving_workbook_with_no_sheets(out_dir):
    pass


if __name__ == '__main__':
    test_dir = 'test_output'
    if not os.path.exists(test_dir):
        os.mkdir(test_dir)

    write_workbook_with_one_worksheet(test_dir)
    test_order_of_cells_in_worksheet_does_not_matter(test_dir)
