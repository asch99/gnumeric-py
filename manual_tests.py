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

def test_read_numbers_only(fpath):
    """
    read only numbers from sheet
    """
    import string

    wb = workbook.Workbook.load_workbook(fpath)

    # Get a sheet by name (skip dummy sheets)
    ws = wb.get_sheet_by_name('info')
    
    truth = []
    for i in range(1,11):
        a =i
        b = 3*i+1
        c = b
        if i>1:
            c = b+truth[-1][-1]
        truth.append([a,b,c])
        
    # construct index like 'B12'
    c0 = string.ascii_uppercase.index('A') # first column
    c1 = string.ascii_uppercase.index('C') # last column
    clist = list(string.ascii_uppercase[c0:c1+1])
    r0 = 2 # first row
    r1 = 11 # last row
    # construct index like 'B12'
    cnames = [ ws['{}{}'.format(c,1)].get_value() for c in clist ]# only data, no expressions
    data = []
    for r in range(r0,r1+1):
        dat = [ int(ws['{}{}'.format(c,r)].get_value(compute_expression=True)) for c in clist ]# only data, no expressions
        data.append(dat)

    assert (data == truth), "Imported data does not match with expected values"
    
if __name__ == '__main__':
    test_dir = 'test_output'
    if not os.path.exists(test_dir):
        os.mkdir(test_dir)

    write_workbook_with_one_worksheet(test_dir)
    test_order_of_cells_in_worksheet_does_not_matter(test_dir)
    test_read_numbers_only(os.path.join('samples', 'test_input.gnumeric'))
