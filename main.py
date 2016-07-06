import sys
from os import listdir
from os.path import isfile, join

from functools import reduce

from xlrd import open_workbook
from xlwt import Workbook

from core import add_sheets, sheet2cells

if __name__ == '__main__':

  if len(sys.argv) < 2:
    print('请输出文件夹名称')
    sys.exit(0)

  # 输入参数:文件夹名称
  dir = sys.argv[1]

  files = [dir + '/' + f for f in listdir(dir) if isfile(join(dir, f))]

  book = open_workbook(files[0])

  sheets = book.sheets()

  result = reduce(add_sheets, [sheet2cells(sheet) for sheet in sheets])

  output = Workbook()
  sheet = output.add_sheet('汇总')

  for r in range(len(result)):
    for c in range(len(result[r])):
      sheet.write(r, c, result[r][c])

  output.save(dir + '.xls')

