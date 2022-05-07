import openpyxl
from openpyxl import load_workbook
import random

print(f'{0:06X}{random.randint(0, 0xFFFFF)}')
print('{0:06X}'.format(random.randint(0, 0xFFFFF)))


def monthlize(path: str):
    print(f'打开文件为{path}')
    wb = load_workbook(path)
    for i in range(12):
        ws = wb.copy_worksheet(wb.active)
        ws.title = f'{i + 1}月{wb.active.title}'
        ws.sheet_properties.tabColor = '{0:06X}'.format(random.randint(0, 0xFFFFF))
        print(f'{ws.title}的颜色为{wb.active.title}')
    wb.remove(wb.active)
    print(path.split('.')[0])
    wb.save(path.split('.')[0] + 'monthly.xlsx')
    print(f"运行完毕!\n文件保存在:{path.split('.')[0] + 'monthly.xlsx'}")


# F:\目标与计划\财务\learn\python-openpyxl-files\video-file\001 访问目录下所有Excel文件\4月报表\4月资产负债表.xlsx
monthlize('F:\\目标与计划\\财务\\learn\\python-openpyxl-files\\video-file\\002 工作簿与工作表操作\\sales.xlsx')
# monthlize('F:\\目标与计划\\财务\\learn\\python-openpyxl-files\\video-file\\001 访问目录下所有Excel文件\\4月报表\\4月资产负债表.xlsx')
