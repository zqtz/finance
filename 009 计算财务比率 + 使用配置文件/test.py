from openpyxl import load_workbook
import configparser

def cal_current_ratio(workseet, config) -> str:
    total_liquid_asset = workseet[config['ratio']['liquid_asset']].value
    total_liquid_liability = workseet[config['ratio']['liquid_liability']].value
    return "{:.2f}%".format(total_liquid_asset / total_liquid_liability*100)


def load_config(file_name: str = None) -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    if not file_name:
        config.read('config.ini')
    else:
        config.read(file_name)
    return config

if __name__ == '__main__':
    wb = load_workbook('600734-2.xlsx')
    workseet = wb.active
    config = load_config()
    result = cal_current_ratio(workseet,config)
    print(result)
