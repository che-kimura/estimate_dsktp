###Excelのマドプロ加盟国からcountry_code.csvを作成する
import csv
import pprint
import openpyxl


def main():
    wb = openpyxl.load_workbook(r'マドプロ加盟国.xlsx')
    ws = wb['code_tbl']

    with open(r'country_code.csv', 'w', newline="",encoding='UTF-8') as csvfile:
        writer = csv.writer(csvfile,quotechar='"',quoting=csv.QUOTE_NONNUMERIC)
        for row in ws.rows:
            writer.writerow( [cell.value for cell in row] )
    
if __name__ == "__main__":
    main()