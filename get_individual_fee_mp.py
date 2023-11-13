import csv
import re

#from hashlib import file_digest
import urllib.request
from bs4 import BeautifulSoup

#---------------------------------------------------------
# WIPOのIndividual Feeのページにアクセスする
url = f"<https://www.wipo.int/madrid/en/fees/ind_taxes.html>"
html_doc = urllib.request.urlopen(url).read()

# BeautifulSoupでHTMLを解析
soup = BeautifulSoup(html_doc, "html.parser")

#-------------------------------------------------------------
# id='table1'の<table>（新規出願）を抽出
table01 = soup.find('table', id='table1')

# 各行を解析する
rows = []
list = []
country = ''
for i, tr in enumerate(table01.find_all('tr')):
	rows = []
	for j, td in enumerate(tr.find_all('td')):
		#1列目の処理
		if j == 0:
			#空でなければ国名を取得
			if re.fullmatch('[a-zA-Z]+',td.text[0:1]):
				country = td.text.strip()
			rows.append(country)
		#2列目の処理
		elif j == 1:
			#料金を配列に
			rows.append(td.get_text(',').split(','))
		elif j == 2:
			rows.append(td.text)
	list.append(rows)
#-------------------------------------------------------------
# データの書き出し
f_path = 'individual_fee.csv'
with open(f_path, 'w', newline='', encoding="utf-8") as f:
	# ヘッダ
	writer = csv.writer(f)
	writer.writerow(['country','fee','type'])
	#データ
	writer.writerows(list)
