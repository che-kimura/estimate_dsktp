import csv
import pprint
import pandas as pd

#country_code.csvを読み込んで配列に振り分ける
f_path = 'country_code.csv'
data = pd.read_csv(f_path,encoding="UTF-8")
df_asia = data.query('area == 1')
df_na = data.query('area == 2')
df_sca = data.query('area == 3')
df_oce = data.query('area == 4')
df_eu = data.query('area == 5')
df_me = data.query('area == 6')
df_af = data.query('area == 7')

print(df_asia)
print(df_na)
print(df_sca)
print(df_oce)
print(df_eu)
print(df_me)
print(df_af)
