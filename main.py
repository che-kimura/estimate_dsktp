#作成日：2023年10月16日
#外国商標の見積りを作成する
import PySimpleGUI as sg
import csv
import pprint
# import os
# import jaconv
# import re
import pandas as pd
# import numpy as np
# import numexpr
# import math
# from collections import Counter
# from openpyxl import Workbook
# import win32com.client
# import pyperclip
# from pathlib import Path


DATAFILE = 'gs_ruijikijun_data.csv'

#
def main():
    #最初の入力フォームを表示する
    display_input_form()

def display_input_form():
    # デザインテーマの設定
    sg.theme('LightBlue3')

    discount = {
        0:True,
        5:False,
        10:False,
        15:False,
        20:False,
        25:False,
        30:False,
    }
    elements = {
        '宛先':'',
        '商標':'',
        '区分':'',
        '見積番号':'',
        '割引':'',
    }
    #区分
    classes = []
    # ウィンドウの部品とレイアウト
    layout = [
        [sg.Text('基本情報を入力してください。')],
        [sg.Text('宛先：',size=(10, 1), pad=((5, 5), (10, 10))),sg.InputText(key='atesaki', pad=((5, 5), (10, 10)))],
        [sg.Text('商標：',size=(10, 1), pad=((5, 5), (10, 10))),sg.InputText(key='mark', pad=((5, 5), (10, 10)))],
        [sg.Text('区分：', pad=((5, 5), (10, 10)))],
        [sg.Checkbox('第'+str(num)+'類',key='kubun'+str(num), pad=((5, 10), (0, 0))) for num in range(1,11)],
        [sg.Checkbox('第'+str(num)+'類',key='kubun'+str(num), pad=((5, 3), (0, 0))) for num in range(11,21)],
        [sg.Checkbox('第'+str(num)+'類',key='kubun'+str(num), pad=((5, 3), (0, 0))) for num in range(21,31)],
        [sg.Checkbox('第'+str(num)+'類',key='kubun'+str(num), pad=((5, 3), (0, 0))) for num in range(31,41)],
        [sg.Checkbox('第'+str(num)+'類',key='kubun'+str(num), pad=((5, 3), (0, 0))) for num in range(41,46)],
        [sg.Text('見積番号：',size=(10, 1), pad=((5, 5), (30, 10))),sg.InputText(key='mitsumri_no', pad=((5, 5), (30, 10)))],
        [sg.Text('割引：',size=(10, 1), pad=((5, 5), (10, 0)))],
        [sg.Radio(str(k)+'%',key=k, group_id='discount', default=v, pad=((5, 5), (0, 10))) for k, v in discount.items()],
        [sg.Button('マドプロ', key='mp', pad=((5, 5), (20, 5))), sg.Button('個別出願', key='kobetsu', pad=((5, 5), (20, 5))), sg.Button('調査', key='search', pad=((5, 5), (20, 5))), sg.Button('クリア', key='clear', pad=((5, 5), (20, 5)))],
    ]
    # ウィンドウの生成
    window = sg.Window('外国商標見積作成', layout, size=(800,500))
    # イベントループ
    while True:
        event, values = window.read()
        #ウィンドウのXボタンを押したときの処理
        if event == sg.WIN_CLOSED:
            break
        #「マドプロ」ボタンが押されたときの処理
        elif event == 'mp':
            #入力された値を取得する
            elements['宛先'] = values['atesaki']
            elements['見積番号'] = values['mitsumri_no']
            elements['商標'] = values['mark']
            #区分
            for num in range(1,46):
             if values['kubun'+str(num)] == True:
                 classes.append(num)
            elements['区分'] = classes
            #割引
            for key in discount:
                if values[key] == True:
                    elements['割引'] = key
            #print(elements)
            #マドプロ入力フォームを表示する
            display_mp_form(elements)
    window.close()

def display_mp_form(elements):
    # デザインテーマの設定
    sg.theme('LightBlue3')
    # #通常／事後指定
    # radio_type = {
    #     'type1':'新規出願',
    #     'type2':'事後指定',
    # }
    # #色彩
    # radio_color = {
    #     'color1':'白黒',
    #     'color2':'色彩あり',
    # }
    # #優先権
    # radio_priority = {
    #     'priority1':'優先権あり',
    #     'priority2':'優先権なし',
    # }
    #基礎出願の見積り
    estimate_base = {
        'しない':True,
        'する（調査無し）':False, 
        'する（通常調査）':False, 
        'する（簡易調査）':False,
    }
    #指定国のチェックボックス
    #アジア
    desig_countries_asia = []
    #北米
    desig_countries_na = []
    #中南米
    desig_countries_sca = []
    #オセアニア
    desig_countries_oce = []
    #欧州
    desig_countries_eu = []
    #中東
    desig_countries_me = []
    #アフリカ
    desig_countries_af = []
    
    
    #country_code.csvを読み込んで配列に振り分ける
    f_path = 'country_code.csv'
    data = pd.read_csv(f_path,encoding="UTF-8",dtype={'code':'object','name':'object','area':'int'})
    df_asia = data.query('area == 1')
    df_na = data.query('area == 2')
    df_sca = data.query('area == 3')
    df_oce = data.query('area == 4')
    df_eu = data.query('area == 5')
    df_me = data.query('area == 6')
    df_af = data.query('area == 7')
    
    #通常／事後指定
    l_syutsugan = [
        #[sg.Radio(item[1], key=item[0], group_id='type') for item in radio_type.items()]
        [sg.Radio('通常出願', key='type1', group_id='type', default=True),sg.Radio('事後指定', key='type2', group_id='type', default=False)]
    ]
    #色彩の有無
    l_color = [
        [sg.Radio('色彩なし', key='color1', group_id='color', default=True),sg.Radio('色彩あり', key='color2', group_id='color', default=False)],
        #[sg.Radio(item[1], key=item[0], group_id='color') for item in radio_color.items()]
    ]
    #優先権
    l_priority = [
        [sg.Radio('優先権なし', key='priority1', group_id='priority', default=True),sg.Radio('優先権あり', key='priority2', group_id='priority', default=False)],
        #[sg.Radio(item[1], key=item[0], group_id='priority') for item in radio_priority.items()]
    ]
    #基礎出願の見積り
    l_kisomitsumori = [
        [sg.Radio(item[0], key=item[0], group_id='estimate_base', default=item[1]) for item in estimate_base.items()]
    ]
    #指定国
    #アジア
    col_asia = [
        [sg.Checkbox(code+' - '+name,key=code)] for code, name in zip(df_asia['code'],df_asia['name'])
    ]
    #欧州
    col_europe = [
        [sg.Checkbox(code+' - '+name,key=code)] for code, name in zip(df_eu['code'],df_eu['name'])
    ]
    #北アメリカ
    col_na = [
        [sg.Checkbox(code+' - '+name,key=code)] for code, name in zip(df_na['code'],df_na['name'])
    ]
    #中南米
    col_sca = [
        [sg.Checkbox(code+' - '+name,key=code)] for code, name in zip(df_sca['code'],df_sca['name'])
    ]
    #オセアニア
    col_oce = [
        [sg.Checkbox(code+' - '+name,key=code)] for code, name in zip(df_oce['code'],df_oce['name'])
    ]
    #中東
    col_me = [
        [sg.Checkbox(code+' - '+name,key=code)] for code, name in zip(df_me['code'],df_me['name'])
    ]
    #アフリカ
    col_af = [
        [sg.Checkbox(code+' - '+name,key=code)] for code, name in zip(df_af['code'],df_af['name'])
    ]
    #フレーム
    frame_asia = sg.Frame('アジア',[
        [sg.Column(col_asia,size=(300,620),scrollable=True)]
    ])
    frame_europe = sg.Frame('欧州',[
        [sg.Column(col_europe,size=(300,620),scrollable=True)]
    ])
    frame_america = sg.Frame('北中南米・オセアニア',[
        [sg.Column(col_na,size=(300,90),scrollable=True)],
        [sg.Column(col_sca,size=(300,370),scrollable=True)],
        [sg.Column(col_oce,size=(300,120),scrollable=True)]
    ], vertical_alignment='TOP')
    frame_africa = sg.Frame('中東・アフリカ',[
        [sg.Column(col_me,size=(300,220),scrollable=True)],
        [sg.Column(col_af,size=(300,380),scrollable=True)]
    ])
    frame_desigcountry = sg.Frame('指定国',[
        [frame_asia,frame_europe,frame_america,frame_africa]
    ])
    
    # ウィンドウの部品とレイアウト
    layout = [
        [sg.Frame('通常／事後指定',l_syutsugan),sg.Frame('色彩の有無',l_color),sg.Frame('優先権',l_priority),sg.Frame('基礎出願の見積り',l_kisomitsumori)],
        [frame_desigcountry],
        [sg.Button('見積作成', key='make_mp', pad=((5, 5), (20, 5))), sg.Button('クリア', key='clear', pad=((5, 5), (20, 5))),sg.Button('指定国の全てチェックする', key='all_ck', pad=((5, 5), (20, 5))), sg.Button('指定国のチェックをすべて外す', key='del_ck', pad=((5, 5), (20, 5)))],
    ]
    # ウィンドウの生成
    window = sg.Window('マドプロ見積作成', layout, size=(1400,900), resizable=True)
    # イベントループ
    while True:
        event, values = window.read()
        #ウィンドウのXボタンを押したときの処理
        if event == sg.WIN_CLOSED:
            break
        #「マドプロ」ボタンが押されたときの処理
        #elif event == 'make_mp':
            #マドプロ見積作成関数を呼び出す

        #「全てチェックする」ボタンが押されたときの処理
        #elif event == 'all_ck':
            #全ての国をチェックする
        #「チェックをすべて外す」ボタンが押されたときの処理
        #elif event == 'del_ck':
            #全ての国のチェックを外す    
        #「クリア」ボタンが押されたときの処理
        elif event == 'clear':
            window['input'].Update('')
            window['output'].Update('')

    window.close()

""" def make_mp_estimate():
    #テンプレートブック（temp_mp.xlsx）を開く
    abspath = str(Path(r"temp_mp.xlsx").resolve())
    wb = xl.Workbooks.Add(abspath)
    #シート「指定商品役務」を選択
    ws = wb.Worksheets('指定商品役務')
    #指定商品役務の値をA3セルに入力
    ws.Range("A3").Value = list[1]
    #シート「指定商品役務一覧」を選択
    ws = wb.Worksheets('指定商品役務一覧')
    #A3からデータを入力
    
def get_unique_list(seq):
    seen = []
    return [x for x in seq if x not in seen and not seen.append(x)]
def Export_to_Excel(text):
    xlEdgeLeft         =  7
    xlEdgeBottom       =  9
    xlEdgeRight        = 10
    xlInsideVertical   = 11
    xlHairline = 1
    xlThin     = 2
    xlContinuous    =  1
    xlAutomatic   = -4105
    xlCenter      = -4108
    xlLeft        = -4131
    index = 0
    #文字列を改行で配列にする
    list = []
    
    #最後に改行コードがあれば削除
    if text[-1] == '\n' or text[-1] == '\r\n':
        list = text.split('\n')[:-1]        
    else:
        list = text.split('\n')
    xl = win32com.client.Dispatch("Excel.Application")
    #動いている様子を見てみる
    xl.Visible = True
    #テンプレートブック（template.xlsx）を開く
    abspath = str(Path(r"template.xlsx").resolve())
    wb = xl.Workbooks.Add(abspath)
    #シート「指定商品役務」を選択
    ws = wb.Worksheets('指定商品役務')
    #指定商品役務の値をA3セルに入力
    ws.Range("A3").Value = list[1]
    #シート「指定商品役務一覧」を選択
    ws = wb.Worksheets('指定商品役務一覧')
    #A3からデータを入力
    for i, n in enumerate(list):
        if i <= 3:
            continue
        if n == '［類似群カウント］':
            ws.Range("A" + str(i-2) + ":E" + str(i-2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range("A" + str(i-2) + ":E" + str(i-2)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            ws.Range("A" + str(i-2) + ":E" + str(i-2)).Borders(xlEdgeBottom).TintAndShade = 0
            ws.Range("A" + str(i-2) + ":E" + str(i-2)).Borders(xlEdgeBottom).Weight = xlThin
            ws.Range("A3:E" + str(i-2)).WrapText = True
            ws.Range("A3:E" + str(i-2)).HorizontalAlignment = xlCenter
            ws.Range("B3:C" + str(i-2)).HorizontalAlignment = xlLeft
            index = i
            break
        #タブで区切る
        element = []
        element = n.split('\t')
        ws.Range("A"+ str(i-1)).Value = element[0]
        if element[0] == '':
            ws.Range("B"+ str(i-1)).Value = element[1]
            ws.Range("C"+ str(i-1)).Value = element[2]
        else:
            ws.Range("B"+ str(i-1)).Value = element[1]
            ws.Range("C"+ str(i-1)).Value = element[2]
            ws.Range("D"+ str(i-1)).Value = element[3]
            if len(element) == 5:
                ws.Range("E"+ str(i-1)).Value = element[4]
        #罫線を引く
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeLeft).TintAndShade = 0
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeLeft).Weight = xlThin
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeRight).LineStyle = xlContinuous
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeRight).TintAndShade = 0
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeRight).Weight = xlThin
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlInsideVertical).LineStyle = xlContinuous
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlInsideVertical).TintAndShade = 0
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlInsideVertical).Weight = xlThin
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeBottom).TintAndShade = 0
        ws.Range("A" + str(i-1) + ":E" + str(i-1)).Borders(xlEdgeBottom).Weight = xlHairline
    #シート「類似群カウント」を選択
    ws = wb.Worksheets('類似群カウント')
    #A3から区分、個数を入力
    for i, m in enumerate(list[index+2:]):
        element = []
        element = m.split('\t')
        ws.Range("A" + str(i+3)).Value = element[0]
        ws.Range("B" + str(i+3)).Value = element[1]
        #罫線を引く
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeLeft).TintAndShade = 0
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeLeft).Weight = xlThin
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeRight).LineStyle = xlContinuous
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeRight).TintAndShade = 0
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeRight).Weight = xlThin
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlInsideVertical).LineStyle = xlContinuous
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlInsideVertical).TintAndShade = 0
        ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlInsideVertical).Weight = xlThin
        if i == len(list)-index-3:
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).TintAndShade = 0
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).Weight = xlThin
            ws.Range("A3:B" + str(i+3)).HorizontalAlignment = xlCenter
        else:
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).TintAndShade = 0
            ws.Range("A" + str(i+3) + ":B" + str(i+3)).Borders(xlEdgeBottom).Weight = xlHairline


# ステップ2. デザインテーマの設定
sg.theme('LightBlue3')

# ステップ3. ウィンドウの部品とレイアウト
layout = [
    [sg.Text('指定商品・役務を入力してください。')],
    [sg.Text('※区切り文字は日本語は「，（全角コンマ）」「改行」、英語は「; （半角セミコロン＋スペース）」「改行」です。')],
    [sg.Multiline(default_text='', size=(100, 10), key='input')],
    [sg.Button('日→英', key='JtoE'), sg.Button('英→日', key='EtoJ'), sg.Button('クリア', key='clear'), sg.Button('Excelで出力', key='export')],
    #[sg.Output(size=(100,30), key='output')]
    [sg.Multiline(default_text='', size=(100, 30), key='output')],
]
# ステップ4. ウィンドウの生成
window = sg.Window('指定商品役務翻訳／一覧作成', layout)
# ステップ5. イベントループ
while True:
    event, values = window.read()
     #ウィンドウのXボタンを押したときの処理
    if event == sg.WIN_CLOSED:
        break
     #「日→英」ボタンが押されたときの処理
    elif event == 'JtoE':
        # window['output'].Update('')
        # out = Translate_JtoE(values["input"], DATAFILE)
        # print('［指定商品役務］')
        # print(out[0][0].upper() + out[0][1:] + '.')
        # print('［指定商品役務一覧］')
        # print('区分\t英語\t日本語\t類似群\tニースコード')
        # for i in sorted(out[1], key=lambda x: (x[0], x[3])):
        #     print(i) 
        # print('［類似群カウント］')
        # print('区分' + '\t' + '個数')
        # c = Counter([x[0] for x in get_unique_list(out[2])])
        # #区分でソートする
        # c2 = sorted(c.items())
        # #print(c2)
        # for key in c2:
        #     print(key[0] + '\t' + str(key[1]))


        window['output'].Update('')
        out = Translate_JtoE(values["input"], DATAFILE)
        text_out = '［指定商品役務］\n'
        text_out = text_out + out[0][0].upper() + out[0][1:] + '.'
        text_out = text_out + '\n［指定商品役務一覧］\n'
        text_out = text_out + '区分\t英語\t日本語\t類似群\tニースコード\n'
        for i in sorted(out[1], key=lambda x: (x[0], x[3])):
            text_out = text_out + i +'\n'
        text_out = text_out + '［類似群カウント］\n'
        text_out = text_out + '区分' + '\t' + '個数\n'
        c = Counter([x[0] for x in get_unique_list(out[2])])
        #区分でソートする
        c2 = sorted(c.items())
        for key in c2:
            text_out = text_out + key[0] + '\t' + str(key[1]) + '\n'
        window['output'].Update(text_out)
     #「英→日」ボタンが押されたときの処理
    elif event == 'EtoJ':
        window['output'].Update('')
        out = Translate_EtoJ(values["input"], DATAFILE)
        text_out = '［指定商品役務］\n'
        text_out = text_out + out[0] + '\n'
        text_out = text_out + '［指定商品役務一覧］\n'
        text_out = text_out + '区分\t英語\t日本語\t類似群\tニースコード\n'
        for i in sorted(out[1], key=lambda x: (x[0], x[3])):
            text_out = text_out + i + '\n'
        text_out = text_out + '［類似群カウント］\n'
        text_out = text_out + '区分' + '\t' + '個数\n'
        c = Counter([x[0] for x in get_unique_list(out[2])])
        #区分でソートする
        c2 = sorted(c.items())
        for key in c2:
            text_out = text_out + key[0] + '\t' + str(key[1]) + '\n'
        window['output'].Update(text_out)
    #「クリア」ボタンが押されたときの処理
    elif event == 'clear':
        window['input'].Update('')
        window['output'].Update('')
    #「エクスポート」が押されたときの処理
    elif event == 'export':
        Export_to_Excel(values["output"]) """

if __name__ == "__main__":
    main()