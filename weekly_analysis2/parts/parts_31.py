import openpyxl as xlpy
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time
from operator import itemgetter
import os
import shutil
import datetime
from datetime import timedelta,date
from webdriver_manager.chrome import ChromeDriverManager
#ーーーーーーーーーーーーーーーーーーーーー|　販売NETスクレイピング |ーーーーーーーーーーーーーーーーーーーーーーーーー

kasiwa = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[3]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[3]','柏','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[3]','柏.CSV',"01001008"]
chiba = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[4]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[4]', '千葉','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[4]','千葉.CSV',"01001009"]
isesaki = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[9]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[9]','伊勢崎','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[9]','伊勢崎.CSV',"01001028"]
nagamachi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[11]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[11]','長町','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[11]','長町.CSV',"01001032"]
hunabashi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[12]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[12]','船橋','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[12]','船橋.CSV',"01001033"]
hujimi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[13]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[13]','富士見','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[13]','富士見.CSV',"01001034"]
reiku = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[15]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[15]','レイク','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[15]','レイク.CSV',"01001036"]
ebina = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[17]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[17]','海老名','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[17]','海老名.CSV',"01001038"]
musashi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[18]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[18]','むさし','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[18]','むさし.CSV',"01001039"]
hiratuka = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[19]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[19]','平塚','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[19]','平塚.CSV',"01001040"]
natori = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[20]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[20]','名取','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[20]','名取.CSV',"01001041"]
otaka = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[21]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[21]','大高','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[21]','大高.CSV',"01001042"]
togocyo = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[22]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[22]','東郷町','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[22]','東郷町.CSV',"01001043"]
ota = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[23]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[23]','太田','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[23]','太田.CSV',"01001044"]
mito = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[24]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[24]','水戸','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[24]','水戸.CSV',"01001045"]
expo = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[25]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[25]','EXPO','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[25]','EXPO.CSV',"01001046"]
kawasaki = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[26]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[26]','川崎','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[26]','川崎.CSV',"01001047"]
sinmisato = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[27]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[27]','新三郷','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[27]','新三郷.CSV',"01001048"]
makuhari = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[28]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[28]','幕張','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[28]','幕張.CSV',"01001049"]
kagamihara = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[29]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[29]','各務原','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[29]','各務原.CSV',"01001050"]
sakai = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[30]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[30]','堺','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[30]','堺.CSV',"01001051"]

tenpo_list = [
  kasiwa,
  chiba,
  isesaki,
  # nagamachi,
  # hunabashi,
  hujimi,
  reiku,
  ebina,
  musashi,
  hiratuka,
  natori,
  otaka,
  togocyo,
  ota,
  mito,
  expo,
  kawasaki,
  sinmisato,
  makuhari,
  kagamihara,
  sakai,
  ]

tenpo = [
    ["01001008 FUN柏","柏"],
    ["01001009 FUN千葉C-one","千葉"],
    ["01001028 FUNスマーク伊勢崎","伊勢崎"],
    # ["01001032 FUNララガーデン長町","長町"],
    # ["01001033 FUNららぽーとTOKYO-BAY","船橋"],
    ["01001034 FUNららぽーと富士見","富士見"],
    ["01001036 FUNイオンレイクタウン","レイク"],
    ["01001038 FUNららぽーと海老名","海老名"],
    ["01001039 FUNイオンモールむさし村山","むさし"],
    ["01001040 FUNららぽーと湘南平塚","平塚"],
    ["01001041 FUNイオンモール名取","名取"],
    ["01001042 FUNイオンモール大高","大高"],
    ["01001043 FUNららぽーと愛知東郷","東郷町"],
    ["01001044 FUNイオンモール太田","太田"],
    ["01001045 FUNイオンモール水戸内原","水戸"],
    ["01001046 FUNららぽーとEXPOCITY","EXPO"],
    ["01001047 FUNラゾーナ川崎プラザ","川崎"],
    ["01001048 FUNららぽーと新三郷","新三郷"],
    ["01001049 FUNイオンモール幕張新都心","幕張"],
    ["01001050 FUNイオンモール各務原","各務原"],
    ["01001051 FUNららぽーと堺","堺"],
    
]

#-----------------------------------------------------------------------------------------------------------------------------
#　ここから　

inventory_folder = "C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory"


dr_files = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder'#今週実績
dr_files2 = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data'#過去実績
dr_files3 = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values'#売上集計
dr_files4 = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/buget'#今週日割予算


dr_read = os.listdir(dr_files)
dr_read2 = os.listdir(dr_files2)
dr_read3 = os.listdir(dr_files3)
dr_read4 = os.listdir(dr_files4)
dr_read5 = os.listdir(inventory_folder)

print(dr_read)
print(dr_files2)
print(dr_files3)
print("データをダウンロードしますか？\nYES ⇒ 0\nNo  ⇒ 1")

swicth_ = input()

if swicth_ == str(0):

  for file_name in dr_read:
    del_f_path = dr_files + '/' + file_name#削除ファイルパスの設定
    os.remove(del_f_path)#dataf内のファイルの削除
    
  for file_name2 in dr_read2:
    del_f_path2 = dr_files2 + '/' + file_name2#削除ファイルパスの設定
    os.remove(del_f_path2)#dataf内のファイルの削除  
    
  for file_name3 in dr_read3:
    del_f_path3 = dr_files3 + '/' + file_name3#削除ファイルパスの設定
    os.remove(del_f_path3)#dataf内のファイルの削除    
    
  for file_name4 in dr_read4:
    del_f_path4 = dr_files4 + '/' + file_name4#削除ファイルパスの設定
    os.remove(del_f_path4)#dataf内のファイルの削除    
    
  # for file_name5 in dr_read5:
  #   del_f_path5 = inventory_folder + '/' + file_name5#削除ファイルパスの設定
  #   os.remove(del_f_path5)#dataf内のファイルの削除   
    #--------------------------------------------------------

  todaytime = datetime.date.today()
  
  tod = '{0:20%y}'.format(todaytime)#今日の日付(西暦)

  print("期間のSTART日を入力して下さい！")
  print("例 20XX0101(20XX年1月1日の場合)")
  #period1 = tod + str(input())
  period1 = str(input())
  print("期間のEND日を入力して下さい！")
  print("例 20XX0107(20XX年1月7日の場合)")
  # = tod + str(input())
  period2 = str(input())
  
  
  y = period1[0:4]
  m = period1[4:6]
  d = period1[6:8]

  ymd = datetime.datetime.strptime(str(y) + '-' + str(m) + '-' + str(d), '%Y-%m-%d')#period1を日付に変換
  point_week = date(int(y),int(m),int(d)).isocalendar().week
  
  print("基準週",point_week)
  
  print(ymd)

  select_day_3 = ymd  + timedelta(days = 7)
  print(select_day_3)
  #day_of_week_3 = select_day_3.weekday()
  #print(day_of_week_3)
  
  print("遡る年数を指定して下さい")
  
  year_n =  int(input())
  #初期値の設定として年度の1月1WEEK目の日付を設定
  up_days = timedelta(days = 365 * year_n)


  #select_day = start_day + up_days

  select_day = ymd - up_days


  select_week_year = select_day.year
  select_week_month = select_day.month
  select_week_day = select_day.day

  select_week_no = datetime.date(select_week_year,select_week_month,select_week_day).isocalendar().week

  adjust = (point_week - select_week_no) * 7
  print(select_day)
  print("調整値",adjust)
  
  select_day = select_day + timedelta(days = adjust)
  
  print("調整日",select_day)
  #調整値
  print(select_week_no)
  
  day_of_week = select_day.weekday()
  print("曜日No",day_of_week)
  
  print(select_day.weekday())
  
  create_filename = "【" + period1 + "-" + period2 + "】週間分析.xlsx"
  

  #条件１曜日が月曜日である事
  #条件２周番号が１である事

  if day_of_week == 0 :
    print("OK")
    
    select_day_2 = select_day - timedelta(days = day_of_week)
    print("ポイント２",select_day_2)
    day_of_week_2 = select_day_2.weekday()
    print(day_of_week_2)
    
    select_week_year_2 = select_day_2.year
    select_week_month_2 = select_day_2.month
    select_week_day_2 = select_day_2.day
    
    period_1_previous = str(select_week_year_2) + str("{:0>2}".format(select_week_month_2)) + str("{:0>2}".format(select_week_day_2))#前年実績１
    
    period_2_1 = select_day_2 + timedelta(days = 6)#７⇒６に修正
    period_2_y = period_2_1.year
    period_2_m = period_2_1.month
    period_2_d = period_2_1.day
    
    period_2_previous = str(period_2_y) + "{:0>2}".format(period_2_m) + "{:0>2}".format(period_2_d)#前年実績2
    
    select_week_no_2 = datetime.date(select_week_year_2,select_week_month_2,select_week_day_2).isocalendar().week

    print(select_week_no)
  #---------------------------------------------------
  #前年翌週実績
  
    select_day_3 = select_day - timedelta(days = day_of_week ) + timedelta(days = 7 )

    print("ポイント",select_day_3)
    day_of_week_3 = select_day_3.weekday()
    print(day_of_week_3)
    
    select_week_year_3 = select_day_3.year
    select_week_month_3 = select_day_3.month
    select_week_day_3 = select_day_3.day
    
    period_1_previous3 = str(select_week_year_3) + str("{:0>2}".format(select_week_month_3)) + str("{:0>2}".format(select_week_day_3))#前年実績１
    
    period_3_1 = select_day_3 + timedelta(days = 6)#７⇒６に修正
    period_3_y = period_3_1.year
    period_3_m = period_3_1.month
    period_3_d = period_3_1.day
    
    period_2_previous3 = str(period_3_y) + "{:0>2}".format(period_3_m) + "{:0>2}".format(period_3_d)#前年実績2
    
    select_week_no_3 = datetime.date(select_week_year_3,select_week_month_3,select_week_day_3).isocalendar().week

    print(select_week_no)
    
    
  #前年来週実績
  
    select_day_4 = select_day - timedelta(days = day_of_week ) + timedelta(days = 14 )

    print("ポイント",select_day_4)
    day_of_week_4 = select_day_4.weekday()
    print(day_of_week_4)
    
    select_week_year_4 = select_day_4.year
    select_week_month_4 = select_day_4.month
    select_week_day_4 = select_day_4.day
    
    period_1_previous4 = str(select_week_year_4) + str("{:0>2}".format(select_week_month_4)) + str("{:0>2}".format(select_week_day_4))#前年実績１
    
    period_4_1 = select_day_4 + timedelta(days = 6)#７⇒６に修正
    period_4_y = period_4_1.year
    period_4_m = period_4_1.month
    period_4_d = period_4_1.day
    
    period_2_previous4 = str(period_4_y) + "{:0>2}".format(period_4_m) + "{:0>2}".format(period_4_d)#前年実績2
    
    select_week_no_4 = datetime.date(select_week_year_4,select_week_month_4,select_week_day_4).isocalendar().week

    print(select_week_no)  
    


  else :
    select_day_2 = select_day - timedelta(days = day_of_week)
    print("ポイント２",select_day_2)
    day_of_week_2 = select_day_2.weekday()
    print(day_of_week_2)
    
    select_week_year_2 = select_day_2.year
    select_week_month_2 = select_day_2.month
    select_week_day_2 = select_day_2.day
    
    period_1_previous = str(select_week_year_2) + str("{:0>2}".format(select_week_month_2)) + str("{:0>2}".format(select_week_day_2))#前年実績１
    
    period_2_1 = select_day_2 + timedelta(days = 6)#７⇒６に修正
    period_2_y = period_2_1.year
    period_2_m = period_2_1.month
    period_2_d = period_2_1.day
    
    period_2_previous = str(period_2_y) + "{:0>2}".format(period_2_m) + "{:0>2}".format(period_2_d)#前年実績2
    
    select_week_no_2 = datetime.date(select_week_year_2,select_week_month_2,select_week_day_2).isocalendar().week

    print(select_week_no)
  #---------------------------------------------------
  #前年翌週実績
  
    select_day_3 = select_day - timedelta(days = day_of_week ) + timedelta(days = 7 )

    print("ポイント",select_day_3)
    day_of_week_3 = select_day_3.weekday()
    print(day_of_week_3)
    
    select_week_year_3 = select_day_3.year
    select_week_month_3 = select_day_3.month
    select_week_day_3 = select_day_3.day
    
    period_1_previous3 = str(select_week_year_3) + str("{:0>2}".format(select_week_month_3)) + str("{:0>2}".format(select_week_day_3))#前年実績１
    
    period_3_1 = select_day_3 + timedelta(days = 6)#７⇒６に修正
    period_3_y = period_3_1.year
    period_3_m = period_3_1.month
    period_3_d = period_3_1.day
    
    period_2_previous3 = str(period_3_y) + "{:0>2}".format(period_3_m) + "{:0>2}".format(period_3_d)#前年実績2
    
    select_week_no_3 = datetime.date(select_week_year_3,select_week_month_3,select_week_day_3).isocalendar().week

    print(select_week_no)
    
    
  #前年来週実績
  
    select_day_4 = select_day - timedelta(days = day_of_week ) + timedelta(days = 14 )

    print("ポイント",select_day_4)
    day_of_week_4 = select_day_4.weekday()
    print(day_of_week_4)
    
    select_week_year_4 = select_day_4.year
    select_week_month_4 = select_day_4.month
    select_week_day_4 = select_day_4.day
    
    period_1_previous4 = str(select_week_year_4) + str("{:0>2}".format(select_week_month_4)) + str("{:0>2}".format(select_week_day_4))#前年実績１
    
    period_4_1 = select_day_4 + timedelta(days = 6)#７⇒６に修正
    period_4_y = period_4_1.year
    period_4_m = period_4_1.month
    period_4_d = period_4_1.day
    
    period_2_previous4 = str(period_4_y) + "{:0>2}".format(period_4_m) + "{:0>2}".format(period_4_d)#前年実績2
    
    select_week_no_4 = datetime.date(select_week_year_4,select_week_month_4,select_week_day_4).isocalendar().week

    print(select_week_no)  
    
    priod_list1 = str(period_1_previous) + " 〜 " + str(period_2_previous)
    priod_list2 = str(period_1_previous3) + " 〜 " + str(period_2_previous3)
    priod_list3 = str(period_1_previous4) + " 〜 " + str(period_2_previous4)
    
    priod_list1_key = [period_1_previous,period_2_previous]
    priod_list2_key = [period_1_previous3,period_2_previous3]
    priod_list3_key = [period_1_previous4,period_2_previous4]
    priod_list_key =  [period1,period2]
    
    
  priod_list = [
    priod_list1,
    priod_list2,
    priod_list3
  ]  
  
  priod_list_2 = [
    priod_list1_key,
    priod_list2_key,
    priod_list3_key,
    priod_list_key
  ]

  #---------------------------------------------------

  url = 'http://tri.hanbai-net.com/system/Login.aspx'
  #driver = webdriver.Chrome('C:/Users/fun-f/Downloads/chromedriver.exe')#旧
  #driver = webdriver.Chrome('C:/Users/fun-f/Desktop/myfile/chromedriver.exe')
  #driver = webdriver.Chrome("C:/Users/fun-f/chromedriver.exe")#2021 0724
  time.sleep(2)
  chrome_options = Options()
  chrome_options.add_experimental_option("detach", True)
  driver = webdriver.Chrome(ChromeDriverManager().install(),options=chrome_options)

  driver.get(url)         

  #id_1 = 'tenpo'
  #id_2 = 'tenpo'
  
  id_1 = 'trinityadmin'
  id_2 = 'AdminTrinity'

  

  loginid_1 = driver.find_element(By.ID, "ContentPlaceHolder1_txtUserCode")
  loginid_2 = driver.find_element(By.ID, "ContentPlaceHolder1_txtPassword")

  loginid_1.send_keys(id_1)#ユーザーIDを入力
  loginid_2.send_keys(id_2)#パスワードを入力

  #ログインボタンをクリック
  driver.find_element(By.ID, "ContentPlaceHolder1_btnLogin").click()
  
  time.sleep(1)

  driver.get('http://tri.hanbai-net.com/system/00000000.aspx')
  time.sleep(2)

  driver.get('http://tri.hanbai-net.com/system/30021901.aspx?id=010199')#品番別売上集計
  
  time.sleep(3)
  
  #--------------------------------------------------------------------------------------------------------
  #過去実績をダウンロード
  
  driver.find_element(By.ID, "ContentPlaceHolder1_txtCond02").clear()#日付クリア

  driver.find_element(By.ID, "ContentPlaceHolder1_txtCond02").send_keys(str(period_1_previous))#日付入力1

  driver.find_element(By.ID, "ContentPlaceHolder1_txtCond03").clear()#日付クリア

  driver.find_element(By.ID, "ContentPlaceHolder1_txtCond03").send_keys(str(period_2_previous))#日付入力2
  
  driver.find_element(By.ID, "ContentPlaceHolder1_btnCSV").click()#CSV出力

  time.sleep(3)#一時待機

  filelists = []
  for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
      base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
      if ext == '.csv':#拡張子csvが一致した場合…
          if base == '品番売上集計':
              filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
              #print("file:{},csv:{}" .format(file,csv))
              filelists.sort(key=itemgetter(0), reverse=True)#
              MAX_CNT = 0
              for i, file in enumerate(filelists):
                  if i > MAX_CNT-1:
                      print(file[0])
                      #file_1 = os.rename(i[0], 'kasi.csv')
                      os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" +'全店1.csv')
                      shutil.move("C:/Users/古内翔平/Downloads/" + '全店1.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data')                        
  time.sleep(1)                    
  
  
  
  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").send_keys(str(period_1_previous3))#日付入力1

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").send_keys(str(period_2_previous3))#日付入力2
  
  driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

  time.sleep(3)#一時待機

  filelists = []
  for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
      base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
      if ext == '.csv':#拡張子csvが一致した場合…
          if base == '品番売上集計':
              filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
              #print("file:{},csv:{}" .format(file,csv))
              filelists.sort(key=itemgetter(0), reverse=True)#
              MAX_CNT = 0
              for i, file in enumerate(filelists):
                  if i > MAX_CNT-1:
                      print(file[0])
                      #file_1 = os.rename(i[0], 'kasi.csv')
                      os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + '全店2.csv')
                      shutil.move("C:/Users/古内翔平/Downloads/" + '全店2.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data')                        
  time.sleep(1)             
  
  #前年来週実績
  
  
  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").send_keys(str(period_1_previous4))#日付入力1

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").send_keys(str(period_2_previous4))#日付入力2
  
  driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

  time.sleep(3)#一時待機

  filelists = []
  for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
      base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
      if ext == '.csv':#拡張子csvが一致した場合…
          if base == '品番売上集計':
              filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
              #print("file:{},csv:{}" .format(file,csv))
              filelists.sort(key=itemgetter(0), reverse=True)#
              MAX_CNT = 0
              for i, file in enumerate(filelists):
                  if i > MAX_CNT-1:
                      print(file[0])
                      #file_1 = os.rename(i[0], 'kasi.csv')
                      os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + '全店3.csv')
                      shutil.move("C:/Users/古内翔平/Downloads/" + '全店3.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data')                        
  time.sleep(1)             
  
  #--------------------------------------------------------------------------------------------------------
  


  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").send_keys(str(period1))#日付入力1

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").send_keys(str(period2))#日付入力2
  
  #----------全店------------

  #driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

  #time.sleep(5)

  driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

  time.sleep(3)#一時待機

  filelists = []
  for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
      base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
      if ext == '.csv':#拡張子csvが一致した場合…
          if base == '品番売上集計':
              filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
              #print("file:{},csv:{}" .format(file,csv))
              filelists.sort(key=itemgetter(0), reverse=True)#
              MAX_CNT = 0
              for i, file in enumerate(filelists):
                  if i > MAX_CNT-1:
                      print(file[0])
                      #file_1 = os.rename(i[0], 'kasi.csv')
                      os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + '全店.csv')
                      shutil.move("C:/Users/古内翔平/Downloads/" + '全店.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder')                        
  time.sleep(1)                    
  
  
  

  #--------店別---------
  #★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
  #品番別売上集計をダウンロード

  for i_1 in tenpo_list:
    target1 = driver.find_element(By.ID,"ContentPlaceHolder1_DropDownListCond04")#.click()#店舗名指定上段
    
    select_target1 = Select(target1)
    select_target1.select_by_value(str(i_1[5]))

    #driver.find_element(str(i_1[0])).click()#店舗選択


    target2 = driver.find_element(By.ID,"ContentPlaceHolder1_DropDownListCond05")#.click()#店舗名指定下段
    select_target2 = Select(target2)
    select_target2.select_by_value(str(i_1[5]))

    #driver.find_element(str(i_1[1])).click()#店舗選択


    #driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索

    driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

    time.sleep(3)#一時待機

    filelists = []
    for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '品番売上集計':
                filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + str(i_1[2]) + '.csv')
                        shutil.move("C:/Users/古内翔平/Downloads/" + str(i_1[2]) + '.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder') 
                      
  time.sleep(1)                    
  
  #--------------------------- 売上週計  -------------------------------  
  
  driver.get('http://tri.hanbai-net.com/system/30026401.aspx?id=010199')#売上集計＊
  
  file_no = 1
  for priod_n in priod_list_2:
    priod_select1 = priod_n[0]
    priod_select2 = priod_n[1]
    
    
    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond01").clear()
    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond01").send_keys(str(priod_select1))
    
    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").clear()
    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").send_keys(str(priod_select2))

    #driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()

    driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

    time.sleep(3)#一時待機
    
    

    filelists = []
    for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '売上集計':
                filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]),"C:/Users/古内翔平/Downloads/" +  '売上実績' + str(file_no) + '.csv')
                        shutil.move("C:/Users/古内翔平/Downloads/" + '売上実績' + str(file_no) + '.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values')
                        
                        file_no += 1
  
  driver.get('http://tri.hanbai-net.com/system/50010201.aspx?id=010199')#販売分析ログ
  time.sleep(2)
  
  #★★★全店実績をダウンロード
  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond07").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond07").send_keys(period1)#日付入力(前)

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond10").clear()#日付クリア

  driver.find_element(By.ID,"ContentPlaceHolder1_txtCond10").send_keys(period2)#日付入力(後)

  driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

  time.sleep(30)#一時待機

  filelists = []
  for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
      base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
      if ext == '.csv':#拡張子csvが一致した場合…
          if base == '販売分析ログ':
              filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
              #print("file:{},csv:{}" .format(file,csv))
              filelists.sort(key=itemgetter(0), reverse=True)#
              MAX_CNT = 0
              for i, file in enumerate(filelists):
                  if i > MAX_CNT-1:
                      print(file[0])
                      #file_1 = os.rename(i[0], 'kasi.csv')
                      os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + '全店顧客データ.csv')
                      shutil.move("C:/Users/古内翔平/Downloads/" + '全店顧客データ.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder')                        
  time.sleep(1)                                
  
  #★★★店別実績をダウンロード
  
  for i_name,i_shopname in zip(tenpo_list,tenpo):
    
    target3 = driver.find_element(By.ID,"ContentPlaceHolder1_DropDownList9")
    select_target3 = Select(target3)
    select_target3.select_by_value(str(i_name[5]))
    #.send_keys(str(i_name[0]))#店舗名を指定
    

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond07").clear()#日付クリア

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond07").send_keys(period1)#日付入力(前)

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond10").clear()#日付クリア

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond10").send_keys(period2)#日付入力(後)

    #driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索
  


    driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

    time.sleep(5)#一時待機
   
    filelists = []
    for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '販売分析ログ':
                filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + str(i_shopname[1]) + '顧客データ.csv')
                        shutil.move("C:/Users/古内翔平/Downloads/" + str(i_shopname[1]) + '顧客データ.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder')                        
    time.sleep(1)                                
  
  
   
  #★★★店別予算表をダウンロード
  
  driver.get("http://tri.hanbai-net.com/system/30020901.aspx?id=010199")
  day_element = ymd + timedelta(days = 7 )
  day_element2 = day_element + timedelta(days = 6 )
  
  start_day = str(day_element.year) + str(day_element.month).zfill(2) + str(day_element.day).zfill(2)
  end_day = str(day_element2.year) + str(day_element2.month).zfill(2) + str(day_element2.day).zfill(2)
 
  for i_name,shop_name in zip(tenpo_list,tenpo):
    
    target4 = driver.find_element(By.ID,"ContentPlaceHolder1_DropDownList1")#.send_keys(str(i_name[0]))#店舗名を指定
    select_target4 = Select(target4)
    select_target4.select_by_value(str(i_name[5]))
    
    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").clear()#日付クリア

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond02").send_keys(str(start_day))#日付入力(前)

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").clear()#日付クリア

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond03").send_keys(end_day)#日付入力(後)

    #driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCondRun"]').click()#検索
  


    driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

    time.sleep(5)
    
    filelists = []
    for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '日別予算設定管理':
                filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" +  str(shop_name[1]) + '日別予算設定管理.csv')
                        shutil.move("C:/Users/古内翔平/Downloads/" + str(shop_name[1]) + '日別予算設定管理.csv','C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/buget')                        
    time.sleep(1)   
    
    
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  #店舗在庫を取得  
    
    
  # for i_1 in tenpo_list:
  #   #店舗入力
  #   driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownListCond01"]').send_keys(i_1[4])
    
  #   #日付入力
  #   driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtCond04"]').clear()    
  #   driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtCond04"]').send_keys(period1)
    
  #   #CSV出力
  #   driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnCSV"]').click()
    
  #   time.sleep(5)#一時待機


  #   filelists = []
  #   for file in os.listdir("C:/Users/fun-f/Downloads"):#ディレクトリ内をfor文で取り出す
  #       base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
  #       if ext == '.csv':#拡張子csvが一致した場合…
  #           if base == '在庫一覧_':
  #               filelists.append([file, os.path.getctime(file)])#filelistsに取り出したfileにダウンロード時間を追加
  #               #print("file:{},csv:{}" .format(file,csv))
  #               filelists.sort(key=itemgetter(0), reverse=True)#
  #               MAX_CNT = 0
  #               for i, file in enumerate(filelists):
  #                   if i > MAX_CNT-1:
  #                       print(file[0])
  #                       #file_1 = os.rename(i[0], 'kasi.csv')
  #                       os.rename(file[0], str(i_1[2]) + '.csv')
  #                       shutil.move(str(i_1[2]) + '.csv',inventory_folder) 
     
    
  # print("SUCCESS!!")     
            
  driver.close()
  
  
  'C:/Users/fun-f/Desktop/analysis/buget'