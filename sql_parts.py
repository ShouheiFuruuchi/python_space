import pyodbc

# データベースに接続する
driver = "SQL Server"
server ="FUN-PC119"
database = 'Competitor'#競合店実績DB
#database = 'TimeZoneSales', #時間帯売上実績DB
trusted_connection = "yes"
conn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';POST=1433;Trusted_Connection='+trusted_connection+';')


cursor = conn.cursor()
def SELECT():
  #cursor.execute('SELECT * FROM Competitor')
  cursor.execute('SELECT * FROM Competitor ')
  for row in cursor:
    print(row)

# データの登録
#競合店実績を登録
def INSERT_COMPETITOR(BRAND,SHOP_TYPE,YEAR,MONTH,DAY,DOW,DOW_TYPE,SALES,NOC):
  driver = "SQL Server"
  server ="FUN-PC119"
  database = 'Competitor'#競合店実績DB
  #database = 'TimeZoneSales', #時間帯売上実績DB
  trusted_connection = "yes"
  conn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';POST=1433;Trusted_Connection='+trusted_connection+';')


  cursor = conn.cursor()
  cursor.execute("INSERT INTO Competitor(BRAND,SHOP_TYPE,YEAR,MONTH,DAY,DOW,DOW_TYPE,SALES,NOC) VALUES ('" 
                 + str(BRAND) + "','" + str(SHOP_TYPE) + "','" + str(YEAR) + "','"+ str(MONTH) + "','" + str(DAY) + "','" + str(DOW) + "','" + str(DOW_TYPE) + "','" + str(SALES) + "','" + str(NOC) + "')")
  conn.commit()
  
  
# データの登録
#売上実績を登録
def INSERT_SALESDATA(
    SHOP_NAME,YEAR,MONTH,DAY,DOW,DOW_TYPE,BUGET,SALES,NOC,QUANTITY,P,CUP,SET_Ratio,
      OP_S,OP_Q,
      CD_S,CD_Q,
      JK_S,JK_Q,
      KT_S,KT_Q,
      CS_S,CS_Q,
      CT_S,CT_Q,
      BL_S,BL_Q,
      SK_S,SK_Q,
      PT_S,PT_Q,
      TR_S,TR_Q,
      INN_S,INN_Q,
      SETUP_S,SETUP_Q
      ,ACC_S,ACC_Q
      ,SH_S,SH_Q
      ,OTHERS_S,OTHERS_Q
      ):
  driver = "SQL Server"
  server ="FUN-PC119"
  database = 'Competitor'#競合店実績DB
  #database = 'TimeZoneSales', #時間帯売上実績DB
  trusted_connection = "yes"
  conn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';POST=1433;Trusted_Connection='+trusted_connection+';')
  cursor = conn.cursor()
  cursor.execute("INSERT INTO Competitor(BRAND,SHOP_TYPE,YEAR,MONTH,DAY,DOW,DOW_TYPE,SALES,NOC) VALUES ('" +
      str(SHOP_NAME) + "','" + str(YEAR) + "','" + str(MONTH) + "','" + str(DAY) + "','" + str(DOW) + "','" + 
      str(DOW_TYPE) + "','" + str(BUGET) + "','" + str(SALES) + "','" + str(NOC) + "','" + str(QUANTITY) + "','" + 
      str(P) + "','" + str(CUP) + "','" + str(SET_Ratio) + "','" + 
      str(OP_S) + "','" + str(OP_Q) + "','" + 
      str(CD_S) + "','" + str(CD_Q) + "','" + 
      str(JK_S) + "','" + str(JK_Q) + "','" + 
      str(KT_S) + "','" + str(KT_Q) + "','" + 
      str(CS_S) + "','" + str(CS_Q) + "','" + 
      str(CT_S) + "','" + str(CT_Q) + "','" + 
      str(BL_S) + "','" + str(BL_Q) + "','" + 
      str(SK_S) + "','" + str(SK_Q) + "','" + 
      str(PT_S) + "','" + str(PT_Q) + "','" + 
      str(TR_S) + "','" + str(TR_Q) + "','" + 
      str(INN_S) + "','" + str(INN_Q) + "','" + 
      str(SETUP_S) + "','" + str(SETUP_Q) + "','" + 
      str(ACC_S) + "','" + str(ACC_Q) + "','" + 
      str(SH_S) + "','" + str(SH_Q) + "','" + 
      str(OTHERS_S) + "','" + str(OTHERS_Q)
            
                 + "')"
                 )
  conn.commit()
    
  
  

# データの更新
def UPDATE():
  cursor.execute("UPDATE table_name SET column1 = ? WHERE id = ?", new_value, id_value)
  conn.commit()

# データの削除
def DELERT():
  cursor.execute("DELETE FROM table_name WHERE id = ?", id_value)
  conn.commit()
  
SELECT()

# 接続を閉じる
conn.close()