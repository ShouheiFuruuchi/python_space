
#作成ファイルをcreate_file に移動するプログラムになります。


import os
import shutil

path = "C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis"
move_path = "C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/create_file"

files = os.listdir(path)


for file in files :
  if ".xlsx" in file and "【" in file:
    if "~$" in file :
      print("No_Move")
      
    else :
      print(file)
      shutil.move(os.path.join(path,file),"C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/create_file/")
    
