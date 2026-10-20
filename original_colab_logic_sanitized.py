# -*- coding: utf-8 -*-
# HISTORY-ONLY / NOT GUARANTEED RUNNABLE
# This is a minimally redacted copy of the user's original logic (up through
# the data-extraction portion), excluding the original Colab URL, mounted
# Drive path, and research-specific LaTeX template. It is included to show
# the development approach, not as a verified, reusable program.
# Actual research files are NOT included. For reproducible demo use ../main.py.

#把檔名弄出來
import os
folder_path = "data"  # 研究原始資料未公開；如要執行，需自行準備 .xls 檔

excel_files = [
    f for f in os.listdir(folder_path)
    if f.lower().endswith((".xls", ".xlsx"))
]

excel_files
#h此時應有n個result裝有xls.

import pandas as pd

result = {}

for file in excel_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path)

    result[file] = df

"""**邏輯 順序 不然想出來亂掉**

* *建立每個關鍵字的字典*

* *把關鍵字加上編號作為鍵*

* *然後對應的欄數做值*

* *取出第幾次出現的*


把尋找的關鍵字與當中的次序設為參數 然後一個函數之後分層使用(?)
"""

#n百分比,n次數,
# 原始研究用變項代碼已隱去
#keyword_col=[n百分比,n次數,n顯著性,nexp,n上界,n下界]

"""建立字典

在此發現若用順序有問題，必須用表格的樣子去篩選


[多條件設定表格模樣]


```
match a:
    case 1:
        print("apple")
    case 2:
        print("banana")
    case _:
        print("other")
```


我要以關鍵字辨別模樣的要有幾個值，上下的邊界我要取這個數據與否，
不過用一堆if實在太不方便

**確定邏輯**

每提取到一次關鍵字要做到三件事:

1.   把他們上下的資料都抓出來(這個要import得靠ai我不熟)

2.   把他的上下邊界格數讀出來建成dict

3.   用名子+出現次數作為鍵值


```
(一套)
nkey={"key1:["上界，下界"]"}
keys = [k for k, v in nkey.items() if v == [上界,下界]]

```
上=上，下=下，才取出來

然後往下讀幾格的數據
"""

from re import U
#把當中的格子抓出來
import xlrd


def tracker(filename,key):
  global book,sheet
  book = xlrd.open_workbook(os.path.join(folder_path, filename))
  sheet = book.sheet_by_index(0)

  keyword = key
  matches = []

  for r in range(sheet.nrows):
      for c in range(sheet.ncols):
          value = sheet.cell_value(r, c)
          if str(value) == keyword:
              matches.append((r, c))

  return matches

n百分比={}
n次數={}
n顯著性={}
nexp={}
n上界={}
n下界={}
關鍵字=["百分比","次數","顯著性。","Exp(B)","上界","下界"]
關鍵字集合=[n百分比,n次數,n顯著性,nexp,n上界,n下界]
temp={}
exe=excel_files
def search(sfilename, skey, temp):
    temp={}
    data = builder(sfilename, skey)
    rev = {v: k for k, v in data.items()}  # 只建一次
    if skey not in ["顯著性。","Exp(B)"]:
      for i in data.values():
          temp[f"{sfilename}{rev[i]}"]=(detective(int(i[0]),int(i[1]),"否"))

    else:
     # print("Yes")
      for i in data.values():
          temp[f"{sfilename}{rev[i]}"]=(detective(int(i[0]),int(i[1]),"是"))
    return temp

#檢查重複項目
def check(tr,ty):
  com1=[]
  com2=[]
  only=[]
  if ty=="字典":
    for a in tr:
      com1.append(a[0])
    ccom=Counter(com1)
 #   print(ccom)
    for k in ccom.keys():
      if ccom[k]>1:
        com2.append(k)
    idxs = [i for i, v in enumerate(tr) if v[0] not in com2]
    for d in idxs:
      only.append(tr[d])
  elif ty=="tuple" :
    for a in tr:
      com1.append(a)
      continue
    ccom=Counter(com1)
    only.extend(ccom.keys())
  else:
    for a in tr:
      com1.append(a[0])
      com1.append(a[3])
      continue
    ccom=Counter(com1)
    only.extend(ccom.keys())
  #  print(only)
  return only

#建立序號與格子的dict
from collections import Counter

def builder(bfilename,bkey):
  tr =tracker(bfilename,bkey)
  if bkey!="顯著性。":
      order_col_row = ocr = { f"{bkey}{tr.index(v)+1}" : v for  v in tr}
      return ocr
  else:
    #建立一個比對函數
    ocr2 = { f"{bkey}{tr.index(v)+1}" : v for  v in check(tr,"字典",)}
    return ocr2

#找到上下的資造

#有點繞但是可用的土法邏輯
def detective(row,col,tell):
    lower=[]
    f=[]
    for i in range(1,8):
      lval=sheet.cell_value(row+i,col)
      lower.append(lval)#加上判斷邏輯 後刪除
 #   print(lower,1)
    if tell=="是":
      lower2=lower[1:7]
 #     print(lower2,2)
      if lower2[0:6]=="" or (all(item == "" for item in lower2[1:6]) and "" != lower2[0]) or (all(item == "" for item in lower2[2:6]) and "" not in lower2[0:2] )or (all(item == "" for item in lower2[3:6]) and "" not in lower2[0:3] )or (all(item == "" for item in lower2[4:6]) and "" not in lower2[0:4] ):
        idx = [i for i, v in enumerate(lower2) if v !=""]
        for k in idx :
          f.append(lower2[k])
 #         print(f,3)
      elif all(item == "" for item in lower2[2:6]):
        f=lower2[0:2]
      elif all(item == "" for item in lower2[3:6]):
        f=lower2[0:3]
      elif all(item == "" for item in lower2[4:6]):
        f=lower2[0:4]
      elif all(item == "" for item in lower2[5:6]):
        f=lower2[0:5]

      else :
        f=lower2
 #     print(f,4)

    else:
      if lower[0:7]=="" or (all(item == "" for item in lower[1:7]) and "" != lower[0]) or (all(item == "" for item in lower[2:7]) and "" not in lower[0:2] )or (all(item == "" for item in lower[3:7]) and "" not in lower[0:3]) or (all(item == "" for item in lower[4:7]) and "" not in lower[0:4]) or (all(item == "" for item in lower[5:7]) and "" not in lower[0:5]) :
        idx = [i for i, v in enumerate(lower) if v !=""]
        for k in idx :
          f.append(lower[k])
          continue
      elif all(item == "" for item in lower[2:7]):
        f=lower[0:2]
      elif all(item == "" for item in lower[3:7]):
        f=lower[0:3]
      elif all(item == "" for item in lower[4:7]):
        f=lower[0:4]
      elif all(item == "" for item in lower[5:7]):
        f=lower[0:5]
      elif all(item == "" for item in lower[6:7]):
        f=lower[0:6]
      else :
        f=lower
#      print(f,6)
    upper=[]
    uval=sheet.cell_value(row-1,col)
    if uval!="":
      upper.append(uval)
#      print(upper)
    return (upper,f)

#函數代換 減少冗長

def selection(run,key,seat,up,down):
  templ,temp2,temph2=[],{},search(run,key,seat)
  templ.extend(temph2.keys())
  for i in templ:
    if temph2[i] !=None :
        if len(temph2[i][0]) == up and len(temph2[i][1]) ==down:
          temp2[i]=temph2[i][1]
          continue

  return temp2

"""邏輯:

情況一:第六格有東西但前面沒有[把空格取代成None]
情況二:第六格沒有但左或右有[把空格取代成None]
情況三:第六格沒有但左或右也沒有[標記注意回傳None]
"""

#分流紀錄符合表格樣式的
def devide(run,key,seat,s):
  match key,s:
    case "百分比",0:
      return selection(run,key,seat,0,6)

    case "次數",0:
      return selection(run,key,seat,0,6)
    case "顯著性。",0:
      return selection(run,key,seat,0,2)  #顯著性很麻煩，多寫一個篩選函數吧
    case "Exp(B)",0:
      return selection(run,key,seat,0,2)
    case "上界",0:
      return selection(run,key,seat,0,2)
    case "下界",0:
      return selection(run,key,seat,1,2)
    case "顯著性。",1:
      return selection(run,key,seat,0,6)  #好醜好醜好醜嗚嗚嗚嗚嗚
    case "Exp(B)",1:
      return selection(run,key,seat,0,6)
    case "上界",1:
      return selection(run,key,seat,0,6)
    case "下界",1:
      return selection(run,key,seat,1,6)

allh,allexp,allup,alldown,alle,alltimes,salle,sallup,salldown,sallexp={},{},{},{},{},{},{},{},{},{}
for i in exe :
    allh[i]=devide(i,"百分比",n百分比,0)
    alltimes[i]=devide(i,"次數",n次數,0)
    alle[i]=devide(i,"顯著性。",n顯著性,0)
    allexp[i]=devide(i,"Exp(B)",nexp,0)
    allup[i]=devide(i,"上界",n上界,0)
    alldown[i]=devide(i,"下界",n下界,0)
    salle[i]=devide(i,"顯著性。",n顯著性,1)
    sallexp[i]=devide(i,"Exp(B)",nexp,1)
    sallup[i]=devide(i,"上界",n上界,1)
    salldown[i]=devide(i,"下界",n下界,1)
    #很丑旦方便檢修:)

for i in [allh,allexp,allup,alldown,alle,alltimes,salle,sallup,salldown,sallexp]:
  print(i)

#有漏一些東西 建立一個新的special 我需要知道後四個數據

#四捨五入測試
from decimal import Decimal, ROUND_HALF_UP
def rounder(value, places):
    if value != "":
      # 1. 先將輸入轉為字串以確保精確
      d = Decimal(str(value))

      # 2. 建立精確度字串，例如 2 位就是 '0.01'
      precision = Decimal(str(places))

      # 3. 執行四捨五入
      return d.quantize(precision, rounding=ROUND_HALF_UP)
    else:
      return ""
print(rounder(7.3453,0.01))

"""如果一開始就都這樣搞，整體城市運行效率下降"""

#四捨五入先 轉字串
import math
from decimal import Decimal, ROUND_HALF_UP
#再依序拆包 驗證 嗚嗚嗚嗚
def package(pfilename,pall):
    #拆包
    spackage=pall[pfilename]
    #過濾重複 重新賦值
    temp3=[]
    temp4=[]
    for k in spackage.values():
    #  print(k)
      temp4.append(tuple(k))
    #print(temp4)    #這裡轉元祖

    return check(temp4,"tuple")

#拆包
alllist,allkeys,alldict=[allh,alltimes,allexp,sallexp,allup,sallup,alldown,salldown,alle,salle],["百分比","次數","Exp(B)2","Exp(B)6","上界2","上界6","下界2","下界6","顯著性2","顯著性6"],{}
for i in exe:
    for k,l in zip(alllist,allkeys):
      alldict[f"{i}{l}"]=package(i,k)

allcheck_keys=ack=[]
allcheck_values=acv=[]

ack.extend(alldict.keys())

for k in alldict.values():
#  print(k)
    acv.append(tuple(k))
print(alldict)
print(acv)
print(ack)

#切片注入分段
def lice(llist):
    llist[1:1]=["", ""]
    llist.extend(["", ""])

from IPython.core.oinspect import Inspector
#如果有長度不合的
import sys
def checklen(clongstr,clist):
  match clongstr:
    case x if  any(k in x for k in ["百分比", "Exp(B)6", "上界6", "下界6", "顯著性6", "次數"]):
      for i in range(0,6):
        if len(clist[i])!=6:
          print(f"{clongstr}/{i+1}數據數量!!!!!!!!!!!!!!!!Error!!!!!!!!!!!!!!檢查,有異常")
          sys.exit(f"{clist[i]}")
        else:
          return f"{clongstr}/{i+1}數據數正常"

    case x if any(k in x for k in ["Exp(B)2" , "上界2" , "下界2" , "顯著性2" ]):
      for i in range(0,2):
        if len(clist[i])!=2:
          print(f"{clongstr}/{i+1}數據數量!!!!!!!!!!!!!!!!Error!!!!!!!!!!!!!!檢查,有異常")
          sys.exit(f"{clist[i]}")
        else:
          return f"{clongstr}/{i+1}數據數正常"
#如果有缺漏值
def checknum(clongstr,clist):
    inchecktuple=ict=[]
    outchecktuple=oct={}
    bekeys=0
    checklist=[]
    if len(clist)!=2:
     # print(clist)
      for i in clist:
        bekeys+=1
       # print(i)
        ict=[]
        for k in i:
          #  print(k)
          if "顯著性6" in clongstr:
            ict.append(rounder(k,0.001))
          else:
            ict.append(rounder(k,0.01))
        #print(f"ict{ict}")
        oct[str(bekeys)]=tuple(ict)
        #print(f"oct{oct}")

      for t in oct.values():
        checklist.append(tuple(t))
      #  print(f"checklist{checklist}")
     # print(f"checklistd{check(checklist,"special")}")
      if len(check(checklist,"special"))==4:
        alldict[clongstr]=[]
        lice1=check(checklist,"special")[0:2]
        lice2=check(checklist,"special")[2:4]
        lice(lice1)
        lice(lice2)
        alldict[clongstr]=[tuple(lice1),tuple(lice2)]
        print(alldict[clongstr])
        return f"{clongstr}資料總量正常,四捨五入前有障礙"
      else:
        if "顯著性6" in clongstr:
           alldict[clongstr]=[clist[0],clist[-1]]
           print(alldict[clongstr])
           return f"{clongstr}資料總量正常,但是建議人工檢查此基因的第二個pvalue"

        else:
          print(f"{clongstr}/{i+1}數據數量!!!!!!!!!!!!!!!!Error!!!!!!!!!!!!!!檢查,有異常")
          sys.exit(f"{clist[i]}")
    else:
      return f"{clongstr}資料總量正常"

"""因此選擇這裡出現問題之後把東西丟出來四捨五入檢查"""

#驗證階段
for n,m in zip(ack,acv):
    print(checknum(n,m))
    print("--------------------------------------------------------------------")
    print(checklen(n,m))
    print("--------------------------------------------------------------------")



#輸出填表(對應變數)#
def simple(efilename,no,fi):
  fi[f"{efilename}notime1"]=rounder(alldict[f"{efilename}次數"][no][0],1)
  fi[f"{efilename}notime2"]=rounder(alldict[f"{efilename}次數"][no][1],1)
  fi[f"{efilename}noper1"]=rounder(alldict[f"{efilename}百分比"][no][0],0.1)
  fi[f"{efilename}noper2"]=rounder(alldict[f"{efilename}百分比"][no][1],0.1)
  fi[f"{efilename}upexp1"]=rounder(alldict[f"{efilename}Exp(B)2"][no][0],0.01)
  fi[f"{efilename}upexp2"]=rounder(alldict[f"{efilename}Exp(B)6"][no][0],0.01)
  fi[f"{efilename}upl1"]=rounder(alldict[f"{efilename}下界2"][no][0],0.01)
  fi[f"{efilename}upl2"]=rounder(alldict[f"{efilename}下界6"][no][0],0.01)
  fi[f"{efilename}upu1"]=rounder(alldict[f"{efilename}上界2"][no][0],0.01)
  fi[f"{efilename}upu2"]=rounder(alldict[f"{efilename}上界6"][no][0],0.01)
  fi[f"{efilename}upex1"]=rounder(alldict[f"{efilename}顯著性2"][no][0],0.001)
  fi[f"{efilename}upex2"]=rounder(alldict[f"{efilename}顯著性6"][no][0],0.001)
  fi[f"{efilename}yestime1"]=rounder(alldict[f"{efilename}次數"][no][3],1)
  fi[f"{efilename}yestime2"]=rounder(alldict[f"{efilename}次數"][no][4],1)
  fi[f"{efilename}yesper1"]=rounder(alldict[f"{efilename}百分比"][no][3],0.1)
  fi[f"{efilename}yesper2"]=rounder(alldict[f"{efilename}百分比"][no][4],0.1)
  fi[f"{efilename}downexp1"]=rounder(alldict[f"{efilename}Exp(B)2"][no][1],0.01)
  fi[f"{efilename}downexp2"]=rounder(alldict[f"{efilename}Exp(B)6"][no][3],0.01)
  fi[f"{efilename}downl1"]=rounder(alldict[f"{efilename}下界2"][no][1],0.01)
  fi[f"{efilename}downl2"]=rounder(alldict[f"{efilename}下界6"][no][3],0.01)
  fi[f"{efilename}downu1"]=rounder(alldict[f"{efilename}上界2"][no][1],0.01)
  fi[f"{efilename}downu2"]=rounder(alldict[f"{efilename}上界6"][no][3],0.01)
  fi[f"{efilename}downex1"]=rounder(alldict[f"{efilename}顯著性2"][no][1],0.001)
  fi[f"{efilename}downex2"]=rounder(alldict[f"{efilename}顯著性6"][no][3],0.001)

  return fi
def export(xfilename,xfi):
  #notime1,noper1,notime2,noper2,upexp1,upl1,upu1,upex1,upexp2,upl2,upu2,upex2,yestime1,yesper1,yestime2,yesper2,downexp1,downl1,downu1,downex1,downexp2,downl2,downu2,downex2
  nfi={}
  if "pfi"in xfi:
    nno=0
    return simple(xfilename,nno,nfi)
  else:
    nno=1
    return simple(xfilename,nno,nfi)


final_pfi={}
final_dfi={}

for ex in exe:
  final_pfi[ex]=export(ex,"pfi")
  final_dfi[ex]=export(ex,"dfi")

print(final_pfi)
