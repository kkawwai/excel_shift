import numpy as np
import pandas as pd
from random import random
import calendar
import datetime


"""
シフトを作成します。日、早、遅、休、の４種類をランダムに配置する。
demo.xlsxの一列目の二行目から行方向に従業員の名前を入力します。一行目の列方向のB～AFに1～31を入力し、AGに「休日数」と書きます。
各従業員の列に希望休を書いていきます。「休日数」の欄には、各従業員の休日数を記載してください。
日数は、自動で取得されます。2月に作成した場合、31日分のシフトが作成されます。10月に作成した場合は、30日分のシフトが作成されます。
"""

def read_excel():
    k = []
    count = 0
    df = pd.read_excel('demo.xlsx', sheet_name='Sheet1')
    for i in range(df.shape[0]):
        if pd.isna(df.iloc[i, 0]):
            k.append(i)
        else:
            count += 1
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month 
    _, days_in_month = calendar.monthrange(year, month + 1)
    df = pd.read_excel('demo.xlsx')
    df = df.iloc[0:count, 1:days_in_month+1]
    df = df.fillna('あ')
    df = df.replace('◎', 0).replace('あ', 2)
    d = pd.read_excel('demo.xlsx')
    holiday = d.iloc[0:count:,32:33].reset_index(drop=True)
    holiday.columns = ['休日数']
    origin = df.iloc[:,0:31].reset_index(drop=True)
    origin.columns = [i + 1 for i in range(len(origin.columns))]
    return origin, holiday


#2 休日数の数だけそれぞれの行の2の値を1にしたoriginをorigin_copyに移して渡す。
def first_gene(origin, holiday):
    days = len(origin.columns)

    origin_copy = origin.copy()


    for k in range(len(origin_copy)):
        h = []

        while len(h) < holiday.loc[k][0]:
            n = np.random.randint(1, days+1)
            if not n in h:
                h.append(n)
        

        for i in h:
            if origin_copy.loc[k][i] == 2:
                origin_copy.loc[k][i] = 1
    return origin_copy



#早出の勤務を追加する。3が最終的に早になって表示されます。
def day_shift(origin_copy):
    for k in range(len(origin_copy)):
        for i in range(len(origin_copy.columns)):
            if origin_copy.iloc[k][i+1] == 2:
                x = True if  0.6 > random() else False
                if x == True:
                    origin_copy.iloc[k][i+1] = 3
    return origin_copy  


#遅出の勤務を追加する。
def day_shifts(origin_copy):
    for k in range(len(origin_copy)):
        for i in range(len(origin_copy.columns)):
            if origin_copy.iloc[k][i+1] == 3:
                x = True if  0.5 > random() else False
                if x == True:
                    origin_copy.iloc[k][i+1] = 4
    return origin_copy     


def holiday_fix(origin_copy, holiday):
    for k_idx in range(len(origin_copy)):
        days = len(origin_copy.columns)
        k = np.count_nonzero(origin_copy[k_idx:k_idx+1] == 0)
        r = np.count_nonzero(origin_copy[k_idx:k_idx+1] == 1)
        e = k + r
        if e != holiday.loc[k_idx][0]:
            s = e - holiday.loc[k_idx][0]
            buf = 0
            if s > 0:
                while buf < s:
                    n = np.random.randint(1, days+1)
                    if origin_copy.loc[k_idx][n] == 1:
                        buf += 1
                        origin_copy.loc[k_idx][n] = np.random.choice([2,3,4])
            if s < 0:    
                while buf < abs(s):
                    n = np.random.randint(1, days+1)
                    if origin_copy.loc[k_idx][n] == 2 or origin_copy.loc[k_idx][n] == 3 or origin_copy.loc[k_idx][n] == 4:
                        buf += 1
                        origin_copy.loc[k_idx][n] = 1
    return origin_copy


def evaluation_1(origin_copy):
    eva = origin_copy.replace(1, 0)
    score = 0
    
    for k in range(len(eva)):
        x = ''.join([str(i) for i in np.array(eva.iloc[k:k+1]).flatten()])
        s = x.split('31')
        count = (len(s) -1)*3
        r = x.split('000')
        count += (1 -len(r))*2
        score += count
        score += np.sum([((2 - len(i))**2)*-1 for i in x.split('0') if len(i) >= 5])
        
    score += np.sum([abs(len(eva)*0.7 - (len(eva) - np.count_nonzero(eva[i] == 0)))*-4 for i in eva.columns]) 
    for i in range(len(eva.columns)):
        count_2 = eva[i+1].tolist().count(2)
        count_3 = eva[i+1].tolist().count(3)
        count_4 = eva[i+1].tolist().count(4)
        if count_2 >= 2 & count_3 >= 2 & count_4 >= 2:
            score += 30
        elif count_2 == 2 & count_3 == 2 & count_4 == 2:
            score += 70
        elif count_2 >=4 & count_3 <= 1 & count_4 <= 1:
            score += -40
        elif count_2 <=1 & count_3 >=4 & count_4 <= 1:
            score += -40
        elif count_2 <= 1 & count_3 <= 1 & count_4 >= 4:
            score += -40
    return score



def evaluation_2(origin_copy):
    eva = origin_copy.replace(1, 0)
    score = 0
    
    for k in range(len(eva)):
        x = ''.join([str(i) for i in np.array(eva.iloc[k:k+1]).flatten()])
        s = x.split('31')
        count = (len(s) -1)*3
        r = x.split('000')
        count += (1 -len(r))*2
        score += count
        score += np.sum([((2 - len(i))**2)*-1 for i in x.split('0') if len(i) >= 5])
        
    score += np.sum([abs(len(eva)*0.7 - (len(eva) - np.count_nonzero(eva[i] == 0)))*-4 for i in eva.columns]) 
    for i in range(len(eva.columns)):
        count_2 = eva[i+1].tolist().count(2)
        count_3 = eva[i+1].tolist().count(3)
        count_4 = eva[i+1].tolist().count(4)
        if count_2 >= 2 & count_3 >= 2 & count_4 >= 1:
            score += 30
        elif count_2 == 2 & count_3 == 2 & count_4 == 1:
            score += 70
        elif count_2 >=4 & count_3 <= 1 & count_4 <= 1:
            score += -40
        elif count_2 <=1 & count_3 >=4 & count_4 <= 1:
            score += -40
        elif count_2 <= 1 & count_3 <= 1 & count_4 >= 4:
            score += -40
    return score


def crossover(ep, sd, p1, p2):
    days = len(p1.columns)
    
    p1 = np.array(p1).flatten()
    p2 = np.array(p2).flatten()
    
    ch1 = []
    ch2 = []
    
    for p1_, p2_ in zip(p1,p2):
        x = True if  ep > random() else False
        
        if x == True:
            ch1.append(p1_)
            ch2.append(p2_)
        else:
            ch1.append(p2_)
            ch2.append(p1_)
            
    ch1, ch2 = mutation(sd,np.array(ch1).flatten(), np.array(ch2).flatten())
    
    ch1 = pd.DataFrame(ch1.reshape(int(len(ch1)/days), days))
    ch2 = pd.DataFrame(ch2.reshape(int(len(ch2)/days), days))
    
    ch1.columns = [i+1 for i in range(len(ch1.columns))]
    ch2.columns = [i+1 for i in range(len(ch1.columns))]
    
    return ch1, ch2



def mutation(sd, ch1, ch2):
    x = True if sd > random() else False
    
    if x == True:
        rand = np.random.permutation([i for i in range(len(ch1))])
        rand = rand[:int(len(ch1)//10)]
        for i in rand:
            if ch1[i] == 1:
                ch1[i] = 2
            elif ch1[i] == 2:
                ch1[i] = 3
            elif ch1[i] == 3:
                ch1[i] = 4
                
                
    x = True if sd > random() else False
    
    if x == True:
        rand = np.random.permutation([i for i in range(len(ch1))])
        rand = rand[:int(len(ch1)//10)]
        for i in rand:
            if ch2[i] == 1:
                ch2[i] = 2
            elif ch2[i] == 2:
                ch2[i] = 3
            elif ch2[i] == 3:
                ch2[i] = 4
    return ch1, ch2 



origin, holiday = read_excel()   
parent = []
for i in range(2):
    origin_copy = first_gene(origin, holiday)
    
    
    origin_copy = day_shift(origin_copy)
    origin_copy = day_shifts(origin_copy)
    origin_copy = holiday_fix(origin_copy, holiday)
    
    parent.append(origin_copy)
    


ep = 0.5
sd = 0.05

ch1,ch2 = crossover(ep, sd, parent[0], parent[1])

ch1 = holiday_fix(ch1,holiday)
ch2 = holiday_fix(ch2,holiday)

origin, holiday = read_excel()   
parent = []


for i in range(100):
    origin_copy = first_gene(origin, holiday)
    origin_copy = day_shift(origin_copy)
    origin_copy = day_shifts(origin_copy)
    origin_copy = holiday_fix(origin_copy, holiday)
    if len(origin_copy) >= 9:
        score = evaluation_1(origin_copy)
    else:
        score = evaluation_2(origin_copy)
    parent.append([score, origin_copy])

elite_length = 10
gene_length = 20

ep = 0.5
sd = 0.05

for i in range(gene_length):
    parent = sorted((parent), key=lambda x: x[0], reverse=True)
    
    parent = parent[:elite_length]
    
    
    if i == 0 or top[0] < parent[0][0]:
        top = parent[0]
    else:
        parent.append(top)


    children = []

    for k1, v1 in enumerate(parent):
        for k2, v2 in enumerate(parent):
            if k1 < k2:
                ch1, ch2 = crossover(ep, sd, v1[1], v2[1])
                ch1 = holiday_fix(ch1, holiday)
                ch2 = holiday_fix(ch2, holiday)
                
                if len(origin_copy) >= 9:
                    score1 = evaluation_1(ch1)
                    score2 = evaluation_1(ch2)  
                else:
                    score1 = evaluation_2(ch1)
                    score2 = evaluation_2(ch2)
            

            
                children.append([score1, ch1])
                children.append([score2, ch2])
            
    parent = children.copy()

x = top[1].replace(0, '休').replace(1, '休').replace(2, '日').replace(3, '早').replace(4, '遅')
x.to_excel('shift.xlsx')
