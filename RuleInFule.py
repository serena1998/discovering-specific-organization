#!/usr/bin/env python
# coding: utf-8

# In[1]:


#生成嫌疑人名单
# -*- coding:utf-8 -*-
import sys,xlrd,xlwt

ipyout = sys.stdout
reload(sys)
sys.setdefaultencoding('utf-8')
sys.stdout = ipyout
#第一批确定人员的个人信息
#fileFirstClientInfo = u'698传销大团伙中头目的客户基本信息表.xlsx' 重复数据
# fileFirstNongAccountInfo = 
# fileFirstIdentifyInfo = 


# In[2]:


firstIdentifyInfo = xlrd.open_workbook(u'全数据/698传销大团伙中头目的身份信息.xlsx')
firstIdentityTable = firstIdentifyInfo.sheets()[0]
print (firstIdentityTable.nrows3)
firstIdentityDictionary = {} #证件号：姓名
for row in  range(firstIdentityTable.nrows):
    if row ==0:
        continue
    identityNum =firstIdentityTable.cell(row,2).value
    if identityNum not in firstIdentityDictionary.keys():
        firstIdentityDictionary[identityNum] = firstIdentityTable.cell(row,0).value
    else:
        print( identityNum )# 230183199210180717，220523199112231617有两手机号，table为重复记录
print (len(firstIdentityDictionary))


# In[3]:


firstAccountInfo=xlrd.open_workbook(u'全数据/698传销大团伙中头目的农行开户基本信息表.xlsx')
firstAccountTable = firstAccountInfo.sheets()[0]
firstAccountDictionary={} #交易卡号:证件号
for row in  range(firstAccountTable.nrows):
    if row ==0:
        continue
    accountNum =firstAccountTable.cell(row,2).value
    if accountNum not in firstAccountDictionary.keys():
        firstAccountDictionary[accountNum] = firstAccountTable.cell(row,1).value
    if firstAccountTable.cell(row,1).value not in firstIdentityDictionary.keys():
        firstIdentityDictionary[firstAccountTable.cell(row,1).value]=firstAccountTable.cell(row,0).value
print (len(firstAccountDictionary), len(firstIdentityDictionary))
# print firstAccountDictionary.keys()
count = 0
# for keyv in firstAccountDictionary.keys():
#     print keyv,":",firstAccountDictionary[keyv]
#     count +=1
#     if count >5:
#         break


# In[4]:


rivalAccountDictionary={} #银行流水中对手账号：对手户名
with open(u'C:\Users\forev\Desktop\data.txt') as f:  #698传销大团伙中头目的农行流水.xlsx对应的txt文件
    dataRecordsline = f.readlines()
    
tags = dataRecordsline[0].strip().split('|')
dataFlueDictionary={}
dataRecords=[]
index = 0
# for tag in tags:
#     print index, tag
#     index += 1
# length = len(tags)

for dataLine in dataRecordsline[1:]:
    items = dataLine.strip().split("|")
    if items[2] == "":
        continue
    price = float(items[5])
    if abs(price) < 100 or items[9] == "":
        continue
    inout = "进"
    if price < 0:
        inout = "出"
#     print "items[9]",items[9]
    tempData = [""]*10
    tempData[0] = items[2] #账卡号
    tempData[1] = items[3] #交易日期
    tempData[2] = abs(price)   #交易金额
    tempData[3] = items[6] #交易后余额
    tempData[4] = inout    #收付标志
    tempData[5] = items[9] #对手账户
    tempData[6] = items[10] #对手名称
    tempData[7] = items[7]  #交易摘要
    tempData[8] = items[17] #交易网点
    tempData[9] = "1" if items[2] in firstAccountDictionary.keys() else "0" #如果对手账号在卡号证件字典中则存1，否则为0
    dataRecords.append(tempData)
    if tempData[0] not in dataFlueDictionary.keys():
        dataFlueDictionary[tempData[0]]=[tempData]
    else:
        dataFlueDictionary[tempData[0]].append(tempData)
    if items[9] not in rivalAccountDictionary.keys():
        if items[10] != "":
#             print items[10],items[9]
            rivalAccountDictionary[items[9]] = items[10]  #当对手户名不为空时，将对应关系加入字典中
print (len(dataRecords),len(dataFlueDictionary) )#记录条数，交易账号个数709个


# In[5]:


# for item in dataRecords[0:5]:
#     print item
# print [keys for keys in rivalAccountDictionary.keys()][1:10],len(rivalAccountDictionary) #对手账号个数


# In[6]:


secondAccountInfo=xlrd.open_workbook(u'全数据/698二次查询农行开户基本信息表.xlsx')
secondAccountTable = secondAccountInfo.sheets()[0]
secondAccountDictionary={} #交易卡号:证件号
secondIdentityDictionary={}#证件号：姓名
for row in  range(secondAccountTable.nrows):
    
    if row ==0:
        continue
    identityNum = secondAccountTable.cell(row,1).value
    if identityNum not in secondIdentityDictionary.keys():
        secondIdentityDictionary[identityNum]= secondAccountTable.cell(row,0).value
    accountNum =secondAccountTable.cell(row,2).value
    if accountNum not in secondAccountDictionary.keys():
        secondAccountDictionary[accountNum] = secondAccountTable.cell(row,1).value 
print (len(secondAccountDictionary), len(secondIdentityDictionary))#第二批交易卡个数652张；第二批交易人450个


# In[7]:


# secondRecordInfo = xlrd.open_workbook(u'全数据/698二次查询农行流水.xlsx')
# secondRecordTable = secondRecordInfo.sheets()[0] #取第一个table
# index = 0
# for item in secondRecordTable.row_values(0):
#     print "index",index,item,
#     index = index+1


# In[8]:


# secondFlueRecords=[] #第二批流水所有记录；修改为继续添加到dataRecors中
secondFlueDictionary={} #交易账号：[0-9]
secondRecordInfo = xlrd.open_workbook(u'全数据/698二次查询农行流水.xlsx')
for sheet in range(12):
    secondRecordTable = secondRecordInfo.sheets()[sheet] #取第一个table
    for row in range(secondRecordTable.nrows):
        if row == 0:
            continue
        tempRecord=secondRecordTable.row_values(row)

        price = float(tempRecord[6].encode('utf-8'))
        if price < 100 or tempRecord[9] == "" :
            continue
        accountNum = tempRecord[0]
        timeValue = "".join(tempRecord[5].split(" ")[0].split("-"))
        tempSecondData = [""]*10
        tempSecondData[0] = accountNum  #交易账卡号
        tempSecondData[1]=timeValue #交易日期
        tempSecondData[2] = price  #交易金额
        tempSecondData[3:9]=tempRecord[7:10]+[tempRecord[12]]+ [tempRecord[15]]+[tempRecord[17]] #交易金额,交易余额,收付标志,对手账号,摘要说明，
        tempSecondData[9] = "0" if accountNum not in firstAccountDictionary.keys() else "1"
        dataRecords.append(tempSecondData)
        if accountNum not in dataFlueDictionary.keys():
            dataFlueDictionary[accountNum] = [tempSecondData]
        else:
            dataFlueDictionary[accountNum].append(tempSecondData)
        if tempRecord[9] not in rivalAccountDictionary.keys():
            if tempRecord[12] != "":
                rivalAccountDictionary[tempRecord[9]] = tempRecord[12]  #当对手户名不为空时，将对应关系加入字典中
print (len(dataRecords),len(dataFlueDictionary),len(rivalAccountDictionary))


# In[1]:


from __future__ import divisino
# print 3/2
sensitivePurchaseAmountList = []
for first in [0,1]:
   for second in range(21):
       if first == 0 and second ==0:
           continue
       amount = first*3800+second*3300
       sensitivePurchaseAmountList.append(amount)
print (sensitivePurchaseAmountList)


# In[10]:


rebateAmountList = [570,190,760,380,1140,152,456,1672,114,1976,7600,5092,1520,4864,836,1900,6612,7904,2394,1596]
sensitiveRebateAmountList=[]
for times in range(2,6):
    sensitiveRebateAmountList=sensitiveRebateAmountList +[item*times for item in rebateAmountList]
#     print sensitiveRebateAmountList

# print  len(sensitiveRebateAmountList)


# In[11]:


sensitivePurchaseAmountList.append([50800,70000])
def MedianAmount(List):
    List.sort()
    half = len(List)//2 
    return (List[half]+List[~half])/2 #中值
#计算各进交易统计值；各出交易统计值
def TransactionStatistic(ItemtransactionList):
    count = len(ItemtransactionList) #次数
    if count ==0:
        return 0,0,0,0,0
    amountList = [item[2] for item in ItemtransactionList] #金额列表
#     print amountList
    amount = sum(amountList) #总额
    averageAmount = amount/count if count !=0 else 0 #均值
    #进中值
    amountList.sort()
    maxAmount = amountList[-1] #最大值
    medianAmount = MedianAmount(amountList)
    return count,amount,averageAmount,maxAmount,medianAmount
#计算进敏感统计值,出敏感统计值
def SensitiveStastic(ItemtransactionList):
    sensitivePurchase = [] #申购对手列表
    sentitiveRebate = []   #返利对手列表
    purchaseCount = 0 #申购金额次数
    purchaseAmount = 0 #申购金额总额
    rebateCount = 0 #返利金额次数
    rebataAmount = 0 #返利金额总额
    for itemTranscation in ItemtransactionList:
        if itemTranscation[2] in sensitivePurchaseAmountList: #如果当前金额为申购金额
            purchaseCount += 1  #申购金额次数
            purchaseAmount += itemTranscation[2] #申购总额
            if itemTranscation[5] not in sensitivePurchase:#当前对手不在申购对手列表中
                sensitivePurchase.append(itemTranscation[5])
        if itemTranscation[2] in sensitiveRebateAmountList: #如果当前金额为返利金额
            if itemTranscation[2] == 3800 and itemTranscation[7].find(u'工资') ==-1:
                continue
            rebateCount += 1   #返利金额次数
            rebataAmount += itemTranscation[2] #返利总额
            if itemTranscation[5] not in sentitiveRebate:#当前对手不在返利对手列表中
                sentitiveRebate.append(itemTranscation[5])
#     print "SensitiveStastic",purchaseCount,purchaseAmount,sensitivePurchase,rebateCount,rebataAmount,sentitiveRebate
    return  [purchaseCount,purchaseAmount,sensitivePurchase,rebateCount,rebataAmount,sentitiveRebate]

#计算月进交易统计值；月出交易统计值
from operator import itemgetter #itemgetter用来去dict中的key，省去了使用lambda函数
from itertools import groupby #itertool还包含有其他很多函数，比如将多个list联合起来。。
def MonthStastic(ItemtransactionList):
    monthStasticList=[]
    for dateItem in ItemtransactionList:
#         print "zhanghao ",dateItem[0],len(ItemtransactionList)
        if dateItem[1][0:4] =="2017" :
            monthDictionary={}
            monthDictionary["month"] = dateItem[1][4:6]
            monthDictionary["day"] = dateItem[1][6:8]
            monthDictionary["amount"] = dateItem[2]
            monthDictionary["rivalAccount"] = dateItem[5]
            monthStasticList.append(monthDictionary)
#     print "monthStasticList ",monthStasticList
    monthStasticList.sort(key = itemgetter("month"))  
    monthSorted = groupby(monthStasticList,key = itemgetter("month")) #========
    monthValueDic = dict([(int(key),sum([g["amount"] for g in group])) for key,group in monthSorted])
    monthGrouped = dict([(int(key),list(group)) for key,group in monthSorted]) #按月分组
    totalCount = []
    rivalCountList=[]
#     mounthMedianRival=0
    dayAccountMaxList =[] #最多次
    dayValueMaxList=[]  #最大总额
    dayValueMaxMoneyList=[] #最大总额数
    dayAccountMaxCountList=[] #最多次总额
    dayAccountMaxMoneyList=[] #最多次总额
    dayValueMaxCountList=[]
    for monthNum in range(1,10):
        if monthGrouped.has_key(monthNum):
            totalCount.append(monthValueDic[monthNum])
            groupedList = monthGrouped[monthNum]
            groupedList.sort(key=itemgetter("rivalAccount"))
            rivalAccoutSorted = groupby(groupedList,itemgetter("rivalAccount"))
            rivalCount=len([key for key,group in rivalAccoutSorted]) # 月对手个数
            rivalCountList.append(rivalCount)         
            
            groupedList.sort(key=itemgetter("day"))
            
            dayAccoutSorted = groupby(groupedList,itemgetter("day"))
            dayAccountCount = dict([(int(key),len(list(group))) for key,group in dayAccoutSorted])
            dayAccountMax = max(dayAccountCount,key = lambda x:dayAccountCount[x])  #每月交易次数最多对应的天数
            dayAccountMaxList.append(dayAccountMax)
            
            dayAccountMaxCount=dayAccountCount[dayAccountMax]  #月交易次数最多一天交易的笔数
            dayAccountMaxCountList.append(dayAccountMaxCount)
            
            dayAccoutSorted = groupby(groupedList,itemgetter("day"))
            dayValueDic = dict([(int(key),sum([g["amount"] for g in group])) for key,group in dayAccoutSorted])
            dayValueMax = max(dayValueDic,key = lambda x:dayValueDic[x])  #每月交易金额最大总金额对应的天数
#             print "dayValueMax",dayValueMax
            dayValueMaxList.append(dayValueMax)
            dayValueMaxMoneyList.append(dayValueDic[dayValueMax]) #月交易最大总金额天的交易总钱数
            
            dayValueMaxCount = dayAccountCount[dayValueMax] 
            dayValueMaxCountList.append(dayValueMaxCount) #月交易金额最大天的交易笔数
            
            dayAccountMaxMoneyList.append(dayValueDic[dayAccountMax]) #月交易次数最多天的交易总钱数
        else:
            totalCount.append(0) #月交易总额

            rivalCountList.append(0)
            
            dayValueMaxList.append(0)  #最大总额天           
            dayValueMaxMoneyList.append(0) #最大总额数天的钱数
            dayValueMaxCountList.append(0) #最大总额天笔数
            
            dayAccountMaxList.append(0)#最多笔交易天
            dayAccountMaxMoneyList.append(0) #最多笔交易天总额数
            dayAccountMaxCountList.append(0) #最多笔交易天的笔数
#         monthList.append(rivalCount)
        mounthMedianRival = MedianAmount(rivalCountList) #月交易对手个数中值
     
    return [totalCount,rivalCountList,mounthMedianRival,dayValueMaxList,dayValueMaxMoneyList,dayValueMaxCountList,                                            dayAccountMaxList,dayAccountMaxMoneyList,dayAccountMaxCountList]

#统计与对手之间流水
def InOutTransactionStastic(ItemtransactionList):
    rivalStasticList=[]
    rivalDictionaryStastic={}

    for rivalItem in ItemtransactionList:
        rivalDictionary={}
        rivalDictionary["amount"] = rivalItem[2]
        rivalDictionary["rivalAccount"] = rivalItem[5]
        rivalStasticList.append(rivalDictionary)
    rivalStasticList.sort(key = itemgetter("rivalAccount"))  
    rivalSorted = groupby(rivalStasticList,key = itemgetter("rivalAccount")) #========
    rivalGrouped = dict([(key,list(group)) for key,group in rivalSorted]) #按对手账号分组    
    for key in rivalGrouped.keys():
        rivalPurchase=0
        rivalRebate=0
        rivalItemList = rivalGrouped[key]
#         print "key,rivalItemList",key,rivalItemList
        rivalCount = len(rivalItemList) #交易次数
        rivalItemList.sort(key=lambda x:x["amount"]) #金额升序排列
        rivalItemSortedList=[item["amount"] for item in rivalItemList ] 
        rivalSumAmount = sum(rivalItemSortedList) #交易总金额
        rivalAverageAmount = rivalSumAmount/rivalCount #交易平均
        rivalMaxAmount = rivalItemSortedList[-1] #交易最大金额
        rivalMedianAmout = MedianAmount(rivalItemSortedList) #交易中值
        for amountItem in rivalItemSortedList:
            if amountItem in sensitivePurchaseAmountList:
                rivalPurchase += 1 
            elif amountItem in sensitiveRebateAmountList:
                rivalRebate +=1
        #该账号与对手账号的交易次数，交易总额，交易平均，交易最大金额，交易中值
        rivalDictionaryStastic[key] = [rivalCount,rivalSumAmount,rivalAverageAmount,rivalMaxAmount,rivalMedianAmout,rivalPurchase,rivalRebate]
    return rivalDictionaryStastic
transactionStatistic = {} #交易总统计
sensitiveStatistic = {} #敏感交易统计
monthStatistic = {}
inRivalStatistic = {} #与对手进交易的统计词典
outRivalStatistic = {} #与对手出交易的统计词典
allRicalStatistic ={} #与对手总交易的统计词典
for account in dataFlueDictionary.keys():
#     if account !="6228480309006236074":
#         continue
    accountItemList = dataFlueDictionary[account]
    inItemRivalList = [] #进对手列表
    inItemTransactionList = [] #进交易流水列表
    outItemRivalList = [] #出对手列表
    outItemTransactionList = [] #出交易流水列表
    inPurchaseRivalList = [] #进申购对手卡列表
    inPurchaseRivalList = [] #进返利对手卡列表
    allItemTransactionList = [] #与对手的进出交易流水列表

    for accountItem in accountItemList:
        if accountItem[6] == "":  #对手名称，为了避免空值
            accountItem[6] = rivalAccountDictionary[accountItem[5]] if accountItem[5] in rivalAccountDictionary.keys()                                                                    else accountItem[7]
        transFlag = accountItem[4]
        allItemTransactionList.append(accountItem)
        if transFlag == u"进": #将对手加入进对手列表
#             print "accountItem",accountItem
            inItemTransactionList.append(accountItem)
            if accountItem[5] not in inItemRivalList:
                inItemRivalList.append(accountItem[5])
        else: #将对手加入出对手列表
            outItemTransactionList.append(accountItem)
            if accountItem[5] not in outItemRivalList:
                outItemRivalList.append(accountItem[5])
    #进出额统计
    inRivalCount = len(inItemRivalList)
    incount,inamount,inaverageAmount,inmaxAmount,inmedianAmount = TransactionStatistic(inItemTransactionList)
    outRivalCount = len(outItemRivalList)
    outcount,outamount,outaverageAmount,outmaxAmount,outmedianAmount = TransactionStatistic(outItemTransactionList)
    transactionStatistic[account]=[incount,inamount,inaverageAmount,inmaxAmount,inmedianAmount,inRivalCount,inItemRivalList,                                  outcount,outamount,outaverageAmount,outmaxAmount,outmedianAmount,outRivalCount,outItemRivalList]
    #进出申购、返利统计

    sensitiveStatistic[account]=SensitiveStastic(inItemTransactionList)+ SensitiveStastic(outItemTransactionList)
    

    #月进出对手、天统计
    monthStatistic[account]= MonthStastic(inItemTransactionList)+ MonthStastic(outItemTransactionList)
    
    #     #进\出对手字典
    inRivalStatistic[account] = InOutTransactionStastic(inItemTransactionList)
    outRivalStatistic[account] = InOutTransactionStastic(outItemTransactionList)
    allRicalStatistic[account] = InOutTransactionStastic(allItemTransactionList)


# In[1]:


Name_Dic={"孙广泛":"1","孙朝森":"2","张博洋":"3","张伟":"4","贾复震":"5"}
for name in Name_Dic.keys():
   name_account = 
   #用于后续绘制特定人物（Name_Dic）之间的资金流关系，生成testLv.js,绘制funds.html
   if inname in Name_Dic.keys() or outname in Name_Dic.keys(): #判断出姓名是否在字典中 
       if inname not in account_dic.keys(): #name_inout[inname]
           account_dic[inname]=name_inout[inname][0]-name_inout[inname][2]
       if  outname not in account_dic.keys(): #name_inout[inname]
           account_dic[outname]=name_inout[outname][0]-name_inout[outname][2]
# print transactionStatistic["6228480309006236074"] #张博洋账号
# print transactionStatistic["6228480308919345279"] #孙广泛还没有
# print sensitiveStatistic["6228480309006236074"]
# print monthStatistic["6228480309006236074"] 
# print inRivalStatistic["6228480309006236074"]


# In[13]:


#     print "SensitiveStastic",purchaseCount,purchaseAmount,sensitivePurchase,rebateCount,rebataAmount,sentitiveRebate
SuspiciousPurchaseList=[]
SuspiciousRebateList=[]
#判断是否为敏感申购账户
for account in dataFlueDictionary.keys():
    incount = transactionStatistic[account][0] #进交易次数
    inamount = transactionStatistic[account][1]  #进交易总额
    inaverageAmount = transactionStatistic[account][2] #进平均
    inRivalCount = transactionStatistic[account][5] #进对手数
    inPurchaseCount = sensitiveStatistic[account][0] #进申购交易次数
    inPurchaseAverage = sensitiveStatistic[account][1]/inPurchaseCount if inPurchaseCount !=0 else 0#跟每个对手的进申购平均值
    inPurchaseRivalCount = len(sensitiveStatistic[account][2]) #进申购对手数
    
    outcount = transactionStatistic[account][7] #出交易次数
    outamount = transactionStatistic[account][8]  #出交易总额
    outaverageAmount = transactionStatistic[account][9] #出平均
    outRivalCount = transactionStatistic[account][12] #出对手数
   
    if (incount-outcount)/(incount+outcount) >0.5: #进出交易次数差值
        if (inRivalCount-outRivalCount)/(inRivalCount+outRivalCount)>0.5: #进出对手数差值
            if (inamount-outamount)/(inamount+outamount) <0.3: #资金主要转出没有沉淀
                if inPurchaseCount>10: #进申购次数
                    maxRival = max(outRivalStatistic[account].items(),key = lambda x:x[1][1]) #(最大出对手)
                    if maxRival[1][1]/outamount >0.5: #给最大出对手的资金比例
                        if maxRival[1][2] >=outaverageAmount: #给最大处对手的出平均大于本账户的总出平均
                            SuspiciousPurchaseList.append(account)
#                             print account,maxRival[0]

    outPurchaseCount = sensitiveStatistic[account][9] #出返利交易次数

    if (outcount-incount)/(incount+outcount) >0.5: #出进次数差值
        if (outRivalCount-inRivalCount)/(inRivalCount+outRivalCount)>0.5: #进出对手数差值
               if (inamount-outamount)/(inamount+outamount) <0.3: #资金主要转出没有沉淀
                    if outPurchaseCount>10:
                         if outaverageAmount < inaverageAmount: #出平均小于进平均
                            #月出最多比日期
                            outdayAccountMaxList = monthStatistic[account][13] #月出最多笔的日期
                            outdayAccountMaxMoneyList = monthStatistic[account][14] #月出最多笔时金额
                            #月进最大笔资金的日期
                            indayValueMaxList = monthStatistic[account][2]
                            indayValueMaxMoneyList=  monthStatistic[account][3]
                                                      
                            for index in range(len(outdayAccountMaxList)):
                                if int(outdayAccountMaxList[index])-int(indayValueMaxList[index])<2 and                                    int(outdayAccountMaxList[index])>=int(indayValueMaxList[index]): #最大进钱日期与最大出钱日期相差不到两天
                                        if indayValueMaxMoneyList[index] !=0 and                                            outdayAccountMaxMoneyList[index]/indayValueMaxMoneyList[index] >0.5:
                                                SuspiciousRebateList.append(account)
#                                                 print ":::",account
                                                break


# In[14]:


def SuspiciousAccountItemList(suspiciouslist):
    ItemList=[]
    for accountItem in suspiciouslist:
        outmaxRival = max(outRivalStatistic[accountItem].items(),key = lambda x:x[1][1])[0] #(出总额最大的对手)
        inmaxRival = max(inRivalStatistic[accountItem].items(),key = lambda x:x[1][2])[0]#(进总额最大的对手)
        flag = dataFlueDictionary[accountItem][0][9]
#         print flag,dataFlueDictionary[accountItem][0]
#         ItemFlag = u"第一批" if flag == "1" else u"第二批"
        if  flag== "1":
#             print "firstAccountDictionary[accountItem]",firstAccountDictionary[accountItem]
            ItemFlag = u"第一批"
            accountName = firstIdentityDictionary[firstAccountDictionary[accountItem]]
        else:
            if accountItem in secondAccountDictionary.keys():
                ItemFlag = u"第二批"
                accountName = secondIdentityDictionary[secondAccountDictionary[accountItem]]
            else:
                ItemFlag = u"待调取"
                accountName = "NULL"
    
        Item = [accountItem,accountName,ItemFlag]+transactionStatistic[accountItem][0:6]            +[inmaxRival]            +transactionStatistic[accountItem][7:13]+[outmaxRival]            +sensitiveStatistic[accountItem][0:2]+sensitiveStatistic[accountItem][3:5]            +sensitiveStatistic[accountItem][6:8]+sensitiveStatistic[accountItem][9:11]
        ItemList.append(Item)
    return ItemList
SuspiciousPurchaseListItem=SuspiciousAccountItemList(SuspiciousPurchaseList)
SuspiciousRebateListItem=SuspiciousAccountItemList(SuspiciousRebateList)
print len(SuspiciousPurchaseListItem),len(SuspiciousRebateListItem)


# In[17]:


wb = xlwt.Workbook(encoding = 'utf8')
ws = wb.add_sheet('Purchase')
itemFlag = [u"账号",u"姓名",u"涉案标记",u"进次数",u"进总额",u"进平均",u"进最大",u"进中值",u"进对手数",u"进最大对手",           u"出次数",u"出总额",u"出平均",u"出最大",u"出中值",u"出对手数",u"出最大对手",           u"进申购金额次数",u"进申购总额",u"进返利金额次数",u"进返利总额",            u"出申购金额次数",u"出申购总额",u"出返利金额次数",u"出返利总额"]
for itemindex in range(len(itemFlag)):
    ws.write(0,itemindex,itemFlag[itemindex])
for row in range(len(SuspiciousPurchaseListItem)):
    for col in range(len(itemFlag)):
        ws.write(row+1,col,str(SuspiciousPurchaseListItem[row][col]))

wr = wb.add_sheet(u'Rebate')
for itemindex in range(len(itemFlag)):
    wr.write(0,itemindex,itemFlag[itemindex])
for row in range(len(SuspiciousRebateListItem)):
    for col in range(len(itemFlag)):
        wr.write(row+1,col,str(SuspiciousRebateListItem[row][col]))   
wb.save(u'全数据/嫌疑名单.xls')


# In[ ]:





# In[ ]:




