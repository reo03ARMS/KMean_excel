import pandas as pd
from sklearn.cluster import KMeans
import openpyxl

excel_data = "./SampleData.xlsx"

df = pd.read_excel(excel_data,index_col=0) #excelデータ入力


#グループ数指定
group_count = 5



#グループごとのデータフレーム作成クラス
class CreateDF():
    def __init__(self,lists,count):
        self.count = count
        self.list = lists
    def create(self):
        sub_df = df.iloc[self.list]

        with pd.ExcelWriter(excel_data,mode="a") as writer:
            sub_df.to_excel(writer, sheet_name='new_sheet'+str(self.count))

            
            

#KMean法でクラスタリング
def KMean(df):
    #excelファイル整理
    wb = openpyxl.load_workbook(excel_data)
    for i in range(group_count):
        try:
            wb.remove(wb['new_sheet'+str(i)])
            wb.save(excel_data)
            # print(wb.worksheets)
        except KeyError:
            # print(wb.worksheets)
            pass
    

    
    kmeans_model = KMeans(n_clusters=group_count,random_state=10).fit(df)#5グループに分ける(後に入力された数に)
    label = kmeans_model.labels_
    label = label.tolist()

    

    for i in range(group_count):
        #グループごとにまとめてDataFrame化
        keep_list = [j for j, x in enumerate(label) if x == i]
        # print(keep_list)
        a = CreateDF(keep_list,i)
        a.create()




KMean(df)



