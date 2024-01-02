import pandas as pd 
df = pd.read_excel('D:\\hc\\xsj\\program\\python\\salary\\top3- 2023年11月工资单.xlsx',skiprows=1, header=None,engine='openpyxl')
print(df)