import pandas as pd

# 读取 Excel 文件
df = pd.read_excel('1.xlsx')

# 删除第一列数据
df = df.drop(df.columns[0], axis=1)

# 保存结果到新的 Excel 文件
df.to_excel('1.xlsx', index=False)
