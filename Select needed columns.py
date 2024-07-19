import pandas as pd
from fuzzywuzzy import process

excel_path = 'test.xlsx'
excel_newpath = 'newtest.xlsx'

# 示例数据
data = {
   'old_headr1': [1, 2, 3],
   'old_header2': [4, 5, 6],
   'another_header': [7, 8, 9],
   'different_header': [10, 11, 12],
   'header_to_keep': [13, 14, 15]
}
#df = pd.DataFrame(data)

df = pd.read_excel(excel_path)
print(df)
# 目标表头列表
target_headers = ['Name', 'Age']
# 创建一个字典来存储匹配的结果
matched_headers = {}
for target in target_headers:
   matched_header, score = process.extractOne(target, df.columns)
   if score >= 70:  # 设置相似度阈值，70表示70%的匹配度
       matched_headers[matched_header] = target
# 显示匹配的结果
print("匹配结果:", matched_headers)
# 使用匹配结果创建新表
new_df = df[list(matched_headers.keys())]
# 重命名新表的表头为目标表头
new_df.rename(columns=matched_headers, inplace=True)
print("新表:")
print(new_df)
#df.to_excel(excel_newpath) #保存