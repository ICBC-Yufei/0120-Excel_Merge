import os
import time
import sys
import pandas as pd

# 将文件读取出来放一个列表里面
pwd = './待合并Excel' # 获取文件目录
# 新建列表，存放文件名
file_list = []
# 新建列表存放每个文件数据(依次读取多个相同结构的Excel文件并创建DataFrame)
dfs = []

print("-----------------------说明----------------------------\n"
      "注1：默认方式适用于第一行即为标题所在行，且表尾不存在无关行。\n"
      "注2：若不选择默认方式执行，则需输入相关参数才能执行此程序。\n"
      "注3：Y代表Yes，N代表No。\n"
      "-------------------------------------------------------")

Begin_Judge = input("是否选择默认方式执行？请输入Y或N：")
if Begin_Judge == "Y" or Begin_Judge == "y":
  begin = 0
  end = None
elif Begin_Judge == "N" or Begin_Judge == "n":
  begin = int(input("请输入表头所在行："))-1
  # Judge = input("表尾是否存在备注行？请输入Y/N：")
  # if Judge == "Y":
  end = int(input("请输入表尾无关行数量："))
  if end == 0:
    end = None
  else:
    end = -1 * end
  # elif Judge == "N":
  #   end = None
  # else:
  #   print("输入有误！请检查")
  #   sys.exit()
else:
  print("输入有误，请检查后重新运行程序！")
  sys.exit()

name = input("请输入文件保存名称（自动补充日期信息）：")

for root,dirs,files in os.walk(pwd): # 第一个为起始路径，第二个为起始路径下的文件夹，第三个是起始路径下的文件。
  for file in files:
    file_path = os.path.join(root, file)
    file_list.append(file_path) # 使用os.path.join(dirpath, name)得到全路径
    df = pd.read_excel(file_path , header = begin, dtype=str) # 设置指定表头名称
    df = df[:end]
    print("Now: " + file + " Merge Completed!")
    dfs.append(df)

# 将多个DataFrame合并为一个
df = pd.concat(dfs)

df.dropna(axis=0, how="all", inplace=True)

print("---------所有文件合并已完成！---------")
# 写入excel文件，不包含索引数据

localtime = time.localtime(time.time())
time = time.strftime('%Y%m%d%H%M',time.localtime(time.time()))
path = os.path.abspath(os.path.dirname(sys.argv[0]))

df.to_excel(path+'\\'+str(time)+name+'.xls', index=False)

del_list = os.listdir(pwd)
for file in del_list:
  file_path = os.path.join(pwd, file)
  if os.path.isfile(file_path):
    os.remove(file_path)

print("----原路径下待合并文件已自动删除！----")
