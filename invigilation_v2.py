# 导入所需的模块
import pandas as pd

# 读取excel文件中的考场信息和老师信息
exam_info = pd.read_excel("exam_info.xlsx", sheet_name="exam_info")
teacher_info = pd.read_excel("exam_info.xlsx", sheet_name="teacher_info")

#监考老师数量
TEACHER_NUM = int(exam_info["监考老师数量"].iloc[0])

# 创建一个空的数据框，用于存储监考表
invigilation_table = pd.DataFrame(columns=["考场号", "科目", "日期", "时间", "监考老师"])

# 创建一个空的字典，用于存储每个老师的总监考时间
teacher_time = {}

#每个老师最多监考次数
max_times= (exam_info.shape[0] * TEACHER_NUM -1)/ teacher_info.shape[0]

# 遍历每个老师，初始化他们的总监考时间为0
for name in teacher_info["姓名"]:
    teacher_time[name] = 0

# 遍历每个考场，使用贪心算法分配监考老师
for index, row in exam_info.iterrows():
    # 获取当前考场的考试日期和时间
    exam_date = row["日期"]
    exam_time = row["时间"]
    # 创建一个空的列表，用于存储分配给当前考场的监考老师
    teachers = []
    # 重复TEACHER_NUM次，为当前考场分配监考老师
    for i in range(TEACHER_NUM):
        # 从老师信息中随机抽取一名老师
        teacher = teacher_info.sample(n=1, replace=False)
        # 获取该老师的姓名
        name = teacher["姓名"].iloc[0]
        # 获取该老师的总监考时间
        time = teacher_time[name]
        # 如果该老师已经被分配给当前考场，或者该老师的总监考时间超过平均值，或者该老师已经被分配到同一日期同一时刻的其他考场，则选择一个总监考时间最少的老师
        while name in teachers or time > max_times or name in invigilation_table[(invigilation_table["日期"] == exam_date) & (invigilation_table["时间"] == exam_time)]["监考老师"].explode().tolist():
            teacher = teacher_info.sort_values(by=["总监考时间"],ascending=[True]).iloc[0]
            name = teacher["姓名"]
            time = teacher_time[name]
        # 将该老师添加到分配给当前考场的列表中
        teachers.append(name)
        # 更新该老师的总监考时间
        teacher_time[name] += 1
        teacher_info.loc[teacher_info["姓名"]==name,"总监考时间"] = teacher_time[name]
    # 将考场信息和老师信息合并为一行，添加到监考表中
    invigilation_table.loc[index] = [row["考场号"], row["科目"], row["日期"], row["时间"], teachers]

# 将监考表保存为excel文件
invigilation_table.to_excel("invigilation_table.xlsx", index=False)

# 打印监考表
print(invigilation_table)
print("-"*20)
print(max_times)
print(teacher_time)