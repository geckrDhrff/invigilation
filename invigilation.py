# 导入pandas模块
import pandas as pd

# 读取excel文件中的考场信息和老师信息
exam_info = pd.read_excel("exam_info.xlsx", sheet_name="exam_info")
teacher_info = pd.read_excel("exam_info.xlsx", sheet_name="teacher_info")

#每个考场的监考老师数量
TEACHER_NUM = int(exam_info["监考老师数量"].iloc[0])

#每个老师最多监考次数
max_times= (exam_info.shape[0] -1) * TEACHER_NUM / teacher_info.shape[0]

# 创建一个空的数据框，用于存储监考时间表
invigilation_table = pd.DataFrame(columns=["日期", "时间", "科目", "考场", "监考教师"])

# 创建一个空的字典，用于记录每个教师的监考次数
teacher_count = dict()

# 遍历教师信息，初始化每个教师的监考次数为0
for index, row in teacher_info.iterrows():
    teacher_count[row["姓名"]] = 0

# 遍历考试信息，为每场考试分配监考教师
for index, row in exam_info.iterrows():
    # 创建一个空的列表，用于存储分配的教师
    assigned_teachers = []
    exam_info["监考教师"] = None
    # 遍历教师信息，按照监考次数从小到大排序
    for teacher, count in sorted(teacher_count.items(), key=lambda x: x[1]):
        # 如果分配的教师数量已经达到要求，就跳出循环
        if len(assigned_teachers) == TEACHER_NUM:
            break
        # 如果教师没有监考冲突，就将其加入分配的教师列表，并更新其监考次数
        if not ((row["日期"] == exam_info["日期"]) & (row["时间"] == exam_info["时间"]) & (teacher in exam_info["监考教师"]) & (teacher_count[teacher] > max_times)).any():
            assigned_teachers.append(teacher)
            teacher_count[teacher] += 1
    # 将分配的教师列表转换为字符串，用逗号分隔
    assigned_teachers = ",".join(assigned_teachers)
    # 将当前考试的信息和分配的教师添加到监考时间表中
    new_row =  pd.DataFrame({'日期': [row["日期"]], '时间': [row["时间"]], '科目': [row["科目"]], '考场': [row["考场号"]], '监考教师': [assigned_teachers]})
    
    invigilation_table = pd.concat([invigilation_table,new_row],axis=0)

# 输出监考时间表
print(invigilation_table)

print(teacher_count)

# 将监考时间表保存为excel文件，假设文件名为invigilation_table.xlsx
invigilation_table.to_excel("invigilation_table.xlsx", index=False)
