import os
import re
import pandas as pd

# 定义文件夹路径和学生名单 Excel 文件路径
folder_path = ".\lab1"  # 替换为你的实验报告文件夹路径
student_list_file = ".\学生名单.xlsx"  # 替换为你的学生名单 Excel 文件路径
# 设置实验编号
experiment_number = 1  # 修改为当前实验的编号


def load_students_from_excel(file_path):
    """从 Excel 文件加载学生名单"""
    df = pd.read_excel(file_path)  # 读取 Excel 文件
    students = {}
    for _, row in df.iterrows():
        student_id = str(row["学号"]).strip()  # 学号
        student_name = row["姓名"].strip()  # 姓名
        students[student_id] = student_name
    return students


def remove_spaces(filename):
    """移除文件名中的多余空格"""
    return re.sub(r'\s+', ' ', filename.strip())


def process_files(folder_path, students, experiment_number):
    """批量重命名文件并找出未提交报告的学生"""
    submitted_ids = set()  # 存储已提交的学号

    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):  # 只处理 .docx 文件
            original_filename = filename
            filename = remove_spaces(filename)  # 移除多余空格
            file_path = os.path.join(folder_path, original_filename)

            new_name = None
            student_id = None
            student_name = None

            # 匹配学号
            for id in students.keys():
                if id in filename:  # 如果学号在文件名中，记录学号
                    student_id = id
                    student_name = students[id]
                    submitted_ids.add(student_id)  # 标记为已提交
                    break

            # 匹配姓名
            if not student_id:
                for name in students.values():
                    if name in filename:  # 如果姓名在文件名中
                        student_name = name
                        # 通过姓名反查学号
                        student_id = [id for id, n in students.items() if n == name][0]
                        submitted_ids.add(student_id)  # 标记为已提交
                        break

            # 如果匹配成功，生成新文件名
            if student_id and student_name:
                new_name = f"{student_id}-{student_name}-实验{experiment_number}.docx"
            else:
                print(f"文件 {original_filename} 不符合任何已知学生名单，跳过。")
                continue

            # 重命名文件
            new_file_path = os.path.join(folder_path, new_name)
            os.rename(file_path, new_file_path)
            print(f"文件 {original_filename} 重命名为 {new_name}")

    # 找出未提交的学生
    missing_students = {id: name for id, name in students.items() if id not in submitted_ids}

    # 将未提交学生名单追加到文件
    with open("未提交实验报告名单.txt", "a", encoding="utf-8") as file:
        file.write(f"\n未提交报告{experiment_number}的学生名单：\n")
        if missing_students:
            for student_id, student_name in missing_students.items():
                file.write(f"{student_id}\t{student_name}\n")
        else:
            file.write("所有学生都已提交报告。\n")

    return missing_students



# 从 Excel 文件加载学生名单
students = load_students_from_excel(student_list_file)

# 调用函数
missing_students = process_files(folder_path, students, experiment_number)

# 输出未提交的学生名单
if missing_students:
    print(f"\n未提交报告{experiment_number}的学生名单已追加到 '未提交实验报告名单.txt' 文件。")
else:
    print(f"\n所有学生都已提交报告{experiment_number}，结果已追加到 '未提交实验报告名单.txt' 文件。")
