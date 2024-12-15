import os
import re


def remove_comments_from_py():
    # 输入文件名（不含扩展名）
    file_name = input("请输入当前目录下的py文件名（不含扩展名）：").strip()
    full_file_name = f"{file_name}.py"

    # 检查文件是否存在
    if not os.path.isfile(full_file_name):
        print(f"文件 {full_file_name} 不存在！请检查后重试。")
        return

    # 读取原文件内容
    with open(full_file_name, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # 用于存储去掉注释的内容
    no_comment_lines = []

    # 正则表达式，用于匹配单行注释（# 开头的）
    comment_pattern = re.compile(r'^\s*#.*$')

    for line in lines:
        # 如果是注释行，跳过
        if comment_pattern.match(line.strip()) or line.strip().startswith('#'):
            continue

        # 去掉行内注释，保留有效代码
        line = re.sub(r'#.*$', '', line)
        no_comment_lines.append(line)

    # 生成新的文件名
    new_file_name = f"no_comments_{file_name}.py"

    # 写入无注释的内容到新文件
    with open(new_file_name, 'w', encoding='utf-8') as new_file:
        new_file.writelines(no_comment_lines)

    print(f"已成功生成无注释文件：{new_file_name}")


# 调用函数
remove_comments_from_py()