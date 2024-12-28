import os
import re


def remove_comments_and_empty_lines_from_py():
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

    # 用于存储处理后的内容
    processed_lines = []

    # 正则表达式，用于匹配单行注释（# 开头的）
    comment_pattern = re.compile(r'^\s*#.*$')

    for line in lines:
        # 去掉两端空白字符
        stripped_line = line.strip()

        # 跳过注释行和空行
        if comment_pattern.match(stripped_line) or not stripped_line:
            continue

        # 去掉行内注释，保留有效代码
        line_without_comment = re.sub(r'#.*$', '', line).rstrip()

        # 如果去掉注释后仍有内容，加入结果列表
        if line_without_comment.strip():
            processed_lines.append(line_without_comment + '\n')

    # 生成新的文件名
    new_file_name = f"new_{file_name}.py"

    # 写入无注释和无空行的内容到新文件
    with open(new_file_name, 'w', encoding='utf-8') as new_file:
        new_file.writelines(processed_lines)

    print(f"已成功生成无注释和空行的文件：{new_file_name}")


# 调用函数
remove_comments_and_empty_lines_from_py()
