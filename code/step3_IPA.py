import xlrd
import numpy as np
import matplotlib.pyplot as plt
import os

wb = xlrd.open_workbook(f"E:/MOOC/MOOC_Improvement_Based_on_Aspect-based_Sentiment_Analysis_Driven_by_Large_Language_Models/results/scores_and_attentions.xls")
sheet_r_scores = wb.sheet_by_name('scores')
sheet_r_attention = wb.sheet_by_name('attention')

rows = sheet_r_scores.nrows
columns = sheet_r_scores.ncols
courses = [sheet_r_scores.cell(i, 0).value for i in range(1, rows)]
aspects = [sheet_r_scores.cell(0, i).value for i in range(1, columns)]

save_folder = "E:/MOOC/MOOC_Improvement_Based_on_Aspect-based_Sentiment_Analysis_Driven_by_Large_Language_Models/results"

scores_list = []
attention_list = []

for i in range(1, rows):
    scores = []
    attention = []
    for j in range(1, columns):
        scores.append(float(sheet_r_scores.cell(i, j).value))
        attention.append(float(sheet_r_attention.cell(i, j).value))
    scores_list.append(scores)
    attention_list.append(attention)

all_scores = [element for row in scores_list for element in row]
all_attention = [element for row in attention_list for element in row]
mean_x = sum(all_scores)/len(all_scores)
mean_y = sum(all_attention)/len(all_attention)
print(mean_x)
print(mean_y)
for i in range(len(scores_list)):

    # 创建图形
    plt.figure(figsize=(8, 6))
    # 绘制散点图
    plt.scatter(scores_list[i], attention_list[i], color='blue', label='Attributes')
    for j, name in enumerate(aspects):
        plt.text(scores_list[i][j], attention_list[i][j], name, fontsize=10, ha='left', va='bottom')

    # 添加均值线
    plt.axhline(mean_y, color='red', linestyle='--', label=f'Mean Importance = {mean_y:.2f}')
    plt.axvline(mean_x, color='green', linestyle='--', label=f'Mean Satisfaction = {mean_x:.2f}')

    # 添加象限标签
    # plt.text(mean_x + 1, mean_y + 1, 'Quadrant I\nKeep up the good work!', fontsize=12, ha='left', va='bottom')
    # plt.text(mean_x - 1, mean_y + 1, 'Quadrant II\nOver-prioritized!', fontsize=12, ha='right', va='bottom')
    # plt.text(mean_x - 1, mean_y - 1, 'Quadrant III\nLow priority!', fontsize=12, ha='right', va='top')
    # plt.text(mean_x + 1, mean_y - 1, 'Quadrant IV\nNeed improvement!', fontsize=12, ha='left', va='top')

    # 添加标题和标签
    plt.title(f'Importance-Performance Analysis (IPA) for {courses[i]}')
    plt.xlabel('Satisfaction')
    plt.ylabel('Importance')
    plt.legend()
    plt.grid(True)

    # 调整布局以避免标签重叠（如果需要）
    plt.tight_layout()

    # 保存图形到文件
    file_name = f'{courses[i]}_ipa.png'
    file_path = os.path.join(save_folder, file_name)
    plt.savefig(file_path)

    # 关闭当前图形以避免内存泄漏（当生成多个图形时很重要）
    plt.close()

