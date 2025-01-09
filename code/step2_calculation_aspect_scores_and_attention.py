import xlrd
import xlwt


# scores_all = []
# attention_all = []
aspects = []
book = xlwt.Workbook(encoding='utf-8')
sheet_w_s = book.add_sheet('scores', cell_overwrite_ok=True)
sheet_w_a = book.add_sheet('attention', cell_overwrite_ok=True)
sheet_w_s.write(0, 0, '课程')
sheet_w_a.write(0, 0, '课程')
row_w = 1
for i in range(1, 6):
    dataname = f'人工智能{i}'
    sheet_w_s.write(row_w, 0, dataname)
    sheet_w_a.write(row_w, 0, dataname)
    wb = xlrd.open_workbook(f"E:/MOOC/MOOC_Improvement_Based_on_Aspect-based_Sentiment_Analysis_Driven_by_Large_Language_Models/results/{dataname}.xls")
    sheet_r = wb.sheet_by_index(0)
    rows = sheet_r.nrows
    positive_list = []
    negative_list = []
    for j in range(2, rows):
        if i == 1:
            aspects.append(sheet_r.cell(j, 1).value)
        positive_list.append(int(sheet_r.cell(j, 2).value))
        negative_list.append(int(sheet_r.cell(j, 3).value))


    scores = [positive_list[k]/(positive_list[k]+negative_list[k]) for k in range(len(positive_list))]
    attention = [(positive_list[k]+negative_list[k])/sum(positive_list + negative_list) for k in range(len(positive_list))]
    for j in range(len(scores)):
        if i == 1:
            sheet_w_s.write(0, j + 1, aspects[j])
            sheet_w_a.write(0, j + 1, aspects[j])
        sheet_w_s.write(row_w, j + 1, scores[j])
        sheet_w_a.write(row_w, j + 1, attention[j])
    row_w += 1
book.save(
    f"E:/MOOC/MOOC_Improvement_Based_on_Aspect-based_Sentiment_Analysis_Driven_by_Large_Language_Models/results/scores_and_attentions.xls")
    # scores_all.append(scores)
    # attention_all.append(attention)

