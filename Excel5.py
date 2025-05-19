# import xlwings as xw
#
#
# app = xw.App(visible=False,add_book=False)
# for i in range(1,6):
#     workbook3 = app.books.add()
#     workbook3.save(f'G:\\Python\\Project\\PythonExcelTest\\NewExcel{i}.xlsx')  # 保存excel
#     workbook3.close()
# app.quit()
#
# import openpyxl
# for i in range(1,6):
#     workbook4 = openpyxl.Workbook()
#     workbook4.save(f'G:\\Python\\Project\\PythonExcelTest\\NewExcelopenxl{i}.xlsx')

text = "This is a sample text for word frequency analysis. This is"
words = text.split()
print(words)
word_count = {}
for word in words:
    # print(word)
    if word in word_count:
        word_count[word] += 1
    else:
        word_count[word] = 1
print(word_count)