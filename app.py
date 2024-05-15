# import xlrd
from flask import Flask
import openpyxl
import pandas as pd
print("helo world")

# dataframe = openpyxl.load_workbook("DPRSample.xlsx")
# book = xlrd.open_workbook('DPRSample.xlsx')
app=Flask(__name__)

@app.route("/getissues",methods=["POST"])
def extractissues():
    file = request.files['file']
    # save file in local directory
    # file.save(file.filename)
    df = pd.read_excel(file) 
    print(df)
    row_data = df.iloc[:, 0].tolist()
    # print(row_data)
    technical_idx = row_data.index('Technical Issues')
    material_idx = row_data.index('Material Issues')

    # start_idx = row_data.find('technical issues') + len('technical issues') + 1
    # end_idx = row_data.find('material issues')
    # values_between = row_data[start_idx:end_idx].strip()
    non_empty_values = [value for value in row_data[technical_idx+1:material_idx] if value]
    return non_empty_values
    print(non_empty_values)
    print(technical_idx,material_idx)
# for name in book.sheet_names():
#     if name.endswith('2'):
#         sheet = book.sheet_by_name(name)

#         # Attempt to find a matching row (search the first column for 'john')
#         rowIndex = -1
#         for cell in sheet.col(0): # 
#             if 'john' in cell.value:
#                 break

        # # If we found the row, print it
        # if row != -1:
        #     cells = sheet.row(row)
        #     for cell in cells:
        #         print cell.value

        # book.unload_sheet(name)
if __name__=="__main__":
    app.run()
