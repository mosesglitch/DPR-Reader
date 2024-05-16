# import xlrd
from flask import Flask,request,jsonify
import openpyxl
import pandas as pd
print("helo world")

# dataframe = openpyxl.load_workbook("DPRSample.xlsx")
# book = xlrd.open_workbook('DPRSample.xlsx')
app=Flask(__name__)

@app.route("/getissues",methods=["POST"])
def extractissues():
    file = request.files['file']
    df = pd.read_excel(file) 
    row_data = df.iloc[:, 0].tolist()
    technical_idx = row_data.index('Technical Issues')
    material_idx = row_data.index('Material Issues')
    non_empty_issues = [value for value in row_data[technical_idx+1:material_idx] if value]

    #fetch days activities
    const_id = df[df.iloc[:, 0] == 'CONSTRUCTION ACTIVITIES'].index[0]
    health_id = df[df.iloc[:, 0] == 'HEALTH, ENVIRONMENTAL,SAFETY AND SECURITY ACTIVITIES'].index[0]
    selected_rows = df.iloc[const_id:health_id] 
    selected_rows_trimmed = selected_rows.iloc[:, 5:-3]
    df_filtered = selected_rows_trimmed.dropna(subset=[selected_rows_trimmed.columns[1]])
    # print("Row ID:", df_filtered)
    column_names = ['Today Activities', 'Planned',"Achieved Today","Cumulative","Total scope","Remarks" ]
    df_filtered.columns = column_names
    df_filtered[['Achieved Today', 'Planned']] = df_filtered[['Achieved Today', 'Planned']].apply(pd.to_numeric, errors='coerce')
    df_filtered['Percentage Complete'] = (df_filtered['Achieved Today'] / df_filtered['Planned']) * 100
    data_list = df_filtered.to_dict(orient='records')
    return jsonify({'issues': non_empty_issues, 'daily_activities':data_list})

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
