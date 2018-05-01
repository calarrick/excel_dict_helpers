import excel_dict
# import json
import pickle
import jsondate
# import jsonpickle

rows = excel_dict.read_sheet('Research Sessions.xlsx', 'Research Sessions', start_row=5)
with open('data_workfile.bin','wb') as file:
    pickle.dump(rows, file)

# with open('data_workfile.bin', 'rb') as file2:
#     rows2 = pickle.load(file2)
# print(rows2)

with open('data_workfile.txt', 'w') as file3:
    jsondate.dump(rows, file3)