# import os
# arr=os.listdir('.')
# for file in arr:
#         if file.endswith(".xlsx"):
#             print(file)

# import glob
# for file in glob.glob('*.xlsx'):
#     print(file)

import glob
for file in glob.glob('*.xlsx'):
    b=[]
    print(file)
    b.append(file)
    #b.extend(file)
    

print(b)