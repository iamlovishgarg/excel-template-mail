from pandas import read_excel
import filenames

# opening template file
with open(filenames.TEMPLATE_FILENAME, "r") as f:
    template = f.read()

# opening excel file
data = read_excel(filenames.EXCEL_FILENAME)

# getting all column names
columns = data.columns.values

for column in columns:
    for i in data[column]:
        template = template.replace("{"+column+"}", str(i))

print(template)