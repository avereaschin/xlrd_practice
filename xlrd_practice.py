import xlrd

path = '2013_ERCOT_Hourly_Load_Data.xls'

# load excel workbook
workbook = xlrd.open_workbook(path)

# load first sheet from the workbook
sheet = workbook.sheet_by_index(0)

# load all data into a list of row lists using a listcomp
data = [[sheet.cell_value(r, col)
            for col in range(sheet.ncols)]
                for r in range(sheet.nrows)]

# returns an int indicating the type of data within the cell. See link for more info: 
# https://pythonhosted.org/xlrd3/cell.html
sheet.cell_type(1, 2)

# retrieves all rows from column 1
data_ = sheet.col_values(1, start_rowx=1, end_rowx=None)

# running basic data analysis (retrieving max, min and avg values from the column)
max_ = max(data_)
# excel stores dates as floats so to convert them into a python tuple we use
# the xlrd.xldate_as_tuple function which can then be converted into a datetime.strftime
# string
max_time = xlrd.xldate_as_tuple(sheet.cell_value(data_.index(max_) + 1, 0), workbook.datemode)

min_ = min(data_)
min_time = xlrd.xldate_as_tuple(sheet.cell_value(data_.index(min_) + 1, 0), workbook.datemode)

avg_ = sum(data_)/len(data_)

# store results in a dict
result = {
            'maxtime': max_time,
            'maxvalue': max_,
            'mintime': min_time,
            'minvalue': min_,
            'avgvalue': avg_
}
    
print(result)