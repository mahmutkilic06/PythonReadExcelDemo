import pandas
import sys

excel_data_df = pandas.read_excel(r"""%s""" % sys.argv[1], sheet_name=sys.argv[2], dtype=str)
print(excel_data_df.to_json(orient='records', date_format='iso'))
