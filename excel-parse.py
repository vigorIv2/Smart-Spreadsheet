import helper_functions
import pandas as pd

def find_row_num(df, row_name_string = 'Total Cash and Cash Equivalent'):
    for index, row in df.iterrows():
        for c in row:
            if isinstance(c, str) and c.strip() == row_name_string:
                return index
    return None

def find_col_num(df, col_value_string = '2023-11-30 00:00:00'):
    for index, row in df.iterrows():
        c_index = 0
        for c in row:
            if str(c) == col_value_string:
                return c_index
            c_index += 1
    return None

def get_cell_value(df, rn, cn):
    specific_row = df.iloc[rn]
    return specific_row[cn]


def print_value(xslx_file, row_name_string = 'Total Cash and Cash Equivalent', col_value_string = '2023-11-30 00:00:00'):

    # Load the Excel file into a DataFrame
    df = pd.read_excel(xslx_file)

    rn = find_row_num(df, row_name_string = row_name_string)

    cn = find_col_num(df, col_value_string = col_value_string)

    print("file = ", xslx_file, "Request : What is the Total Cash and Cash Equivalent of Nov. 2023? =>", get_cell_value(df, rn, cn))
    return df

df = print_value("./tests/example_0.xlsx", row_name_string = 'Total Cash and Cash Equivalent', col_value_string = '2023-11-30 00:00:00')

oct = get_cell_value(df, 88, 4)
print("oct=", oct)

df = print_value("./tests/example_1.xlsx", row_name_string = 'Total Cash and Cash Equivalent', col_value_string = '2023-11-30 00:00:00')

df = print_value("./tests/example_2.xlsx", row_name_string = 'Total Cash and Cash Equivalent', col_value_string = '2023-11-30 00:00:00')

