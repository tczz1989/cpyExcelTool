import excelFunc as e
import readConfig as r


if __name__ == '__main__':
    input_config = r.get_input_config()
    output_config = r.get_output_config()
    print(input_config, output_config)
    r.get_input_filenames()
    workbook = e.open_excel(input_config[0][1])
    sheet = e.open_sheet(workbook)
    # e.read_excel(filename)
    rows = e.get_row_name(input_config[1][1])
    cols = e.get_col_name(input_config[2][1])
    rows1 = e.get_row_name(output_config[1][1])
    cols1 = e.get_col_name(output_config[2][1])
    print(rows, cols, rows1, cols1)
    print(e.input_output_match(rows, cols, rows1, cols1))

    data = e.read_excel(sheet, rows, cols)
    e.write_excel("Excel_test.xls", data, rows, cols)
