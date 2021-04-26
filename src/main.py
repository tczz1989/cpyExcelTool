import excelFunc as e
import readConfig as r
import os


if __name__ == '__main__':
    input_config = r.get_input_config()
    output_config = r.get_output_config()
    print(input_config, output_config)
    r.get_input_filenames()
    files = r.get_output_filenames()
    if output_config[0][1] not in files:
        e.create_excel_xlsx(output_config)
    e.cpy_excel_main(input_config, output_config)
