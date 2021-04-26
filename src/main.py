import excelFunc as e
import readConfig as r


if __name__ == '__main__':
    input_config = r.get_input_config()
    output_config = r.get_output_config()
    print(input_config, output_config)
    r.get_input_filenames()
    e.cpy_excel_main(input_config, output_config)
