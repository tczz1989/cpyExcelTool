import excelFunc as e
import readConfig as r
import sys


if __name__ == '__main__':
    try:
        input_config, output_config = r.get_config()
        files = r.get_output_filenames()
        if output_config[0][1] not in files:
            e.create_excel_xlsx(output_config)
        e.cpy_excel_main(input_config, output_config)
        print("对excel的操作成功!\n")
        sys.exit(0)
    except Exception as e:
        print(e)
        print("对excel的操作失败，出现错误！！！！\n")
        sys.exit(-1)
