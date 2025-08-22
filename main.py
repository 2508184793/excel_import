import excel_service

file = r'C:\Users\cj\Desktop\temp\myTest\test.xlsx'

if __name__ == '__main__':
    try:
        # excel_service.load_excel(r'C:\Users\cj\Desktop\temp\myTest\test.xlsx')
        # excel_service.get_excel_info(r'C:\Users\cj\Desktop\temp\myTest\test.xlsx')
        print(excel_service.get_title_info(sheet_name='Sheet5', row_index=1,
                                           file_path=file))
    except Exception as e:
        print(f"加载Excel文件时发生错误: {e}")

