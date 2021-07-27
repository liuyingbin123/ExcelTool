"""
    author:         liuyingbin
    project_name:   ExcelTool
    data:           2021/7/27
    company:        SouthSmart
    CopyRight:      Doc LIU
"""
import argparse
import xlrd
"""
...主逻辑处理入口
"""


def main():
    parser = argparse.ArgumentParser(description='输出参数列表')
    parser.add_argument('-dataTablePath', default='./data/data.xlsx',type=str, help='基础数据表输入路径... '
                                                                        'for example:-dataTablePath test.xlsx')
    parser.add_argument('-outputPath', default='./test.xlsx', type=str, help='数据输出路径')
    args = parser.parse_args()
    # 创建结果表
    csvfile = file('csv_test.csv', 'wb')
    csvfile.write(codecs.BOM_UTF8)
    writer = csv.writer(csvfile)
    list = ['姓名', '年龄', '电话']
    writer.writerow(list)
    csvfile.close()

if __name__ == '__main__':
    main()
