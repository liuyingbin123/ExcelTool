"""
    author:         liuyingbin
    project_name:   ExcelTool
    data:           2021/7/27
    company:        SouthSmart
    CopyRight:      Doc LIU
"""
import argparse
from utils.resultTable import ResultTable
import json
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
    result = ResultTable(args.outputPath)
    # 格式化输出
    # 设置表头, 对于当前表而言有30列
    result.setTableHead('XXX水库建设征地移民安置补偿投资估算表(征收土地补偿费和安置补偿费及征用土地补偿费)', 1, 1, 1, 30)
    # 设置列名
    result.setData(2, 1, '序号', 2, 1, 4, 1)
    result.setData(2, 2, '项目', 2, 2, 4, 2)
    result.setData(2, 3, '数量单位', 2, 3, 4, 3)
    result.setData(2, 4, '单价', 2, 4, 4, 4)
    result.setData(2, 5, '数量', 2, 5, 2, 17)
    result.setData(3, 5, '合计', 3, 5, 4, 5)
    result.setData(3, 6, '水库淹没影响区', 3, 6, 3, 13)
    result.setData(3, 14, '枢纽工程建设区', 3, 14, 3, 17)
    result.setData(4, 6, '小计')
    result.setData(4, 7, '淹没区')
    result.setData(4, 8, '浸没区')
    result.setData(4, 9, '坍岸区')
    result.setData(4, 10, '滑坡区')
    result.setData(4, 11, '内涝区')
    result.setData(4, 12, '水库渗透区')
    result.setData(4, 13, '其他')
    result.setData(4, 14, '小计')
    result.setData(4, 15, '库区提前征用')
    result.setData(4, 16, '永久')
    result.setData(4, 17, '临时')
    result.setData(2, 18, '费用（万元）', 2, 18, 2, 30)
    result.setData(3, 18, '合计', 3, 18, 4, 18)
    result.setData(3, 19, '水库淹没影响区', 3, 19, 3, 26)
    result.setData(4, 19, '小计')
    result.setData(4, 20, '淹没区')
    result.setData(4, 21, '浸没区')
    result.setData(4, 22, '坍岸区')
    result.setData(4, 23, '滑坡区')
    result.setData(4, 24, '内涝区')
    result.setData(4, 25, '水库渗透区')
    result.setData(4, 26, '其他')
    result.setData(3, 27, '枢纽工程建设区', 3, 27, 3, 30)
    result.setData(4, 27, '小计')
    result.setData(4, 28, '库区提前征用')
    result.setData(4, 29, '永久')
    result.setData(4, 30, '临时')
    # 根据json配置行
    with open('./config/LandUseType.json', 'r', encoding='utf8')as jsonfile:
        json_data = json.load(jsonfile)
    start = 5
    for i in range(len(json_data)):
        key = json_data
        result.setData(5, 1, )
        start += 1
        for j in range(len(json_data[i])):
            result.setData()
            start += 1
    result.save()


if __name__ == '__main__':
    main()
