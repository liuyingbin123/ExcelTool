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
import pandas as pd
import copy
STATICS_LAND_ACQUISITION_AREA = ["水库淹没影响区", "枢纽工程建设区"]
STATICS_FUNCTIONAL_AREA1 = ["淹没区","浸没区","塌岸区","滑坡区","内涝区","水库渗透区", "其他"]
STATICS_FUNCTIONAL_AREA2 = ["库区提前征用","永久","临时"]


def generate_result_table(output_path):
    """
    :生成目标表结构
    :param output_path: 输出路径
    :return: 目标表
    """
    # 创建结果表
    result = ResultTable(output_path)
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
    result.setData(4, 9, '塌岸区')
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
    result.setData(4, 22, '塌岸区')
    result.setData(4, 23, '滑坡区')
    result.setData(4, 24, '内涝区')
    result.setData(4, 25, '水库渗透区')
    result.setData(4, 26, '其他')
    result.setData(3, 27, '枢纽工程建设区', 3, 27, 3, 30)
    result.setData(4, 27, '小计')
    result.setData(4, 28, '库区提前征用')
    result.setData(4, 29, '永久')
    result.setData(4, 30, '临时')
    # 生成数据表区域
    start_row = 5
    end_row = 173
    start_col = 1
    end_col = 31
    for row in range(start_row, end_row):
        for col in range(start_col, end_col):
            # 占位
            result.setData(row, col, '')
    # 根据json配置行
    with open('./config/LandUseType.json', 'r', encoding='utf8')as jsonfile:
        json_data = json.load(jsonfile)
    start = 5
    for i, key in enumerate(json_data):
        data = json_data[key]
        result.setData(start, 2, key)
        start += 1
        if type(data) == list:
            for j in range(len(data)):
                subdata = data[j]
                if type(subdata) == dict:
                    for z, subkey in enumerate(subdata):
                        landuse_data = subdata[subkey]
                        for k in range(len(landuse_data)):
                            landuse_group = landuse_data[k]
                            for landuse_group_key in landuse_group.keys():
                                result.setData(start, 1, landuse_group_key)
                                landuse_types = landuse_group[landuse_group_key]
                                for landuse_type in landuse_types:
                                    result.setData(start, 2, landuse_type)
                                    start += 1

    return result


def get_landuse_price(input_path):
    """
    读取单价表
    :param input_path:输入路径
    :return:单价字典
    """
    with open(input_path, 'r', encoding='utf8')as jsonfile:
        json_data = json.load(jsonfile)
    return json_data


def get_data_statics(df):
    """
    0:OBJECTID *
    1:省
    2:市
    3:区
    4:街道办
    5:社区
    6:征地区域
    7:用地性质
    8:功能区1
    9:功能区2
    10:区类
    11:三大类
    12:一级类
    13:二级类
    14:面积_亩
    15:地类备注1
    16:地类备注2
    17:地类备注3
    18:边长
    19:面积
    """
    with open("./config/LandUseArea.json", 'r', encoding='utf8')as json_file:
        land_use_area_statics = json.load(json_file)
    temps1 = {}
    temps2 = {}
    statics = {}
    for i in range(len(STATICS_FUNCTIONAL_AREA1)):
        temps1[STATICS_FUNCTIONAL_AREA1[i]] = copy.deepcopy(land_use_area_statics)
    for i in range(len(STATICS_FUNCTIONAL_AREA2)):
        temps2[STATICS_FUNCTIONAL_AREA2[i]] = copy.deepcopy(land_use_area_statics)
    statics["水库淹没影响区"] = temps1
    statics["枢纽工程建设区"] = temps2
    # 获取当前表的数据行列
    rows = df.shape[0]
    for row in range(rows):
        row_data = df.loc[row].values
        # 征地区域
        if row_data[6] in STATICS_LAND_ACQUISITION_AREA and row_data[6]=="水库淹没影响区":
            try:
                # 用地类型, 只在json指明的地类中进行统计
                if row_data[8] in STATICS_FUNCTIONAL_AREA1:
                    statics[row_data[6]][row_data[8]][row_data[12]][row_data[13]] += row_data[19]
                else:
                    statics[row_data[6]]["其他"][row_data[12]][row_data[13]] += row_data[19]
            except:
                pass
        elif row_data[6] in STATICS_LAND_ACQUISITION_AREA and row_data[6]=="枢纽工程建设区":
            try:
                # 用地类型, 只在json指明的地类中进行统计
                if row_data[7] in STATICS_FUNCTIONAL_AREA2:
                    statics[row_data[6]][row_data[7]][row_data[12]][row_data[13]] += row_data[19]
            except:
                pass
    return statics


def fill_data(statics, table, price):
    # 填充统计字典的每一列, 从第6行第7列开始
    start_col = 7
    for i, key in enumerate(statics):
        first = statics[key]
        for j, key2 in enumerate(first):
            start_row = 7
            second = first[key2]
            for z, key3 in enumerate(second):
                thrid = second[key3]
                for y, key4 in enumerate(thrid):
                    table.setData(start_row,start_col,thrid[key4])
                    start_row += 1
                start_row += 1
            start_col += 1
        start_col += 1
    second_row = start_row + 1
    start_col = 7
    for i, key in enumerate(statics):
        first = statics[key]
        for j, key2 in enumerate(first):
            start_row = second_row
            second = first[key2]
            for z, key3 in enumerate(second):
                thrid = second[key3]
                for y, key4 in enumerate(thrid):
                    table.setData(start_row,start_col,thrid[key4])
                    start_row += 1
                start_row += 1
            start_col += 1
        start_col += 1



def main():
    """
    ...主逻辑处理入口
    """
    parser = argparse.ArgumentParser(description='输出参数列表')
    parser.add_argument('-dataTablePath', default='./data/data.xlsx',type=str, help='基础数据表输入路径... '
                                                                        'for example:-dataTablePath test.xlsx')
    parser.add_argument('-outputPath', default='./test.xlsx', type=str, help='数据输出路径')
    args = parser.parse_args()
    result = generate_result_table(args.outputPath)
    land_use_price = get_landuse_price("./config/LandUnitPrice.json")
    # 读取基础数据
    # 一行数据的结构如下
    df = pd.read_excel("./data/data.xlsx")
    # 统计数据
    statics = get_data_statics(df)
    with open("./config/LandUnitPrice.json", 'r', encoding='utf8')as json_file:
        land_use_price = json.load(json_file)
    # 填充数据到结果表
    fill_data(statics, result, land_use_price)
    result.save()


if __name__ == '__main__':
    main()
