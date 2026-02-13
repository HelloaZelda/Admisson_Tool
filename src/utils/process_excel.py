import xlrd
import xlwt
from collections import defaultdict

def read_excel(file_path):
    """读取Excel文件并返回数据"""
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    
    # 获取表头
    headers = [sheet.cell_value(0, i) for i in range(sheet.ncols)]
    
    # 获取数据
    data = []
    for row in range(1, sheet.nrows):
        row_data = [sheet.cell_value(row, i) for i in range(sheet.ncols)]
        data.append(dict(zip(headers, row_data)))
    
    return data

def process_data(data):
    """处理数据，统计每个专业的选择情况"""
    major_stats = defaultdict(lambda: {'first': 0, 'second': 0, 'third': 0})
    
    for student in data:
        # 统计第一志愿
        major_stats[student['第一志愿']]['first'] += 1
        # 统计第二志愿
        major_stats[student['第二志愿']]['second'] += 1
        # 统计第三志愿
        major_stats[student['第三志愿']]['third'] += 1
    
    return major_stats

def write_results(stats, output_file):
    """将统计结果写入新的Excel文件"""
    wb = xlwt.Workbook()
    ws = wb.add_sheet('专业统计')
    
    # 写入表头
    headers = ['专业', '第一志愿人数', '第二志愿人数', '第三志愿人数', '总人数']
    for i, header in enumerate(headers):
        ws.write(0, i, header)
    
    # 写入数据
    for row, (major, counts) in enumerate(stats.items(), 1):
        total = sum(counts.values())
        ws.write(row, 0, major)
        ws.write(row, 1, counts['first'])
        ws.write(row, 2, counts['second'])
        ws.write(row, 3, counts['third'])
        ws.write(row, 4, total)
    
    wb.save(output_file)

def main():
    input_file = 'data/example_students.xls'
    output_file = 'data/major_statistics.xls'
    
    # 读取数据
    data = read_excel(input_file)
    
    # 处理数据
    stats = process_data(data)
    
    # 写入结果
    write_results(stats, output_file)
    
    print(f"处理完成！结果已保存到 {output_file}")

if __name__ == '__main__':
    main() 