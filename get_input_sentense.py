from docx import Document
import re
import xml.etree.ElementTree as ET
import os
import win32com.client as win32

def convert_doc_to_docx(doc_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(doc_path + "x", FileFormat=16)  # 16 表示 .docx 格式
    doc.Close()
    word.Quit()


def docx_input_gb(file_path):
    l = {}
    str1 = ''
    str2 = ''
    doc = Document(file_path)
    flag = False
    for paragraph in doc.paragraphs:
        data = paragraph.text.replace(' ', '')
        # print(data)
        # print(paragraph.text)
        if not flag:
            # 匹配标题
            pattern = r'\d+\.\d+\.\d+\.\d+(.*)'
            match = re.search(pattern, data)
            if match:
                flag = True
                str1 = match.group(0)
        elif data.find('漏洞描述：') != -1:
            str2 = data[data.find('漏洞描述：') + len('漏洞描述：') :]
            l[str2]=str1
            flag = False
            print(str2+' '+str1)
    return l


def docx_input_zbg(file_path):
    l = []
    doc = Document(file_path)
    for table in doc.tables:
        flag = 0
        # print(table)
        for row in table.rows:  # 遍历每一行
            for cell in row.cells:  # 便利每个单元格
                if cell.text == '漏洞类别':
                    flag = 1
                if flag == 0:
                    continue
                txt = (cell.text.translate(str.maketrans('', '', '0123456789')))
                l.append(txt)
                # print(cell.text)
    return list(i for i in set(l) if i.strip())


def xml_imput_kubo(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    problem_summary = root.find('你的问题汇总')
    rules_totals = problem_summary.findall('ruleTotal')
    result = []
    for rule_total in rules_totals:
        rule_total_name = rule_total.get('规则集')

        for rule in rule_total.findall('rule'):
            rule_data = {
                "序号": rule.get("序号"),
                "名称": rule.get("名称"),
                "严重等级": rule.get("严重等级"),
                "数量": rule.get("数量")
            }
            # result.append(''.join(re.findall('[\u4e00-\u9fa5]',rule_data["名称"])))
            result.append(rule_data["名称"].split(' ', 1)[1])
            # print(rule_data["名称"].split(' ', 1)[1])
            # print(rule_data)
    # print(len(result))
    # print(result)
    return result


def docx_input_keda(file_path):
    l = []
    doc = Document(file_path)
    for table in doc.tables:
        flag = 0
        # print(table)
        for row in table.rows:  # 遍历每一行
            for cell in row.cells:  # 便利每个单元格
                if cell.text == '序号':
                    flag = 1
                if flag == 0:
                    continue
                txt = (cell.text.translate(str.maketrans('', '', '0123456789')))
                if txt != '重大' and txt != '严重' and txt != '次要' and txt != '序号':
                    l.append(txt)
                # print(cell.text)
    return list(i for i in set(l) if i.strip())


def output_txt(file_path, data):
    with open(file_path, 'w', encoding='utf-8') as f:
        for i in data:
            f.write(i + '\n')

if __name__ == '__main__':

    # gb提取数据集
    #dx = docx_input_gb('E:/codespace/GBT34943(1).docx')

    # zbg提取数据集
    # l = []
    # for i in range(1,501):
    #     dx = docx_input_zbg(f'E:/database/lrj-report/1/lrj{i}-综合审计报告.docx')
    #     l += dx
    #     # print(f'{i}:\n{dx}')
    # for i in range(501, 600):
    #     dx = docx_input_zbg(f'E:/database/lrj-report/2/lrj{i}-综合审计报告.docx')
    #     print(f'{i}:\n{dx}')
    #     l += dx
    #
    # with open('E:/codespace/report_text.txt', 'w', encoding='utf-8') as f:
    #     for i in list(set(l)):
    #         if i == '危险' or i == '高' or i == '中' or i == '低' or i == '漏洞数' or i == '漏洞类别':
    #             continue
    #         f.write(i + '\n')
    # print(list(set(l)))

    # 库博报告提取数据
    # report = []
    # xml_kubo_filepath = 'E:/database/库博项目报告'
    # xml_kubo_file = [ xml_kubo_filepath + '/' + i for i in os.listdir(xml_kubo_filepath) if i.find('.xml') != -1]
    # a = 0
    # for i in xml_kubo_file:
    #     report += xml_imput_kubo(i)
    #
    #
    # report = list(set(report))
    # output_txt(xml_kubo_filepath + '/kubo.txt', report)

    # 科大报告提取数据
    # convert_doc_to_docx('E:/database/科大项目报告/1000-报告/1000-检测报告-概述汇总.doc')
    file_path_keda = 'E:/database/科大项目报告/1000-报告/1000-检测报告-概述汇总.docx'
    report = docx_input_keda(file_path_keda)
    print(report)

