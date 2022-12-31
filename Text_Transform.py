import comtypes.client
from pdf2docx import Converter
import PySimpleGUI as sg
import pdfkit
from markdown import markdown
import os
import xlrd
import datetime
from mailmerge import MailMerge


# pdf转word
def pdf2word(file_path):
    file_name = file_path.split('.')[0]
    doc_file = f'{file_name}.docx'
    p2w = Converter(file_path)
    p2w.convert(doc_file, start=0, end=None)
    p2w.close()
    return doc_file


# word转pdf
def word2pdf(file_path):
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = 0
    file_name = file_path.split('.')[0]
    pdf_file = f'{file_name}.pdf'
    w2p = word.Documents.Open(file_path)
    w2p.SaveAs(pdf_file, FileFormat=17)
    w2p.Close()
    return pdf_file


# markdown转pdf
def md2pdf(file_path):
    file_name = file_path.split('\\')[-1].split('.m')[0]
    _pdf_file = f'{file_name}.pdf'

    with open(file_path, encoding='utf-8') as f:
        text = f.read()

    html = markdown(text, output_format='html', extensions=['tables'])  # MarkDown转HTML

    htmltopdf = r'D:\software\wkhtmltopdf\bin\wkhtmltopdf.exe'
    configuration = pdfkit.configuration(wkhtmltopdf=htmltopdf)
    pdfkit.from_string(html, output_path=_pdf_file, configuration=configuration,
                       options={'encoding': 'utf-8', 'enable-local-file-access': None})  # HTML转PDF

    return _pdf_file


# excel转word
def excel2word(file_path):
    file = xlrd.open_workbook(file_path)
    table = file.sheet_by_name(file.sheet_names()[0])
    nrows = table.nrows
    path_base_name = r'C:\Users\cherich\Desktop'
    word_li = []
    for i in range(nrows):
        if i > 0:
            # 打开固定模板文件
            template_path = path_base_name + '//' + 'demo.docx'
            doc = MailMerge(template_path)
            print(doc.get_merge_fields())
            username = str(table.row_values(i)[1]).strip()
            linkuser = str(table.row_values(i)[2]).strip()
            phone = str(table.row_values(i)[3])
            address = str(table.row_values(i)[4])
            productname = str(table.row_values(i)[5]).strip()
            sku = str(table.row_values(i)[6]).strip()
            num = str(table.row_values(i)[7]).strip()
            money = str(table.row_values(i)[8])

            # 以下为填充模板中对应的域，
            doc.merge(
                username=username,
                linkuser=linkuser,
                address=address,
                linkphone=phone,
                productname=productname,
                num=num,
                sku=sku,
                money=money
            )
            path_name = os.path.join(path_base_name, datetime.datetime.now().strftime("%Y-%m-%d"))
            if not os.path.exists(path_name):
                os.makedirs(path_name)
            word_name = path_name + "\\" + username + '.docx'

            doc.write(word_name)
            doc.close()
            word_li.append(word_name)
    return word_li


def main():
    # 选择主题
    sg.theme('LightBlue5')

    layout = [
        [sg.Text('文本互转小工具', font=('微软雅黑', 12)),
         sg.Text('', key='filename', size=(50, 1), font=('微软雅黑', 10), text_color='blue')],
        [sg.Output(size=(100, 10), font=('微软雅黑', 10))],
        [sg.FilesBrowse('选择文件', key='file', target='filename'), sg.Button('pdf转word'), sg.Button('word转pdf'),
         sg.Button('markdown转pdf'), sg.Button('excel转word'),
         sg.Button('退出')]]
    # 创建窗口
    window = sg.Window("Python与数据分析_青青", layout, font=("微软雅黑", 15), default_element_size=(50, 1))
    # 事件循环
    while True:
        # 窗口的读取，有两个返回值（1.事件；2.值）
        event, values = window.read()
        print(event, values)

        if event == 'pdf转word':
            if values['file'] and values['file'].split('.')[2] == 'pdf':
                pdf_filename = pdf2word(values['file'])
                print('pdf文件个数 ：1')
                print('\n' + '转换成功！' + '\n')
                print('文件保存位置：', pdf_filename)
            elif values['file'] and values['file'].split(';')[0].split('.')[2] == 'pdf':
                print('pdf文件个数 ：{}'.format(len(values['file'].split(';'))))
                for f in values['file'].split(';'):
                    pdf_filename = pdf2word(f)
                    print('\n' + '转换成功！' + '\n')
                    print('文件保存位置：', pdf_filename)
            else:
                print('请选择pdf格式的文件哦!')
        if event == 'word转pdf':
            # excel转word {'file': 'C:/Users/cherich/Desktop/合同/demo.xlsx;C:/Users/cherich/Desktop/合同/demo1.xlsx'}
            if values['file'] and values['file'].split('.')[1] == 'docx':
                word_filename = word2pdf(values['file'])
                print('word文件个数 ：1')
                print('\n' + '转换成功！' + '\n')
                print('文件保存位置：', word_filename)
            elif values['file'] and values['file'].split(';')[0].split('.')[1] == 'docx':
                print('word文件个数 ：{}'.format(len(values['file'].split(';'))))
                for f in values['file'].split(';'):
                    filename = word2pdf(f)
                    print('\n' + '转换成功！' + '\n')
                    print('文件保存位置：', filename)
            else:
                print('请选择docx格式的文件哦!')
        if event == 'markdown转pdf':
            if values['file'] and values['file'].split()[1].split('.')[1] == 'md':
                md_filename = md2pdf(values['file'])
                print('markdown文件个数 ：1')
                print('\n' + '转换成功！' + '\n')
                print('文件保存位置：', md_filename)
            elif values['file'] and values['file'].split(';')[0].split()[1].split('.')[1] == 'md':
                print('markdown文件个数 ：{}'.format(len(values['file'].split(';'))))
                for f in values['file'].split(';'):
                    _filename = md2pdf(f)
                    print('\n' + '转换成功！' + '\n')
                    print('文件保存位置：', _filename)
            else:
                print('请选择md格式的文件哦!')
        if event == 'excel转word':
            if values['file'] and values['file'].split('.')[1] == 'xlsx':
                word_filename = excel2word(values['file'])
                print('excel文件个数 ：1')
                print('\n' + '转换成功！' + '\n')
                print('文件保存位置：', word_filename)
            elif values['file'] and values['file'].split(';')[0].split('.')[1] == 'xlsx':
                print('excel文件个数 ：{}'.format(len(values['file'].split(';'))))
                for f in values['file'].split(';'):
                    word_filename = excel2word(f)
                    print('\n' + '转换成功！' + '\n')
                    print('文件保存位置：', word_filename)
            else:
                print('请选择xlsx格式的文件哦!')
        if event in (None, '退出'):
            break

    window.close()


if __name__ == '__main__':
    main()
