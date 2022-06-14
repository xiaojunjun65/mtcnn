# -*- encoding: utf-8 -*-
import os
from win32com import client


# pip install win32com
def doc2pdf(doc_name, pdf_name):
    """
    :word文件转pdf
    :param doc_name word文件名称
    :param pdf_name 转换后pdf文件名称
    """
    try:
        word = client.DispatchEx("Word.Application")
        
        worddoc = word.Documents.Open(doc_name, ReadOnly=1)
        worddoc.SaveAs(pdf_name, FileFormat=17)
        worddoc.Close()
        return pdf_name
    except:
        return 1


def main():
    input = r'C:\Users\zhou\Desktop\11\周国军.docx'
    print(input)


output = r'C:\Users\zhou\Desktop\周国军.pdf'
print(output)
rc = doc2pdf(input, output)
print(rc)
# rc = doc2html(input, output)
# rc = pdf2doc(input, output)
if rc:
    print('转换成功')
else:
    print('转换失败')

if __name__ == '__main__':
    main()
