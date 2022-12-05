from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from urllib.parse import unquote
import re
import pandas as pd
import subprocess, sys


def parse_one_docx():

    #script load

    # docx_file = r"C:\Temp5\Document1"
    docx_file = r"C:\Users\Professional\Desktop\Programming\Parse project\FOR_PATHS_1\Документ"

    document = Document(docx_file + ".docx")

    rels = document.part.rels.items()

    count = 1
    for key, value in rels:
        # part = paragraph.part
        # r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        # run = document.text.run.Run

        if value.reltype == RT.HYPERLINK and "sharepoint.com" in str(value.target_ref):
            print("\n Origianl link id -", key)
            print(f"URL № {count}: {value.target_ref}")
            print(value.reltype)

            count += 1

            #здесь должен появляться результат выполнения скрипта - новая ссылка
            #скрипт сохранен под именем "MakeAbsoluteUrl"
            #вызываем, передаем параметры: адрес сайта, подсайт (listname), название файла

        # ==========================

            siteUrlLink = "http://stpa01z502/sites/TestOne"
            driveListname, filename = make_data_frame(value.target_ref)

            # раскомментировать на локальном компьютере:
            ps = r"C:\Users\Professional\Desktop\Programming\foo.ps1 " + siteUrlLink + " " + driveListname + " " + filename

            #ps = r"C:\Temp5\MakeAbsoluteUrl.ps1 " + siteUrlLink + " " + driveListname + " " + filename
            p = subprocess.Popen(["powershell.exe",
                                  ps],
                                 stdout=subprocess.PIPE)

            try:
                outs, errs = p.communicate(timeout=15)
            except subprocess.TimeoutExpired:
                p.kill()
                outs, errs = p.communicate()

            print("Возврат из скрипта:")
            print(outs)
            print(type(outs))

            # string = eval(outs)
            string = outs.decode('Windows-1251')
            print(type(string))
            print(string)

            pattern = r'result: .+ end'
            # pattern = re.compile('result: \.+ end')
            # string = outs.decode("utf-8")

            search = re.search(pattern, string)
            result_from_script = search[0].removeprefix("result: ").removesuffix(" end")

            print(result_from_script)

            # ==========================

            #убрать хардкод:
            new_url=result_from_script
            document.part.rels.add_relationship(value.reltype, new_url, value.rId, value.is_external)

    out_file=docx_file + "-out.docx"

    document.save(out_file)

    print("\n File saved to: ", out_file)

    #script update

def make_data_frame(link):
    df = pd.read_csv('99999999.csv', delimiter=';')

    data = df[['site', 'drive', 'path', 'weburl', 'link']]
    print(data)
    df['number'] = df['site'].index
    print(df)

    link_dict = dict(zip(df['number'], df['link']))

    pattern = None
    pattern_for_search = None
    filename = None
    row_index_for_collecting_data = None

    drive_dict = dict(zip(df['number'], df['drive']))

    if ".docx" in str(link):
        pattern = r'\?d=\w{4}'
        string = unquote(link)
        search = re.search(pattern, string)
        pattern_for_search = search[0].removeprefix("?d=")
        isNormalLink = True
        row_index_for_collecting_data = 0
        filename = take_filename_from_url(link, df, row_index_for_collecting_data, isNormalLink)
        powershell_listname = re.search(r'999/\w+/', link)[0].removesuffix("/").removeprefix("999/")
    else:
        pattern = r'999/\w{4}'
        # string = unquote(link)
        search = re.search(pattern, link)
        print(link)
        print(search)
        pattern_for_search = search[0].removeprefix("999/")
        for index, url in link_dict.items():
            if pattern_for_search in str(url):
                row_index_for_collecting_data = int(index)
        isNormalLink = False
        filename = take_filename_from_url(link, df, row_index_for_collecting_data, isNormalLink)
        powershell_listname = drive_dict[row_index_for_collecting_data]

    print("pattern_for_search:")
    print(pattern_for_search)
    print("filename:")
    print(filename)
    print("powershell_listname:")
    print(powershell_listname)

    return powershell_listname, filename

def take_filename_from_url(link, df, row_index_for_collecting_data, isNormalLink):
    filename = None

    weburl_dict = dict(zip(df['number'], df['weburl']))

    if isNormalLink:
        pattern = r'/%.+.docx'
        string = link
        search = re.search(pattern, string)
        encoded_filename = search[0].removesuffix(".docx").removeprefix("/")
        filename = unquote(encoded_filename)

    else:
        link_in_sp_with_filename = weburl_dict[row_index_for_collecting_data]
        print(link_in_sp_with_filename)
        pattern = r'file=\w+.docx'
        string = unquote(link_in_sp_with_filename)
        search = re.search(pattern, string)

        filename = search[0].removesuffix(".docx").removeprefix("file=")

    #todo доделать чтобы из столбца path вытаскивал подсайты после "root:": root:/InnerTest/MoreInnerTest

    return filename


if __name__ == '__main__':
    parse_one_docx()