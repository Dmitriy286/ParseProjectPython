import parse_docx
import powershell_init
import ps_open_file_script
from urllib.parse import unquote
import re



if __name__ == '__main__':
    parse_docx.parse_one_docx()
    # powershell_init.simple_command()
    # powershell_init.run_script_from_file()
    # ps_open_file_script.open_file()
    # powershell_init.run_script_from_test_file()

    # print(unquote('Р”РѕРєСѓРјРµРЅС‚3.docx', 'cp1251'))
    # print(unquote('Р”РѕРєСѓРјРµРЅС‚3.docx', 'utf-8'))
    # print(unquote('Р”РѕРєСѓРјРµРЅС‚3.docx', 'unicode'))
    # print(unquote('https://csit.sharepoint.com/sites/99999999/_layouts/15/Doc.aspx?sourcedoc=%7B8277E30D-6E66-4B64-A0AC-611CE612E893%7D&file=%D0%94%D0%BE%D0%BA%D1%83%D0%BC%D0%B5%D0%BD%D1%823.docx&action=default&mobileredirect=true'))
    # pattern = r'file=\w+.docx'
    # string = unquote('https://csit.sharepoint.com/sites/99999999/_layouts/15/Doc.aspx?sourcedoc=%7B8277E30D-6E66-4B64-A0AC-611CE612E893%7D&file=%D0%94%D0%BE%D0%BA%D1%83%D0%BC%D0%B5%D0%BD%D1%823.docx&action=default&mobileredirect=true')
    # search = re.search(pattern, string)
    # print(search[0].removesuffix(".docx").removeprefix("file="))

    # parse_docx.make_data_frame()
    #
    # s = "techiedelight.com"
    # ch = 'coms'
    #
    # if ch in s:
    #     print("Character found")
    # else:
    #     print("Character not found")
