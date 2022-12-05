import os
import subprocess, sys

def open_file():
    command = r"""
    cd 'C:\Users\Professional\Desktop\Programming'
    'NewDoc.docx'
    """
    # command2 = r"""
    # start %windir%\explorer.exe 'C:\Users\Professional\Desktop\Programming',
    #
    # start %windir%\explorer.exe 'C:\Users\Professional\Desktop\Картотека\Алгоритмы и структуры данных'
    # """
    command2 = r"start %windir%\explorer.exe 'C:\Users\Professional\Desktop\Programming', " \
               r"start %windir%\explorer.exe 'C:\Users\Professional\Desktop\Картотека\Алгоритмы и структуры данных'"

    # os.system(r"powershell cd 'C:\Users\Professional\Desktop\Programming'")
    # os.system(r"powershell Get-Content -Path 'C:\Users\Professional\Desktop\Programming\NewDoc.docx'")
    # os.system("powershell " + command)
    # os.system("powershell " + command2)

    p = subprocess.Popen(["powershell.exe",
                          command2]
                         )
    p.communicate()
