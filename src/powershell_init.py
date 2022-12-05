import os
import subprocess, sys

def simple_command():
    os.system("powershell Get-Process")
    os.system("powershell Write-Host 'Hello!'")


def run_script_from_file():
    ps = r"C:\Users\Professional\Desktop\Programming\Get-Process.ps1"
    p = subprocess.Popen(["powershell.exe",
                          ps],
                         stdout=sys.stdout)
    p.communicate()

def run_script_from_test_file():
    siteUrl = "www.o.ru"
    fileName = "Документ.docx"
    ps = r"C:\Users\Professional\Desktop\Programming\foo.ps1 " + siteUrl + " " + fileName
    p = subprocess.Popen(["powershell.exe",
                          ps],
                         stdout=sys.stdout)
    p.communicate()





    # ps ="C:\Users\Professional\Desktop\Programming\Get-Process.ps1"
    #
    #
    # System.Diagnostics.Process.Start( @ "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe", " -executionpolicy Unrestricted -File " + ps);
    #
    #
    # subprocess.Popen("powershell Get-Process", shell=True)
    # print(os.system("ipconfig"))