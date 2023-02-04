from cx_Freeze import setup, Executable


files = ["icon.ico", "themes/", "icons/"]

target = Executable(
    script="main.py",
    base="Win32GUI",
    icon="icons\icon.ico"   
)

setup(
    name="PyMailSender",
    version="1.0",
    description="PyGmailSender using SMTP Server",
    author = "MB",
    options = {'build_exe' : {'include_files' : files }},
    executables=[target]
)