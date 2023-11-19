# HERE IS THE ONE FILE COMPILER FOR YOUR APP
import os

os.chdir("C:/GitHub/shari_genetics/app_v1_0")

# build the executable
os.system("pyinstaller --onefile runme.py")

# build the /reports folder
os.mkdir("dist/reports")

# copy the source and the destination folders
os.system("copy random_values.xlsx dist\\random_values.xlsx")
