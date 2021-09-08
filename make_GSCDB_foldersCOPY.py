from datetime import date
import datetime
import os
import shutil
import pendulum
desktop_path = 'C:\\Users\kschwartz\Desktop'
GSCDB_folder_path = "C:\\Users\kschwartz\Desktop\GSCD"
batch_names = []
weekday = datetime.datetime.today().weekday()
#TODO- FINISH MODIFYING MAKE GSCDB FOLDERS SO IT CAN MAKE THE ARCHIVE FOR OLD EXCEL SHEETS.
# if weekday == 0:
#     weekday == monday
#
# today = date.today()
# today = today.strftime('%m-%d-%y')
for i in range(1,6):
    batchnameI = "_batch_" + str(i)
    batch_names.append(batchnameI)

def make_todays_folders():
    today_folder_names = []
    today = date.today()
    today = today.strftime('%m-%d-%y')
    for i in batch_names:
        foldernameI = today + i
        today_folder_names.append(foldernameI)
    print(f"new folder names are: \n {today_folder_names}")
    for i in today_folder_names:
        os.mkdir(os.path.join(desktop_path,i))
    return today
def move_yesterdays_folders():
    yesterday_folder_names = []
    yesterday = date.today() - datetime.timedelta(days=1)
    yesterday = yesterday.strftime('%m-%d-%y')
    # print(yesterday)
    for i in batch_names:
        foldernameI = str(yesterday) + i
        yesterday_folder_names.append(foldernameI)
    for i in yesterday_folder_names:
        original = os.path.join(desktop_path,i)
        if os.path.exists(original) == False:
            return False
        target = os.path.join(GSCDB_folder_path, i)
        shutil.move(original, target)

def move_fridays_folders():
    fri_folder_names = []
    fri = date.today() - datetime.timedelta(days=3)
    fri = fri.strftime('%m-%d-%y')
    for i in batch_names:
        foldernameI = str(fri) + i
        fri_folder_names.append(foldernameI)
    for i in fri_folder_names:
        original= os.path.join(desktop_path,i)
        if os.path.exists(original) == False:
            return False
        target = os.path.join(GSCDB_folder_path,i)
        shutil.move(original, target)
def make_week_folders(today):
    targetdir = "C:\\Users\kschwartz\Desktop\searches&followups_excel_records"
    os.chdir(targetdir)
    if weekday == 0: #for testing change this to 4
        weekof = "WeekOf_"
        newWeekFolder = weekof + today
        print(f"new week directory is {newWeekFolder}")
        os.mkdir(newWeekFolder)
        newWeekPath = os.path.join(targetdir,newWeekFolder)
        os.chdir(newWeekPath)
def make_todaysFM_folders(today):
    if weekday == 0:
        dOfWeek = "Monday_"
        FMfolder = dOfWeek + today + "_FM_excelSheets"
        FollowupFolder = dOfWeek + today + "_followup_excelSheets"
        print(f"new folders are [{FollowupFolder},{FMfolder}]")
        os.mkdir(FMfolder)
        os.mkdir(FollowupFolder)
    elif weekday == 1:
        dOfWeek = "Tuesday_"
        FMfolder = dOfWeek + today + "_FM_excelSheets"
        FollowupFolder = dOfWeek + today + "_followup_excelSheets"
        print(f"new folders are [{FollowupFolder, FMfolder}]")
        os.mkdir(FMfolder)
        os.mkdir(FollowupFolder)
    elif weekday == 2:
        dOfWeek = "Wednesday_"
        FMfolder = dOfWeek + today + "_FM_excelSheets"
        FollowupFolder = dOfWeek + today + "_followup_excelSheets"
        print(f"new folders are [{FollowupFolder, FMfolder}]")
        os.mkdir(FMfolder)
        os.mkdir(FollowupFolder)
    elif weekday == 3:
        dOfWeek = "Thursday_"
        FMfolder = dOfWeek + today + "_FM_excelSheets"
        FollowupFolder = dOfWeek + today + "_followup_excelSheets"
        print(f"new folders are [{FollowupFolder, FMfolder}]")
        os.mkdir(FMfolder)
        os.mkdir(FollowupFolder)
    elif weekday == 4:
        dOfWeek = "Friday_"
        FMfolder = dOfWeek + today + "_FM_excelSheets"
        FollowupFolder = dOfWeek + today + "_followup_excelSheets"
        print(f"new folders are [{FollowupFolder,FMfolder}]" )
        os.mkdir(FMfolder)
        os.mkdir(FollowupFolder)

if weekday == 0:
    fridaysfolders = move_fridays_folders()
    if fridaysfolders == False:
        print("Cannot find Friday's folders\n")
else:
    yesterdaysfolders = move_yesterdays_folders()
    if yesterdaysfolders == False:
        print("Cannot find yesterday's folders\n")
make_todays_folders()


def makeFMfolder():
    today = date.today()
    today = today.strftime('%m-%d-%y')
    followUps = "followUps_batch"
    FM = "FM_batch"
    today = str(today)
    make_week_folders(today)
    make_todaysFM_folders(today)
makeFMfolder()