# import pyperclip3
# import clipboard
# from PIL import Image
# import pyautogui
# import webbrowser,time
# from pathlib import Path
# import openpyxl
# import pyperclip3
# import clipboard
# from PIL import Image
import pyautogui
import webbrowser,time
# from pathlib import Path
import openpyxl
# from console.utils import wait_key

#TODO: MAKE A FUNCTION THAT WILL PRINT IDS TO ID_LIST AND TO THE GA TEMPALTE SHEET
# nexttab = pyautogui.hotkey('ctrl','tab', interval=.1)
# prev_tab = pyautogui.hotkey('ctrl','shift','tab', interval=.1)
# tab = pyautogui.press('tab',interval=.02)
# enter = pyautogui.press('enter',interval=.02)

#TODO LATEST: MAKE A PROGRAM AFTER THIS OPEN IDS PROGRAM THAT CAN ACTIVATE FUNCTIONS BY WATCHING KEYBOARD INPUT, SPECIFIC SHORTCUTS WILL ACTIVE MOUSE MOVEMENTS TO CERTAIN PLACES.
#TODO , START WORKING ON A WATCH AND RESPOND PROGRAM THAT JUST WRITES HELLO WHEN A CERTAIN KEY COMBO IS PRESSED, THEN MOVE ONTO HAVING IT MOVE TO A LOCATION WHEN THE KEY COMBO IS PRESSED

# def auto_start():
#     webbrowser.open('https://atm.accuratebackground.com/atm/findSearch.html')
#     time.sleep(1)
#     for i in range(1):
#
#         time.sleep(.05)
#     time.sleep(.5)
#     time.sleep(.5)
#     pyautogui.hotkey('winleft', 'left', interval=.07)
#     pyautogui.hotkey('winleft', 'left', interval=.07)
#     time.sleep(.5)
#     webbrowser.open('https://atm.accuratebackground.com/atm/findSearch.html')
#TODO: make a wxglaze age calculator

def get_ids():
    ids_wb = openpyxl.load_workbook('py_id_excel2.xlsm')
    ids_sheet = ids_wb['ids_list']

    ids_list = []
    for rowOfCellObjects in ids_sheet['A2':'A51']:
        for cellObj in rowOfCellObjects:
            # print(cellObj.coordinate, cellObj.value)
            if cellObj.value != None:
                ids_list.append(cellObj.value)
    return ids_list

# def drag_vars():
#     lastdrag = [(1051,374),(1151,370)]
#     akalastdrag = [(1069,391),(1154,389)]
#     firstdrag = [(1196,375), (1340,372)]
#     akafirstdrag = [(1243,389),(1341,387)]
#     dobdrag = [ (1573,392), (1758,389)]
#     lastlist = []
#     firstlist = []
#     doblist = []
#     akafirstlist = []
#     akalastlist = []
#     draglist = [lastdrag,firstdrag,dobdrag,akalastdrag,akafirstdrag]
# #     listlist = [lastlist,firstlist,doblist,akalastlist,akafirstlist]
# def get_clipboard():
#     clipboard_content = clipboard.paste()
#     clipboard_content = (clipboard_content.replace(' ', ''))
#     print(clipboard_content)
# def movedrag(start,end):
#     pyautogui.click(start, duration = .25)
#     pyautogui.dragTo(end, duration = .15)
#     clipboard_content = clipboard.paste()
#     print(clipboard_content)
#     return clipboard_content
# def relative_drag():
#     pyautogui.click()
#     pyautogui.drag(100, 0)
#     clipboard_content = clipboard.paste()
#     print(clipboard_content)
#     return clipboard_content
# def internal_note_box(consent_form):
#     pyautogui.click(1008, 374)
#     for i in range(1,41):
#         tab = pyautogui.hotkey('ctrl', 'w', interval=.07)
#     right = pyautogui.hotkey('right', interval=.07)
#     pyautogui.click()
#     if consent_form == True:
#         pyautogui.typewrite("Emailed consent form to Palo Alto Police Department. Turnaround time is 3 business days.")
#     elif consent_form == False:
#         pyautogui.typewrite("Please obtain a completed GA Release, including full name (First, Middle, and Last) with all Aka's, and email to PR Mailbox. Wet/Digital Signatures are acceptable.")
#     pyautogui.click(1008, 374)
#     # while True:
#     #     print("Y if consent form exists, N if it doesnt")
#     #     check_key = input("Y if consent form exists, N if it doesnt")
#     #     if check_key.upper() != "Y":
#     #         return "Emailed consent form to Palo Alto Police Department. Turnaround time is 3 business days."
#     #         time.sleep(.3)
#     #         pyautogui.hotkey('ctrl', 'r', interval=.07)
#     #         pyautogui.hotkey('enter', interval=.07)
#     #         return True
#     #     elif check_key.upper() != "N":
#     #         pyautogui.typewrite("Please obtain a completed GA Release, including full name (First, Middle, and Last) with all Aka's, and email to PR Mailbox. Wet/Digital Signatures are acceptable.")
#     #         time.sleep(.3)
#     #         pyautogui.hotkey('ctrl', 'r', interval=.07)
#     #         pyautogui.hotkey('enter', interval=.07)
#     #         return False
# def GSCDB(id_list):
#     # prev_tab = pyautogui.hotkey('ctrl', 'shift', 'tab', interval=.07)
#     for i in range(1):
#         time.sleep(.03)
#         next_tab = pyautogui.hotkey('ctrl', 'tab', interval=.07)
#     for i in range(len(id_list)):
#         docinfostring = 'https://atm.accuratebackground.com/atm/documentInfoDocStorage.html?method=view&searchId=' + str(id_list[i])
#         webbrowser.open(docinfostring)
#         time.sleep(.5)
#         while True:
#             "press any key if you see the GA Consent form"""
#             check_key = wait_key()
#             if check_key.upper() != "Y":
#                 return True
#
#
#             elif check_key.upper() != "N":
#                 return False
#                 break
#         close_tab = pyautogui.hotkey('ctrl', 'w', interval=.07)
# # enter == selects status
# # select_awaiting_assignment = up_arrow_x6
# # 28 == status
# # Internal_note_box == 41 tabs
# #
# #         nexttab = pyautogui.hotkey('ctrl', 'tab', interval=.1)
# def DJSON():
#     webbrowser.open_new('https://www.nsopw.gov/en/Search/Results')
def make_chrome_window():
    fw = pyautogui.getWindowsWithTitle('ATM - Google Chrome')
    pyautogui.scroll(200)
    if len(fw) == 0:
        print("l is 0")
        webbrowser.open('https://atm.accuratebackground.com/atm/login.jsp')
        fw = pyautogui.getWindowsWithTitle('Vendor Login | Accurate Background - Google Chrome')
    fw = fw[0]
    fw.width = 974
    fw.topleft = (953, 0)
# def find_picture_location(png_name):
#     finder_pngs = Path("C:\\Users\kschwartz\Desktop\extra_stuff_i_guess\state-py3\GUI_FIND_PNGS")
#     openedImage = Image.open(finder_pngs/png_name)
#     # searchbox = (1201, 391, 390, 23)
#     imagefinder = pyautogui.locateOnScreen(openedImage, confidence=.6, grayscale = True)#, region=searchbox)
#     print(f"imagefinger is:{imagefinder}")
#     imagefinder = pyautogui.center(imagefinder)
#     pyautogui.moveTo(imagefinder, duration=.15)
# def goto_first_prof(id_list):
#     make_chrome_window()
#     for i in range(len(id_list)):
#         time.sleep(.05)
#
#         #
#         pyautogui.click(1008, 374)
#         for h in range(4):
#             pyautogui.press('down')
#     # for i in range(len(id_list)):
#     #     time.sleep(.05)
#     #     # test_pngs = ["dob_png.png","first_png.png","last_png.png"]
#     #     # for i in test_pngs:
#     #     #     find_picture_location(i)
#     #     #     pyautogui.move(15, 0)
#     #     #     time.sleep(.05)
#     #     #     clipboard2 = relative_drag()
#     #     #     if i == 0:
#     #     #         doblist.append(clipboard2)
#     #     #     elif i == 1:
#     #     #         firstlist.append(clipboard2)
#     #     #     elif i == 2:
#     #     #         lastlist.append(clipboard2)
#     #     #     time.sleep(.15)
#     #     # pyautogui.hotkey('ctrl', 'shift', 'tab', interval=.1)
#     #     pyautogui.click(1008, 374)
#     #     time.sleep(.05)
#     #     last = movedrag(lastdrag[0],lastdrag[1])
#     #     lastlist.append(last)
#     #     first = movedrag(firstdrag[0],firstdrag[1])
#     #     firstlist.append(first)
#     #     dob = movedrag(dobdrag[0],dobdrag[1])
#     #     doblist.append(dob)
#     #     akafirst = movedrag(akafirstdrag[0],akafirstdrag[1])
#     #     akafirstlist.append(akafirst)
#     #     akalast = movedrag(akalastdrag[0],akalastdrag[1])
#     #     akalastlist.append(akalast)
#     #     pyautogui.hotkey('ctrl', 'shift', 'tab', interval=.1)
#     #     # draglist = [lastdrag, firstdrag, dobdrag, akalastdrag, akafirstdrag]
#     #     # listlist = [lastlist, firstlist, doblist, akalastlist, akafirstlist]
#
# for i in range(0,len(draglist)):
#     adder = movedrag(draglist[i[0]],draglist[i[1]])
#     listlist[i].append(adder)
#this function is older
def get_new_tab():
    # make_chrome_window()
    # new_tab = pyautogui.hotkey('ctrl', 't', interval=.15)
    webbrowser.open('https://atm.accuratebackground.com/atm/findSearch.html')
    time.sleep(1.1)
    # webbrowser.open('https://www.google.com/')
    # # time.sleep(.15)
    # # open_bm_manager = pyautogui.hotkey('alt', 'e', interval=.05)
    # # time.sleep(.05)
    # # press_b = pyautogui.hotkey('b', interval=.05)
    # # time.sleep(.05)
    # # press_left = pyautogui.hotkey('right', interval=.05)
    # # time.sleep(1)
    # # press_down = pyautogui.press('down', presses=6, interval=.01)
    # # press_enter = pyautogui.press('enter', interval=.05)
    # # prev_tab = pyautogui.hotkey('ctrl','shift','tab', interval=.07)
    # # close_tab = pyautogui.hotkey('ctrl','w', interval=.07)
# def find_picture_location(png_name):
#     finder_pngs = Path("C:\\Users\kschwartz\Desktop\extra_stuff_i_guess\state-py3\Finder_pngs")
#     openedImage = Image.open(png_name)
#     searchbox = (1201,391,390,23)
#     imagefinder = pyautogui.locateOnScreen(openedImage,confidence = .6,region = searchbox )
#     print(f"imagefinger is:{imagefinder}")
#     imagefinder = pyautogui.center(imagefinder)
#     pyautogui.moveTo(imagefinder,duration = .15)
def search_id_fetch(id_list):
    ids = id_list
    ids_str = []
    for i in ids:
        i = str(i)
        ids_str.append(i)
    ids_list = ids_str
    # for h in range(0,len(ids_list)):
    make_chrome_window()
    for h in ids_list:
        time.sleep(.05)
        get_new_tab()
        findsearch = ["enter_id_box.png", "search_press_box.png"]
        time.sleep(.05)
        enter_id_box = (1340,405)
        press_search_button = (1625,405)
        pyautogui.click(enter_id_box)
        time.sleep(.05)
        pyautogui.typewrite(h)
        time.sleep(.15)
        pyautogui.click(press_search_button)
# auto_start()
id_list = get_ids()
search_id_fetch(id_list)
pyautogui.hotkey('ctrl','2', interval=.07)
Testimport = "test succeeded"
# goto_first_prof(id_list)
# for i in range(0,len(listlist)):
#     print(listlist[i])

# last = movedrag(lastdrag[0],lastdrag[1])
# lastlist.append(last)

#SHIFT TAB == FIELD BACK IN CHROME
#ALT LEFT ARROW == GO BACK TO PREVIOUS PAGE IN CHROME
#ALT RIGHT ARROW == GO FORWARD PAGE IN CHROME