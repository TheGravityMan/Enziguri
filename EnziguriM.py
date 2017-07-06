# INITIALIZE, IMPORT TTK, MESSAGEBOX, PYAUTOGUI, TIME, WIN32GUI, PYWINAUTO
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import pyautogui
import time
from win32 import win32gui
import re
from pywinauto.win32functions import ShowWindow
from pywinauto.win32defines import SW_MAXIMIZE


# CREATES OBJECTS WHICH WILL REPRESENT CURRENT TESSITURA WINDOW TO MAKE ACTIVE
class WindowMgr:
    def __init__(self):
        # INITIALIZE HANDLE
        self._handle = None

    def find_window(self, class_name, window_name=None):
        # FIND A WINDOW BY IT'S SPECIFIC CLASS NAME
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        # PASS TO WIN32GUI.ENUMWINDOWS() TO CHECK ALL OPEN WINDOWS FOR WILDCARD
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        # SET WINDOW TO FOREGROUND, MAXIMIZE WINDOW
        win32gui.SetForegroundWindow(self._handle)
        ShowWindow(self._handle, SW_MAXIMIZE)
        #print(self._handle)

    def value_get(self):
        return self._handle


# CREATE OBJECT OF CLASS WINDOWMGR, FIND ACTIVE TESSY WINDOW, MAKE ACTIVE AND MAXIMIZE
def find_tessy():
    act = WindowMgr()
    act.find_window_wildcard('.*Tessitura.*')
    act.set_foreground()
    # ZERO OUT TESSITURA TO MAIN WINDOW
    pyautogui.press('esc')
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'l')

    return


# CHECKS OPEN WINDOW TITLES IN ORDER TO AVOID MULTIPLE OCCURRENCES OF MAS PROGRAM
def run_check():
    win_run = WindowMgr()
    win_run.find_window_wildcard('-----Swipe MN ID Card-----')
    if win_run.value_get() is not None:
        closewindow()
    win_run2 = WindowMgr()
    win_run2.find_window_wildcard('New Member Information----228')
    if win_run2.value_get() is not None:
        closewindow()

    return


# DURING EMAIL AND ADDRESS CHECK, DETERMINE IF MEMBER EXISTS
def email_exist_check(if_found,loop_count):
    email_ex = WindowMgr()
    email_ex.find_window_wildcard('Information')
    if email_ex.value_get() is not None:
        if_found = True
        return if_found,loop_count
    else:
        loop_count = loop_count + 1
        time.sleep(1)
        return if_found,loop_count


# EXIT SCRIPT UPON CLICKING 'X'
def closewindow():
    sys.exit()


# DOUBLE CHECK EMAIL, MAKE DEEP CHECK FOR ACCURACY
def email_check(val, mail):
    if '@' not in mail or '.' not in mail:
        val = False
    if mail[0:1] == '@' or mail[0:1] == '.':
        val = False
    if mail.find('@') > mail.find('.'):
        val = False
    if mail[len(mail) - 1] == '.':
        val = False
    if val:
        mail_list = mail.split('@')
        mail_list2 = mail_list[1].split('.')
        if len(mail_list2[0]) < 2 or len(mail_list2[1]) < 2:
            val = False

    return val


# WHEN OK BUTTON IS CLICKED, RUN A FINAL VALUE CHECK, AND SAVE ALL VALUES INTO MASTER
# LIST OF NECESSARY VALUES. ONCE COMPLETE, DESTROY THE INPUT WINDOW
def clickbutton(first, mid, last, address, apartment, city, zipcode, email, Sfirst, Smid, Slast,
                adult, Pphone, Sphone, MEMlvl, NumKids1, NumKids2, NumKids3, school, school2, district,
                district2, donation, NPselect, NorPCA, NPCAname, root):
    valid = True
    error_reason = empty_str

    # SLICE MEMBERSHIP LEVEL VALUE DOWN TO SINGLE VALUE REPRESENTING IT'S LEVEL
    MEMlvl = MEMlvl[0:1]

    # CHECK ENTRIES TO INSURE VALIDITY OF SPECIFIC VALUES
    list_check = [first, last, address, city, Pphone]
    count = 0
    for i in list_check:
        if i == empty_str:
            valid = False
            if count == 0:
                error_reason = 'Primary First Name'
            elif count == 1:
                error_reason = 'Primary Last Name'
            elif count == 2:
                error_reason = 'Address'
            elif count == 3:
                error_reason = 'City'
            elif count == 4:
                error_reason = 'Primary Phone'
        count = count+1

    # CHECK VALIDITY OF MEMBERSHIP LEVEL
    if MEMlvl =='-':
        valid = False
        error_reason = 'Membership Level'

    # CHECK VALIDITY OF SECOND ADULT IF APPLICABLE
    if adult is True and (Sfirst == empty_str or Slast == empty_str):
        valid = False
        error_reason = 'Second Adult Name'

    # CHECK LENGTH/CONTENTS OF ZIPCODE
    if len(zipcode) != 5 and len(zipcode) != 9:
        valid = False
        error_reason = 'Zipcode'
    if not zipcode.isdigit():
        valid = False
        error_reason = 'Zipcode'

    # CHECK VALIDITY OF EMAIL ADDRESS
    if email != empty_str:
        valid = email_check(valid, email)
        if not valid:
            error_reason = 'E-mail'

    # CHECK VALIDITY FOR 'NUMBER OF CHILDREN' ENTRY
    if MEMlvl == 'A' or MEMlvl == 'E' or MEMlvl == 'G' or MEMlvl == 'I':
        if NumKids1 == empty_str:
            valid = False
            error_reason = 'Number of Kids'
    elif MEMlvl == 'C':
        if NumKids2 == empty_str:
            valid = False
            error_reason = 'Number of Kids'
    elif MEMlvl == 'K' or MEMlvl == 'L' or MEMlvl == 'M' or MEMlvl == 'N':
        if NumKids3 == empty_str:
            valid = False
            error_reason = 'Number of Kids'

    # SAVE ALL VALUES INTO THE MASTER LIST
    if valid:
        master['PrimaryFirst'] = first
        master['PrimaryMiddle'] = mid
        master['PrimaryLast'] = last
        master['SecondFirst'] = Sfirst.title()
        master['SecondMiddle'] = Smid.title()
        master['SecondLast'] = Slast.title()
        master['Address'] = address
        master['Apartment'] = apartment
        master['City'] = city
        master['Zipcode'] = zipcode
        master['Email'] = email
        master['AdultState'] = adult
        master['PrimaryPhone'] = Pphone
        master['SecondaryPhone'] = Sphone
        master['MembershipLevel'] = MEMlvl
        master['DonationAmount'] = donation
        master['NannyState'] = NPselect
        master['NannyOrPca'] = NorPCA
        master['NPCAName'] = NPCAname.title()

        # DEPENDING ON WHICH MEMBERSHIP LEVEL IS SELECTED, NUMBER OF KIDS VALUE IS SELECTED ACCORDINGLY
        if MEMlvl == 'A' or MEMlvl == 'E' or MEMlvl == 'G' or MEMlvl == 'I':
            master['NumberKids'] = NumKids1
        elif MEMlvl == 'C':
            master['NumberKids'] = NumKids2
            master['NameOfSchool'] = school.title()
            master['DistrictNumber'] = district
        elif MEMlvl == 'D':
            master['NameOfSchool'] = school2.title()
            master['DistrictNumber'] = district2
        elif MEMlvl == 'K' or MEMlvl == 'L' or MEMlvl == 'M' or MEMlvl == 'N':
            master['NumberKids'] = NumKids3
        elif MEMlvl == 'B' or MEMlvl == 'F' or MEMlvl == 'H' or MEMlvl == 'J':
            master['NumberKids'] = empty_str

        root.destroy()
        return()

    else:
        # IF ANY PART IS INVALID, RAISE ERROR MESSAGE BOX TO USER
        messagebox.showerror(title='Invalid Entry', message='The value entered for '+error_reason+' is invalid.'
                                                            ' Please check all entries and try again.')
        return()


# IF SECOND ADULT BOX IS CHECKED, SHOW SECOND ADULT ENTRY FORM. IF IT IS UNCHECKED,
# HIDE SECOND ADULT ENTRY FORMS.
def second_ad(adultstate, frame):
    if adultstate:
        print('cooooool')
        frame.grid(column=1, row=2, sticky=(N, W, E, S))
    elif not adultstate:
        frame.grid_remove()
    return()


# CHECKS INPUT OF MEMBERSHIP LEVEL, SHOWS THE APPROPRIATE INPUT FORMS DEPENDING
# ON THE SELECTED LEVEL. IF NO FURTHER FORMS ARE REQUIRED, DESTROY CURRENTLY
# DISPLAYED FORMS
def mem_level(select, frame6, frame7, frame8, frame12):
    if select[0:1] == 'A' or select[0:1] == 'E' or select[0:1] == 'G' or select[0:1] == 'I':
        frame7.grid_remove()
        frame8.grid_remove()
        frame12.grid_remove()
        frame6.grid(column=1, row=6, sticky=(N, W, E, S))
    elif select[0:1] == 'C':
        frame6.grid_remove()
        frame8.grid_remove()
        frame12.grid_remove()
        frame7.grid(column=1, row=6, sticky=(N, W, E, S))
    elif select[0:1] == 'D':
        frame6.grid_remove()
        frame7.grid_remove()
        frame8.grid_remove()
        frame12.grid(column=1, row=6, sticky=(N, W, E, S))
    elif select[0:1] == 'K' or select[0:1] == 'L' or select[0:1] == 'M' or select[0:1] == 'N':
        frame6.grid_remove()
        frame7.grid_remove()
        frame12.grid_remove()
        frame8.grid(column=1, row=6, sticky=(N, W, E, S))
    else:
        frame6.grid_remove()
        frame7.grid_remove()
        frame8.grid_remove()
        frame12.grid_remove()

    return True


# CHECKS STATE OF NANNY/PCA CHECKBOX. IF CHECKED, SHOW FORM FOR NAME INPUT. IF UNCHECKED,
# HIDE FORM FOR INPUT
def add_nanny(value, frame):
    if value:
        frame.grid(column=1, row=8, sticky=W)
    else:
        frame.grid_remove()

    return()


# PROCESS TO PARSE THE ID CARD INFORMATION AND SAVE IT TO IT'S CORRESPONDING
# VARIABLES
def swipe_save(member_info, idwindow):
    try:
        id_list = member_info.split('^')
        id_city = id_list[0][3:len(id_list[0])]
        id_name = id_list[1]
        names_list = (id_name.split(' '))
        num_names = len(names_list)
        id_first = names_list[0]
        id_last = names_list[num_names - 1]
        id_mid = ''
        for i in names_list[1:num_names-1]:
            id_mid = id_mid + ' ' + i
        id_mid = (id_mid[1:len(id_mid)])
        id_address = id_list[2]

        # SAVE VALUES INTO CORRESPONDING VARIABLES
        master['PrimaryFirst'] = id_first.title()
        master['PrimaryMiddle'] = id_mid.title()
        master['PrimaryLast'] = id_last.title()
        master['Address'] = id_address.title()
        master['City'] = id_city.title()

        idwindow.destroy()
        return()

    # HANDLES EXCEPTIONS IN ID SWIPE, OR IF LEFT BLANK
    except IndexError:
        idwindow.destroy()
        return()


# BUILD WINDOW TO ACCEPT SWIPE OF ID CARD. INFORMATION IS HIDDEN WITH '*' AND CANNOT BE
# ACCESSED BY THE USER
def IDswipe():
    root2 = Tk()
    root2.title('-----Swipe MN ID Card-----')

    idmain = ttk.Frame(root2, padding="10 10 12 12")
    idmain.grid(column=1, row=1, sticky=(N, W, E, S))
    idmain.rowconfigure(1, weight=3)
    idmain.columnconfigure(1, weight=1)

    identry = ttk.Entry(idmain, width=30, show='*')
    identry.grid(column=2, row=1, sticky=W)

    ttk.Label(idmain, text='Please Swipe ID Card: ').grid(column=1, row=1, sticky=W)
    ttk.Label(idmain, text="'OK' to continue without ID Swipe").grid(column=2, row=2, sticky=W)

    idbutton = ttk.Button(idmain, text='OK', command=lambda: swipe_save(identry.get(), root2))
    idbutton.grid(column=1, row=2, sticky=W)
    identry.focus_set()

    # SETS INITIAL SCREEN POSITION OF ID SWIPE WINDOW
    root2.geometry('+700+400')
    root2.protocol('WM_DELETE_WINDOW', closewindow)
    root2.attributes('-topmost', True)
    root2.mainloop()


# BUILD MAIN INPUT FORM FOR USER. DEPENDING ON IF THIS IS THE FIRST ENTRY, WILL AUTO-FILL
# ANY FIELEDS ALREADY COMPLETED.
def firstbox():

    # BUILD ROOT WINDOW, DISPLAY TITLE
    root = Tk()
    root.title('New Member Information----228')

    # BUILD ALL FRAMES, WHETHER OR NOT THEY ARE VISIBLE AT THE START
    # MAIN--NAME OF FIRST ADULT
    main = ttk.Frame(root, padding="10 10 12 12")
    main.grid(column=1, row=1, sticky=(N, W, E, S))
    main.rowconfigure(1, weight=3)
    main.columnconfigure(1, weight=1)

    # MAIN2-ADDRESS INFORMATION
    main2 = ttk.Frame(root, padding="10 0 12 12")
    main2.grid(column=1, row=3, sticky=W)

    # MAIN3-SECOND ADULT
    main3 = ttk.Frame(root, padding='10 0 12 12')

    # MAIN4-OTHER INFORMATION
    main4 = ttk.Frame(root, padding='10 0 12 12')
    main4.grid(column=1, row=4, sticky=W)

    # MAIN5-MEMBERSHIP LEVEL
    main5 = ttk.Frame(root, padding='10 10 12 12')
    main5.grid(column=1, row=5, sticky=W)

    # MAIN6-HOUSEHOLD LEVELS
    main6 = ttk.Frame(root, padding='10 0 12 12')

    # MAIN7-EDUCATOR,HOUSEHOLD LEVEL
    main7 = ttk.Frame(root, padding='10 0 12 12')

    # MAIN8-SUPPORTING LEVELS
    main8 = ttk.Frame(root, padding='10 0 12 12')

    # MAIN9-NANNY/PCA SELECT
    main9 = ttk.Frame(root, padding='10 0 12 12')
    main9.grid(column=1, row=7, sticky=W)

    # MAIN10-NANNY/PCA FORM BUILD
    main10 = ttk.Frame(root, padding='10 0 12 12')

    # MAIN11-SUBMIT BUTTON
    main11 = ttk.Frame(root, padding='10 0 12 12')
    main11.grid(column=1, row=9, sticky=W)

    # MAIN12-EDUCATOR,DUAL LEVEL (IMPLEMENTED LATER)
    main12 = ttk.Frame(root, padding='10 0 12 12')

    # BUILD OF MAIN, CHECKBOX FOR ADDITIONAL ADULT SELECTION
      #firstname = StringVar()
    Pfirst_entry = ttk.Entry(main, width=30)
    Pfirst_entry.grid(column=2, row=1, sticky=W)
    Pmid_entry = ttk.Entry(main, width=30)
    Pmid_entry.grid(column=2, row=2, sticky=W)
    Plast_entry = ttk.Entry(main, width=30)
    Plast_entry.grid(column=2, row=3, sticky=W)
    adult_state = BooleanVar()
    secondno = ttk.Checkbutton(main, text='Second Adult?', variable=adult_state,
                               command=lambda: second_ad(adult_state.get(), main3))
    secondno.grid(column=2, row=4, sticky=W)

    # LABELS FOR MAIN
    ttk.Label(main, width=13, text='Second Adult?').grid(column=1, row=4, sticky=W)
    ttk.Label(main, text='First: ').grid(column=1, row=1, sticky=W)
    ttk.Label(main, text='Middle: ').grid(column=1, row=2, sticky=W)
    ttk.Label(main, text='Last: ').grid(column=1, row=3, sticky=W)

    # MAIN3 BUILD
    Sfirst_entry = ttk.Entry(main3, width=30)
    Sfirst_entry.grid(column=2, row=5, sticky=W)
    Smid_entry = ttk.Entry(main3, width=30)
    Smid_entry.grid(column=2, row=6, sticky=W)
    Slast_entry = ttk.Entry(main3, width=30)
    Slast_entry.grid(column=2, row=7, sticky=W)

    # LABELS FOR MAIN3
    ttk.Label(main3, width=13, text='First: ').grid(column=1, row=5, sticky=W)
    ttk.Label(main3, text='Middle: ').grid(column=1, row=6, sticky=W)
    ttk.Label(main3, text='Last: ').grid(column=1, row=7, sticky=W)

    # MAIN2 BUILD
    Add_label = '----------------Enter Street Address-----------------'
    ttk.Label(main2, text=Add_label).grid(column=1, row=1, sticky=W)
    ttk.Label(main2, text='Street Address: ').grid(column=1, row=2, sticky=W)
    ttk.Label(main2, text='Apt. #: ').grid(column=1, row=4, sticky=W)
    ttk.Label(main2, text='City: ').grid(column=1, row=6, sticky=W)
    ttk.Label(main2, text='Zipcode: ').grid(column=1, row=8, sticky=W)

    street_add = ttk.Entry(main2, width=45)
    street_add.grid(column=1, row=3, sticky=W)
    apt_num = ttk.Entry(main2, width=10)
    apt_num.grid(column=1, row=5, sticky=W)
    city_ent = ttk.Entry(main2, width=20)
    city_ent.grid(column=1, row=7, sticky=W)
    zip_ent = ttk.Entry(main2, width=10)
    zip_ent.grid(column=1, row=9, sticky=W)

    # MAIN4 BUILD
    Other_label = '-------------Enter Addition Information--------------'
    ttk.Label(main4, text=Other_label).grid(column=1, row=1, sticky=W)
    ttk.Label(main4, text='Primary Phone: ').grid(column=1, row=2, sticky=W)
    ttk.Label(main4, text='Second Phone: ').grid(column=1, row=4, sticky=W)
    ttk.Label(main4, text='Email Address: ').grid(column=1, row=6, sticky=W)

    pphone = ttk.Entry(main4, width=15)
    pphone.grid(column=1, row=3, sticky=W)
    sphone = ttk.Entry(main4, width=15)
    sphone.grid(column=1, row=5, sticky=W)
    email_ent = ttk.Entry(main4, width=40)
    email_ent.grid(column=1, row=7, sticky=W)

    # MAIN5 BUILD
    Empty_label = '                    Select Membership Level                           '
    ttk.Label(main5, text=Empty_label).grid(column=1, row=1)

    level_select = ttk.Combobox(main5, width=30,
                                values=('A   Household:  $99', 'B   Dual:  $69', 'C   Educator Household:  $84',
                                        'D   Educator Dual:  $58', 'E    Senior Household:  $94',
                                        'F    Senior Dual:  $69', 'G   Great-Tix Household:  $49.50',
                                        'H   Great-Tix Dual:  $34.50', "I     Children's Museum Household:  $89",
                                        "J     Children's Museum Dual:  $59", 'K    Darwin:  $150-$249',
                                        'L    Carver:  $250-$499', 'M   Hopper:  $500-$999',
                                        'N    Einstein:  $1000-$2499'),
                                validatecommand=lambda: mem_level(level_select.get(), main6, main7, main8, main12),
                                validate='focus')

    # DISABLES ABILITY FOR USER TO CHANGE VALUE INSIDE COMBOBOX AND INITIALIZE
    level_select.state(['!disabled', 'readonly'])
    level_select.set('-----------------------------------------------------')
    level_select.grid(column=1, row=2, sticky='N W E S')

    # MAIN6 BUILD
    ttk.Label(main6, text='Number of Kids: ').grid(column=1, row=1, sticky=W)

    num_kids1 = ttk.Entry(main6, width=5)
    num_kids1.grid(column=2, row=1, sticky=W)

    # MAIN7 BUILD
    ttk.Label(main7, text='Number of Kids: ').grid(column=1, row=1, sticky=W)
    ttk.Label(main7, text='Name of School: ').grid(column=1, row=2, sticky=W)
    ttk.Label(main7, text='District #: ').grid(column=1, row=3, sticky=W)

    num_kids2 = ttk.Entry(main7, width=5)
    num_kids2.grid(column=2, row=1, sticky=W)
    school = ttk.Entry(main7)
    school.grid(column=2, row=2, sticky=W)
    district = ttk.Entry(main7)
    district.grid(column=2, row=3, sticky=W)

    # MAIN12 BUILD (IMPLEMETED LATER)
    ttk.Label(main12, text='Name of School: ').grid(column=1, row=1, sticky=W)
    ttk.Label(main12, text='District #: ').grid(column=1, row=2, sticky=W)

    school2 = ttk.Entry(main12)
    school2.grid(column=2, row=1, sticky=W)
    district2 = ttk.Entry(main12)
    district2.grid(column=2, row=2, sticky=W)

    # MAIN8 BUILD
    ttk.Label(main8, text='Number of Kids: ').grid(column=1, row=1, sticky=W)
    ttk.Label(main8, text='Donation Amount ($): ').grid(column=1, row=2, sticky=W)

    num_kids3 = ttk.Entry(main8, width=5)
    num_kids3.grid(column=2, row=1, sticky=W)
    donation = ttk.Entry(main8, width=10)
    donation.grid(column=2, row=2, sticky=W)

    # MAIN9 BUILD
    n_or_pca = BooleanVar()
    nanny_pca = ttk.Checkbutton(main9, text='Add Nanny/PCA?', variable=n_or_pca,
                                command=lambda: add_nanny(n_or_pca.get(), main10))
    nanny_pca.grid(column=1, row=1, sticky=W)

    # MAIN10 BUILD
    ttk.Label(main10, text='Enter Name:').grid(column=2, row=1, sticky=W)
    choose_n_or_pca = ttk.Combobox(main10, width=6, values=('Nanny', 'PCA'))

    # DISABLES ABILITY FOR USER TO CHANGE VALUE INSIDE OF COMBOBOX
    choose_n_or_pca.state(['!disabled', 'readonly'])
    choose_n_or_pca.set('Nanny')
    choose_n_or_pca.grid(column=1, row=2, sticky=W)

    name_npca = ttk.Entry(main10, width=38)
    name_npca.grid(column=2, row=2, sticky=W)
    name_npca.setvar('----------------Enter Name---------------')

    # Main11 Build
    bbutton = ttk.Button(main11, text="OK",
                         command=lambda: clickbutton(Pfirst_entry.get(), Pmid_entry.get(), Plast_entry.get(),
                                                     street_add.get(), apt_num.get(), city_ent.get(),
                                                     zip_ent.get(), email_ent.get(), Sfirst_entry.get(),
                                                     Smid_entry.get(), Slast_entry.get(), adult_state.get(),
                                                     pphone.get(), sphone.get(), level_select.get(), num_kids1.get(),
                                                     num_kids2.get(), num_kids3.get(), school.get(), school2.get(),
                                                     district.get(), district2.get(), donation.get(), n_or_pca.get(),
                                                     choose_n_or_pca.get(), name_npca.get(), root))
    bbutton.grid(column=1, row=1, sticky=W)
    bbutton.grid_configure(padx=5, pady=10)

    # ADDS ADDITIONAL SPACE AROUND PARTICULAR FRAMES
    for child in main.winfo_children(): child.grid_configure(padx=5, pady=5)
    for child in main3.winfo_children(): child.grid_configure(padx=5, pady=5)
    for child in main7.winfo_children(): child.grid_configure(padx=5, pady=5)
    for child in main8.winfo_children(): child.grid_configure(padx=5, pady=5)
    for child in main9.winfo_children(): child.grid_configure(padx=5, pady=5)
    for child in main12.winfo_children(): child.grid_configure(padx=5, pady=5)

    # INITIALIZES THE FORM TO FILL IN ANY INFORMATION ALREADY OBTAINED FROM ID SWIPE
    Pfirst_entry.insert(END, master['PrimaryFirst'])
    Pmid_entry.insert(END, master['PrimaryMiddle'])
    Plast_entry.insert(END, master['PrimaryLast'])
    street_add.insert(END, master['Address'])
    city_ent.insert(END, master['City'])

    # SETS INITIAL SCREEN POSITION FOR ROOT WINDOW
    root.geometry('+700+50')
    root.protocol('WM_DELETE_WINDOW', closewindow)
    root.mainloop()

    return()


def tessyfill(master):

    # INITIALIZE VARIABLES
    sMatch = True
    valid = True
    found = False
    ecount = 0
    acount = 0
    try:
        pyautogui.FAILSAFE = TRUE

        pyautogui.hotkey('ctrl', 'l', 't')

        # FOLLOW THREAD IF EMAIL IS NOT BLANK, SEARCH EMAIL IN TESSY
        if master['Email'] != empty_str:
            pyautogui.hotkey('alt', 'd')
            pyautogui.typewrite(['e', 'tab', 'tab'])
            pyautogui.typewrite(master['Email'])
            pyautogui.typewrite('\n')

            # SEARCH FOR MESSAGE BOX ON SCREEN TO CONFIRM OR DENY MATCH
            while not found:
                found,ecount = email_exist_check(found,ecount)
                if found:
                    pyautogui.typewrite('\n')
                if ecount > 10:
                    found = True
                    sMatch = False

        # SEARCH FOR NAME, ADDRESS, AND CONFIRM OR DENY MATCH
        if sMatch:
            pyautogui.hotkey('alt', 'b')
            time.sleep(1)
            pyautogui.typewrite('\t')
            pyautogui.typewrite(master['PrimaryLast'])
            pyautogui.typewrite('\t')
            pyautogui.typewrite(master['PrimaryFirst'])
            pyautogui.typewrite('\t')
            pyautogui.typewrite(master['Address'])
            if master['Apartment'] != empty_str:
                pyautogui.press('space')
                pyautogui.typewrite('#')
                pyautogui.typewrite(master['Apartment'])
            pyautogui.typewrite('\t')
            pyautogui.typewrite(master['Zipcode'])
            pyautogui.typewrite('\n')
            found = False

            # SECOND CHECK FOR MATCH
            while not found:
                found,acount = email_exist_check(found,acount)
                if found:
                    pyautogui.typewrite('\t')
                    pyautogui.typewrite('\n')
                    tessyfill2(master)
                if acount > 10:
                    found = True
                    valid = False

        # IF MATCH FOUND, MOVE TO END SCRIPT
        else:
            print('This member exists')
            valid = False
            return valid

        return valid

    # STOP SCRIPT USING FAILSAFE, MOVE MOUSE TO TOP LEFT CORNER OF SCREEN
    except pyautogui.FailSafeException:
        print('coooooooooooool')
        valid = False
        return valid


# CONTINUE TO FILL TESSITURA
def tessyfill2(master):
    if master['AdultState']:
        pyautogui.typewrite('\n')
        pyautogui.typewrite(['\t', '\t'])
        pyautogui.typewrite(master['PrimaryMiddle'])
        pyautogui.typewrite(['\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t'])
        pyautogui.typewrite(master['SecondFirst'])
        pyautogui.typewrite('\t')
        pyautogui.typewrite(master['SecondMiddle'])
        pyautogui.typewrite('\t')
        pyautogui.typewrite(master['SecondLast'])
        pyautogui.typewrite(['\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t'])

    else:
        pyautogui.press('down')
        pyautogui.typewrite('\n')
        pyautogui.typewrite(['\t', '\t'])
        pyautogui.typewrite(master['PrimaryMiddle'])
        pyautogui.typewrite(['\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t'])

    pyautogui.typewrite(master['City'])
    pyautogui.typewrite(['\t', '\t', '\t', '\t', '\t'])
    pyautogui.typewrite(master['PrimaryPhone'])
    pyautogui.typewrite(['\t', '\t', '\t'])
    pyautogui.typewrite(master['SecondaryPhone'])
    pyautogui.typewrite(['\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', '\t', 'n', '\t'])
    pyautogui.typewrite(master['Email'])
    pyautogui.hotkey('ctrl', 'e')
    pyautogui.press('\n')

    # ADD KIDS IF APPLICABLE
    if master['NumberKids'] != '':
        addkids(master['NumberKids'])

    # ADD NANNY/PCA IF APPLICABLE
    if master['NannyState']:
        addnanny(master['NannyOrPca'], master['NPCAName'])

    # ADD EDUCATOR INFORMATION IF APPLICABLE
    if master['MembershipLevel'] == 'C' or master['MembershipLevel'] == 'D':
        addeducator(master['NameOfSchool'], master['DistrictNumber'])

    return()


# ADD KIDS IN ATTRIBUTES
def addkids(kids):
    pyautogui.hotkey('ctrl', 't')
    pyautogui.typewrite('\n')
    pyautogui.hotkey('alt', 'a')
    pyautogui.typewrite(['n', 'u', '\t'])
    pyautogui.typewrite(kids)
    pyautogui.hotkey('ctrl', 't')
    pyautogui.press('\n')

    return()

# ADD NANNY TO CONNECTIONS
def addnanny(which, name):
    pyautogui.hotkey('ctrl', 'o')
    pyautogui.hotkey('alt', 'a')
    pyautogui.typewrite(['\t', '\t', '\t', 'p', '\t', 'i', '\t', 'n', '\t', '\t'])
    pyautogui.typewrite(which)
    pyautogui.typewrite(': ')
    pyautogui.typewrite(name)
    pyautogui.hotkey('alt', 's')
    pyautogui.typewrite('\n')
    pyautogui.hotkey('alt', 'c')

    return()


# ADD EDUCATOR INFORMATION
def addeducator(school, district):
    pyautogui.hotkey('ctrl', 'r')
    pyautogui.hotkey('alt', 'a')
    pyautogui.typewrite(['e', '\t'])
    pyautogui.typewrite('School: ')
    pyautogui.typewrite(school)
    pyautogui.typewrite('\n')
    pyautogui.typewrite('District: ')
    pyautogui.typewrite(district)
    pyautogui.hotkey('alt', 's')

    return()

# PREPARE TESSITURA TO ADD MEMBERSHIP LEVEL INFORMATION
def applymember1(pfirst, plast, astate, sfirst, slast):
    if not astate:
        pyautogui.typewrite('\n')

        return()

    # CHECK LITERAL VALUES OF NAMES TO DETERMINE LIST ORDER IN CONSTITUENTS WINDOW
    if pfirst != plast:
        if plast.lower() > slast.lower():
            pyautogui.typewrite(['down', 'down', '\n'])
        else:
            pyautogui.typewrite('\n')

        return()

    else:
        if pfirst.lower() > slast.lower():
            pyautogui.typewrite(['down', 'down', '\n'])
        else:
            pyautogui.typewrite('\n')

        return()


# DEPENDING ON MEMBVERSHIP LEVEL, FIND NECESSARY DETAILS
def applymember2(levelmem, donation):
    pyautogui.hotkey('ctrl', 'b')
    pyautogui.hotkey('alt', 'a')

    if levelmem[0:1] == 'A':
        mem_suffix = 'a'
        mem_amt = '99'
    elif levelmem[0:1] == 'B':
        mem_suffix = 'b'
        mem_amt = '69'
    elif levelmem[0:1] == 'C':
        mem_suffix = 'eh'
        mem_amt = '84'
    elif levelmem[0:1] == 'D':
        mem_suffix = 'ed'
        mem_amt = '58'
    elif levelmem[0:1] == 'E':
        mem_suffix = 'c'
        mem_amt = '94'
    elif levelmem[0:1] == 'F':
        mem_suffix = 'd'
        mem_amt = '64'
    elif levelmem[0:1] == 'G':
        mem_suffix = 'h'
        mem_amt = '49.50'
    elif levelmem[0:1] == 'H':
        mem_suffix = 'i'
        mem_amt = '34.50'
    elif levelmem[0:1] == 'I':
        mem_suffix = 'j'
        mem_amt = '89'
    elif levelmem[0:1] == 'J':
        mem_suffix = 'k'
        mem_amt = '59'
    else:
        # supporting level not yet implemented
        mem_suffix = 'z'
        mem_amt = '00'
        return

    # ADD MEMBERSHIP DETAILS TO CONTRIBUTIONS TAB
    pyautogui.typewrite('mem')
    pyautogui.typewrite(mem_suffix)
    pyautogui.typewrite('\t')
    pyautogui.typewrite(mem_amt)
    pyautogui.typewrite(['\t', '\n'])
    pyautogui.hotkey('ctrl', 'q')

    return()


# IF MEMBER EXISTS, GIVE OPTION TO RECALL VALUES PREVIOUSLY ENTERED
def info_recall(master_info):
    root3 = Tk()
    root3.title('Recall Member Information')

    rec1 = ttk.Frame(root3, padding='10 10 12 12')
    rec1.grid(column=1, row=1, sticky=(N, W, E, S))

    # BUTTON TO SHOW INFORMATION
    recbutton = ttk.Button(rec1, text='Recall Member Information',
                           command=lambda: show_recall(master_info, rec1, root3))
    recbutton.grid(column=1, row=1, sticky=W)

    # BUTTON TO QUIT PROGRAM
    nobutton = ttk.Button(rec1, text='NOPE!', command=lambda: closewindow())
    nobutton.grid(column=2, row=1, sticky=W)

    # WINDOW APPEARS ON TOP RIGHT OF SCREEN
    root3.geometry('+1500+100')
    root3.protocol('WM_DELETE_WINDOW', closewindow)
    root3.mainloop()

    return


# SHOWS INFORMATION IF USER DESIRES
def show_recall(mstr, recframe, rec):
    recframe.grid_remove()

    rec2 = ttk.Frame(rec, padding='10 10 12 12')
    rec2.grid(column=1, row=1, sticky=(N, W, E, S))

    ttk.Label(rec2, text='Primary Adult: ').grid(column=1, row=1, sticky=W)
    ttk.Label(rec2, text=mstr['PrimaryFirst']+' '+mstr['PrimaryMiddle']+' '+
                         mstr['PrimaryLast']).grid(column=2, row=1,sticky=W)

    ttk.Label(rec2, text='Second Adult: ').grid(column=1, row=2, sticky=W)
    ttk.Label(rec2, text=mstr['SecondFirst']+' '+mstr['SecondMiddle']+' '+
                         mstr['SecondLast']).grid(column=2, row=2, sticky=W)

    ttk.Label(rec2, text='Address: ').grid(column=1, row=3, sticky=W)
    ttk.Label(rec2, text=mstr['Address']+' #'+mstr['Apartment']).grid(column=2, row=3, sticky=W)
    ttk.Label(rec2, text=mstr['City']+', '+mstr['Zipcode']).grid(column=2, row=4, sticky=W)

    ttk.Label(rec2, text='Email Address: ').grid(column=1, row=5, sticky=W)
    ttk.Label(rec2, text=mstr['Email']).grid(column=2, row=5, sticky=W)

    ttk.Label(rec2, text='Phone Number: ').grid(column=1, row=6, sticky=W)
    ttk.Label(rec2, text=mstr['PrimaryPhone']).grid(column=2, row=6, sticky=W)

    # SHOWS EDUCATOR INFORMATION DEPENDING ON WHETHER OR NOT IT EXISTS
    if mstr['NameOfSchool'] != empty_str:
        ttk.Label(rec2, text='Name of School: ').grid(column=1, row=7, sticky=W)
        ttk.Label(rec2, text=mstr['NameOfSchool']).grid(column=2, row=7, sticky=W)

        ttk.Label(rec2, text='District: ').grid(column=1, row=8, sticky=W)
        ttk.Label(rec2, text=mstr['DistrictNumber']).grid(column=2, row=8, sticky=W)

    return


#   MAIN    #
master = {'PrimaryFirst': '', 'PrimaryMiddle': '', 'PrimaryLast': '', 'AdultState': '', 'Address': '',
          'Apartment': '', 'City': '', 'Zipcode': '', 'Email': '', 'SecondFirst': '', 'SecondMiddle': '',
          'SecondLast': '', 'PrimaryPhone': '', 'SecondaryPhone': '', 'MembershipLevel': '', 'NumberKids': '',
          'NameOfSchool': '', 'DistrictNumber': '', 'DonationAmount': '', 'NannyState': 'False', 'NannyOrPca': '',
          'NPCAName': ''}
empty_str = ''

run_check()
find_tessy()
IDswipe()
firstbox()
find_tessy()
success = tessyfill(master)
print(success)
if success:
    pyautogui.press('esc')
    applymember1(master['PrimaryFirst'], master['PrimaryLast'], master['AdultState'], master['SecondFirst'],
                 master['SecondLast'])
    applymember2(master['MembershipLevel'], master['DonationAmount'])
    closewindow()
else:
    info_recall(master)
