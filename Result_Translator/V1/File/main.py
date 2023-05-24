import os
import pandas as pd
import xlsxwriter as xl
from datetime import datetime

current_datetime = datetime.now()


participants = []
names = []

PAN_IDs_Participants = []
Data_participants = []

selected_range_events = []

index_file = "1625"

file_directory = ""

Exer_title = "Exercise Title"
Exer_day = "Day X"
# Exer_date = "dd/mm/yyyy"
Exer_date = str(current_datetime.day) + "/" + str(current_datetime.month) + "/" + str(current_datetime.year)

NSA_logo_scale = 0.095
NSA_logo_offset_x = 0
NSA_logo_offset_y = 0

RBAT_logo_scale = 0.65
RBAT_logo_offset_x = 0
RBAT_logo_offset_y = 0


# ///////////////////////Translated
Catastrophic_kill = "ضربة قاتلة"
Heavy_damage = "دمار كثير"
Medium_damage = "دمار متواسط"
Light_damage = "N/A"
Light_wounded = "N/A"
Heavy_wounded = "N/A"
Wounded_sitting = "N/A"
Wounded_laying = "N/A"
Mobility_Damage = "معطل"
Near_miss = "N/A"
direct_fire = "N/A"
Weapon = "N/A"
pistol = "N/A"
Friendly_fire = "N/A"
Party = "فريق"
Blue_Party = "فريق أزرق"
Blue = "أزرق"
Red_Party = "فريق أحمر"
Red = "أحمر"
Company = "سرية"
Platoon = "فصيل"
Section = "حضيرة"
Soldier = "جندي"
Section_Commander = "قائد حضيرة"
True_w = "N/A"
False_w = "N/A"
Yes = "نعم"
No = "لا"
Distance = "مسافحة"
Unknown = "N/A"
ORBAT = "N/A"
Vulnerability_State = "N/A"
Originator = "N/A"
Victim = "N/A"
Suicide_Bomber = "N/A"
IED = "N/A"
Hit = "ضرب"


# /////////////////////////////////

def open_txt():
    name = "test_csv_names2.txt"
    file = open(name, encoding='utf-8')
    for i in range(100):
        line = file.readline()
        print(line)


def open_txt_AR():
    File_parsings = []
    name = "AR_names.txt"
    file = open(name, encoding='utf-8')
    for i in range(4):
        line = file.readline()
        data = line.split("\t")
        for i in range(len(data)):
            if (i == 4):
                new = data[i].replace('\n', '')
                data[4] = new
        print(line)
        print(data)
        participants.append(data)
    print(participants)


def write_txt():
    write_file = "names_file.txt"
    file_w = open(write_file, 'w+', encoding='utf-8')
    for i in range(len(participants)):
        print(participants[i])
        names_sep = []
        names_sep.append(participants[i][3])
        names_sep.append(participants[i][4])
        names.append(names_sep)
    print("*******///////////************")
    print(names)
    for l in range(len(names)):
        EN_name = str(names[l][0])
        AR_name = str(names[l][1])
        file_w.write(EN_name)
        file_w.write(" -- ")
        file_w.write(AR_name)
        file_w.write("\n")
    file_w.close()


# //////////////////////Test Cases Above


class participant():
    def __init__(self, panid, team, en_name, ar_name):
        self.PAN_ID = panid
        self.TEAM = team
        self.EN_NAME = en_name
        self.AR_NAME = ar_name

    def get_panid(self):
        return self.PAN_ID

    def get_en_name(self):
        return self.EN_NAME

    def get_ar_name(self):
        return self.AR_NAME

    def get_team(self):
        return self.TEAM


def visuals_init():
    os.system("cls")


def write_range_events_txt(directory):
    range_selected_file = open(directory + "\Filtered_txt.txt", 'w+', encoding='utf-8')
    for i in range(len(selected_range_events)):
        # print(selected_range_events[i])
        range_selected_file.writelines(selected_range_events[i])
        range_selected_file.writelines("\n")
    print(50 * "---")
    print(f"{len(selected_range_events)} events filtered ")
    range_selected_file.close()


def write_solider_events_files_initial(directory, panid):
    solider_file = open(directory + "\Data" + "\\" + str(panid) + ".txt", 'w+', encoding='utf-8')
    # solider_file.writelines("Killed: ")
    solider_file.writelines("\n")
    solider_file.close()


def write_solider_events_files(directory, panid, event, victim, FF):
    solider_file = open(directory + "\Data" + "\\" + str(panid) + ".txt", 'a', encoding='utf-8')
    if(FF == "Yes"):
        solider_file.writelines(str(event) + ": " + str(victim) + " (FF)")
    else:
        solider_file.writelines(str(event) + ": " + str(victim))
    solider_file.writelines("\n")
    solider_file.close()


def read_solider_events_files(directory, filename):
    Solider_file = open(directory + "\Data" + "\\" + str(filename) + ".txt", encoding='utf-8')
    data = Solider_file.read()
    return data


def translator(word):
    # print(Types)
    ret_word = ""
    if (word.lower() == "catastrophic kill"):
        ret_word = word + " / " + Catastrophic_kill
    elif(word.lower() == "heavy damage"):
        ret_word = word + " / " + Heavy_damage
    elif (word.lower() == "medium damage"):
        ret_word = word + " / " + Medium_damage
    elif (word.lower() == "light damage"):
        ret_word = word + " / " + Light_damage
    elif (word.lower() == "light wounded"):
        ret_word = word + " / " + Light_wounded
    elif (word.lower() == "heavy wounded"):
        ret_word = word + " / " + Heavy_wounded
    elif (word.lower() == "wounded sitting"):
        ret_word = word + " / " + Wounded_sitting
    elif (word.lower() == "wounded laying"):
        ret_word = word + " / " + Wounded_laying
    elif (word.lower() == "mobility damage"):
        ret_word = word + " / " + Mobility_Damage
    elif (word.lower() == "near miss"):
        ret_word = word + " / " + Near_miss
    elif (word.lower() == "direct fire"):
        ret_word = word + " / " + direct_fire
    elif (word.lower() == "weapon"):
        ret_word = word + " / " + Weapon
    elif (word.lower() == "pistol"):
        ret_word = word + " / " + pistol
    elif (word.lower() == "friendly fire"):
        ret_word = word + " / " + Friendly_fire
    elif (word.lower() == "party"):
        ret_word = word + " / " + Party
    elif (word.lower() == "blue"):
        ret_word = word + " / " + Blue_Party
    elif (word.lower() == "red"):
        ret_word = word + " / " + Red_Party
    elif (word.lower() == "company"):
        ret_word = word + " / " + Company
    elif (word.lower() == "platoon"):
        ret_word = word + " / " + Platoon
    elif (word.lower() == "section"):
        ret_word = word + " / " + Section
    elif (word.lower() == "soldier"):
        ret_word = word + " / " + Soldier
    elif (word.lower() == "section commander"):
        ret_word = word + " / " + Section_Commander
    elif (word.lower() == "distance"):
        ret_word = word + " / " + Distance
    elif (word.lower() == "true"):
        ret_word = word + " / " + True_w
    elif (word.lower() == "false"):
        ret_word = word + " / " + False_w
    elif (word.lower() == "yes"):
        ret_word = word + " / " + Yes
    elif (word.lower() == "no"):
        ret_word = word + " / " + No
    elif (word.lower() == "unkown"):
        ret_word = word + " / " + Unknown
    elif (word.lower() == "orbat"):
        ret_word = word + " / " + ORBAT
    elif (word.lower() == "victim"):
        ret_word = word + " / " + Victim
    elif (word.lower() == "suicide bomber"):
        ret_word = word + " / " + Suicide_Bomber
    elif (word.lower() == "ied"):
        ret_word = word + " / " + IED
    elif (word.lower() == "hit"):
        ret_word = word + " / " + Hit
    else:
        ret_word = word + " / -"
    return ret_word

def time_adjuster(time):
    try:
        time_split = time.split("T")
        Event_date = time_split[0]
        Event_time = time_split[1]
        Event_time_split = Event_time.split(":")
        Event_time_split_hr = Event_time_split[0]
        Event_time_split_hr_corrected = int(Event_time_split_hr) + 3
        Event_time = str(Event_time_split_hr_corrected) + ":" + Event_time_split[1] + ":" + Event_time_split[2]
        Event_moment = Event_date + " @ " + Event_time
        return Event_moment
    except:
        Event_moment = "Time Error"
        return Event_moment
def exc_header_V2(directory):
    xl_name = directory + "\Af_results_Header_v2.xlsx"
    filtered_text_file = open(directory + "\Filtered_txt.txt", encoding='utf-8')
    filtered_text_file = filtered_text_file.read().split("\n")
    file_lines = len(filtered_text_file)
    workbook = xl.Workbook(xl_name)
    sheet = workbook.add_worksheet()
    sheet.set_column('A:A', 8)
    sheet.set_column('B:B', 25)
    sheet.set_column('C:C', 19)
    sheet.set_column('D:D', 27)
    sheet.set_column('E:E', 17)
    sheet.set_column('F:F', 26.5)
    sheet.set_column('G:G', 17)
    sheet.set_column('H:H', 20)
    sheet.set_column('I:I', 30)
    sheet.set_column('J:J', 15)
    sheet.set_column('K:K', 25)
    sheet.set_row(0, 71)
    sheet.set_row(1, 25)
    cell_format_top = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4'})
    cell_format_top_right = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 2,
         'right': 1, 'top': 2})
    cell_format_top_middle = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 1,
         'right': 1, 'top': 2})
    cell_format_top_left = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 1,
         'right': 2, 'top': 2})
    cell_format_detail_right = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'font_size': '12', 'bottom': 1, 'left': 2, 'right': 1})
    cell_format_detail_right_No = workbook.add_format(
        {'bold': False,'align': 'center', 'font_color': 'black', 'font_size': '12', 'bottom': 1, 'left': 2, 'right': 1})
    cell_format_detail_middle = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 1, 'left': 1,
         'right': 1})
    cell_format_detail_Score_names = workbook.add_format(
        {'bold': True, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
         'font_size': '14', 'bottom': 1, 'left': 2, 'right': 1, 'top': 2})
    cell_format_detail_Score_nums = workbook.add_format(
        {'bold': True, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
         'font_size': '18', 'bottom': 2, 'left': 1, 'right': 1, 'top': 1})
    cell_format_detail_middle_wrap = workbook.add_format(
        {'bold': False, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 2,
         'left': 1, 'right': 2, 'top': 1})
    cell_format_detail_left = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 1, 'left': 1,
         'right': 2})
    cell_format_detail_middle_party_blue = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'blue', 'bottom': 1,
         'left': 1, 'right': 1})
    cell_format_detail_middle_party_red = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'red', 'bottom': 1,
         'left': 1, 'right': 1})
    cell_format_detail_middle_FF_yellow = workbook.add_format(
        {'bold': False,'text_wrap': True, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'yellow', 'bottom': 1,
         'left': 1, 'right': 1})
    cell_format_banner_seperator_right = workbook.add_format(
        {'bold': True,'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '16', 'bg_color': '#FCD5B4', 'bottom': 1,
         'left': 0, 'right': 2})
    cell_format_banner_seperator_middle = workbook.add_format(
        {'bold': True, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '16',
         'bg_color': '#FCD5B4', 'bottom': 1,
         'left': 0, 'right': 0})
    cell_format_banner_seperator_left = workbook.add_format(
        {'bold': True, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '14',
         'bg_color': '#FCD5B4', 'bottom': 1,
         'left': 2, 'right': 0})
    # //////////////////////
    cell_format_no_events_right = workbook.add_format(
        {'bold': False, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '14',
         'bg_color': '#FF7C80', 'bottom': 1,
         'left': 0, 'right': 2})
    cell_format_no_events_middle = workbook.add_format(
        {'bold': False, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '14',
         'bg_color': '#FF7C80', 'bottom': 1,
         'left': 0, 'right': 0})
    cell_format_no_events_left = workbook.add_format(
        {'bold': False, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '14',
         'bg_color': '#FF7C80', 'bottom': 1,
         'left': 2, 'right': 0})

    #//////////////////////////////////
    cell_format_no_events_GR_right = workbook.add_format(
        {'bold': False, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '14',
         'bg_color': '#00FF99', 'bottom': 1,
         'left': 0, 'right': 2})
    cell_format_no_events_GR_middle = workbook.add_format(
        {'bold': False, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '14',
         'bg_color': '#00FF99', 'bottom': 1,
         'left': 0, 'right': 0})
    cell_format_no_events_GR_left = workbook.add_format(
        {'bold': False, 'text_wrap': False, 'font_color': 'black', 'align': 'center', 'font_size': '14',
         'bg_color': '#00FF99', 'bottom': 1,
         'left': 2, 'right': 0})

    # sheet.write(0, 0, "Hit Type")
    Title = Exer_title + ": "+ Exer_day + "\n" + "        "+Exer_date
    sheet.insert_image('A1',"NSA_logo.png",  {"x_scale": NSA_logo_scale, "y_scale": NSA_logo_scale, 'x_offset': NSA_logo_offset_x, 'y_offset': NSA_logo_offset_y})
    sheet.insert_image('I1',"RBAT_logo.jpeg",  {"x_scale": RBAT_logo_scale, "y_scale": RBAT_logo_scale, 'x_offset': RBAT_logo_offset_x, 'y_offset': RBAT_logo_offset_y})
    text = Title   #"Excercise Title: Excercise \n        dd/mm/yy"
    options = {
        "x_offset": 15,
        "y_offset": 0,
        "width": 497,
        "height": 88,
        "fill": {"none": True},
        "font": {
            "bold": False,
            "italic": False,
            "name": "Calibri (Body)",
            "color": "black",
            "size": 24,
        },
        "align": {"vertical": "middle", "horizontal": "center"},
    }
    sheet.insert_textbox(0, 3, text, options)
    sheet.write(1, 0, "No.", cell_format_top_right)
    sheet.write(1, 1, "Shooter Name (الرامي)", cell_format_top_middle)
    sheet.write(1, 2, "Shooter PAN ID", cell_format_top_middle)
    sheet.write(1, 3, "Shooter Party", cell_format_top_middle)
    sheet.write(1, 4, "Status (حالة الإصابة)", cell_format_top_middle)
    sheet.write(1, 5, "Victim Name (المصاب)", cell_format_top_middle)
    sheet.write(1, 6, "Victim PAN ID", cell_format_top_middle)
    sheet.write(1, 7, "Victim Party", cell_format_top_middle)
    sheet.write(1, 8, "Time", cell_format_top_left)
    m = 2
    # ////////Friendly fire Banner
    sheet.write(m, 0, "", cell_format_banner_seperator_left)
    sheet.write(m, 1, "", cell_format_banner_seperator_middle)
    sheet.write(m, 2, "", cell_format_banner_seperator_middle)
    sheet.write(m, 3, "", cell_format_banner_seperator_middle)
    sheet.write(m, 4, "Friendly Fire (Blue > Blue)", cell_format_banner_seperator_middle)
    sheet.write(m, 5, "", cell_format_banner_seperator_middle)
    sheet.write(m, 6, "", cell_format_banner_seperator_middle)
    sheet.write(m, 7, "", cell_format_banner_seperator_middle)
    sheet.write(m, 8, "", cell_format_banner_seperator_right)
    Num_counter = 0
    Freindly_fire_events = 0
    Catastrophic_kill_events_B_R = 0
    Catastrophic_kill_events_R_B = 0
    Wounding_events_B_R = 0
    Wounding_events_R_B = 0
    Near_miss_events_B_R = 0
    Near_miss_events_R_B = 0

    # ////////////////////////////////////////////Friendly Fire Gen
    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)): #check for invalid chars in a parsed array
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        if(((parsed_event[5] == "Yes") & (parsed_event[8] != "Near miss")) & (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue") & (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue")):
            m = m + 1
            Num_counter = Num_counter + 1
            Freindly_fire_events = Freindly_fire_events + 1
            print("FoundFF")
            sheet.write(m, 0, Num_counter, cell_format_detail_right_No)  # Hit Type Call arabic method
            # sheet.write(m, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
            sheet.write(m, 2, int(parsed_event[11]), cell_format_detail_middle)
            try:
                sheet.write(m, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle)
            except:
                sheet.write(m, 3, "Party", cell_format_detail_middle)
            if (parsed_event[5] == "No"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            elif (parsed_event[5] == "Yes"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
            else:
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            # sheet.write(m, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
            sheet.write(m, 6, int(parsed_event[3]), cell_format_detail_middle)
            try:
                sheet.write(m, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
            # sheet.write(m, 8, int(parsed_event[3]), cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 7, "Party", cell_format_detail_middle)
            except:
                sheet.write(m, 7, "N/A", cell_format_detail_middle)
            # sheet.write(m, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
            Event_Times = time_adjuster(parsed_event[1])
            sheet.write(m, 8, Event_Times, cell_format_detail_left)
            write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3], parsed_event[5])
    # Freindly_fire_events = 0 # Test Code
    if(Freindly_fire_events < 1):
        m = m + 1
        sheet.write(m, 0, "", cell_format_no_events_GR_left)
        sheet.write(m, 1, "", cell_format_no_events_GR_middle)
        sheet.write(m, 2, "", cell_format_no_events_GR_middle)
        sheet.write(m, 3, "", cell_format_no_events_GR_middle)
        sheet.write(m, 4, "No Friendly Fire Events", cell_format_no_events_GR_middle)
        sheet.write(m, 5, "", cell_format_no_events_GR_middle)
        sheet.write(m, 6, "", cell_format_no_events_GR_middle)
        sheet.write(m, 7, "", cell_format_no_events_GR_middle)
        sheet.write(m, 8, "", cell_format_no_events_GR_right)

    print("FF: " + str(Freindly_fire_events))
    # /////////////////////////Cat Kill (Blue > Red)
    # ////////Catastrophic Killed Banner
    m = m + 1
    sheet.write(m, 0, "", cell_format_banner_seperator_left)
    sheet.write(m, 1, "", cell_format_banner_seperator_middle)
    sheet.write(m, 2, "", cell_format_banner_seperator_middle)
    sheet.write(m, 3, "", cell_format_banner_seperator_middle)
    sheet.write(m, 4, "Catastrophic Killed (ضربة قاتلة) (Blue > Red)", cell_format_banner_seperator_middle)
    sheet.write(m, 5, "", cell_format_banner_seperator_middle)
    sheet.write(m, 6, "", cell_format_banner_seperator_middle)
    sheet.write(m, 7, "", cell_format_banner_seperator_middle)
    sheet.write(m, 8, "", cell_format_banner_seperator_right)

    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)): #check for invalid chars in a parsed array
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
                if (crc == 3):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        if((parsed_event[8] == "Catastrophic kill") & (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue") & (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red")):
            m = m + 1
            Num_counter = Num_counter + 1
            Catastrophic_kill_events_B_R = Catastrophic_kill_events_B_R + 1
            print("Found CK")
            sheet.write(m, 0, Num_counter, cell_format_detail_right_No)  # Hit Type Call arabic method
            # sheet.write(m, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
            sheet.write(m, 2, int(parsed_event[11]), cell_format_detail_middle)
            try:
                sheet.write(m, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle)
            except:
                sheet.write(m, 3, "Party", cell_format_detail_middle)
            if (parsed_event[5] == "No"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            elif (parsed_event[5] == "Yes"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
            else:
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            # sheet.write(m, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
            sheet.write(m, 6, int(parsed_event[3]), cell_format_detail_middle)
            try:
                sheet.write(m, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
            # sheet.write(m, 8, int(parsed_event[3]), cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 7, "Party", cell_format_detail_middle)
            except:
                sheet.write(m, 7, "N/A", cell_format_detail_middle)
            # sheet.write(m, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
            Event_Times = time_adjuster(parsed_event[1])
            sheet.write(m, 8, Event_Times, cell_format_detail_left)
            write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3], parsed_event[5])
    # Freindly_fire_events = 0 # Test Code
    if(Catastrophic_kill_events_B_R < 1):
        m = m + 1
        sheet.write(m, 0, "", cell_format_no_events_left)
        sheet.write(m, 1, "", cell_format_no_events_middle)
        sheet.write(m, 2, "", cell_format_no_events_middle)
        sheet.write(m, 3, "", cell_format_no_events_middle)
        sheet.write(m, 4, "No Catastrophic Killed Events", cell_format_no_events_middle)
        sheet.write(m, 5, "", cell_format_no_events_middle)
        sheet.write(m, 6, "", cell_format_no_events_middle)
        sheet.write(m, 7, "", cell_format_no_events_middle)
        sheet.write(m, 8, "", cell_format_no_events_right)
    print("CK_BR : " + str(Catastrophic_kill_events_B_R))
    # /////////////////////////Cat Kill (Red > Blue)
    # ////////Catastrophic Killed Banner
    m = m + 1
    sheet.write(m, 0, "", cell_format_banner_seperator_left)
    sheet.write(m, 1, "", cell_format_banner_seperator_middle)
    sheet.write(m, 2, "", cell_format_banner_seperator_middle)
    sheet.write(m, 3, "", cell_format_banner_seperator_middle)
    sheet.write(m, 4, "Catastrophic Killed (ضربة قاتلة) (Red > Blue)", cell_format_banner_seperator_middle)
    sheet.write(m, 5, "", cell_format_banner_seperator_middle)
    sheet.write(m, 6, "", cell_format_banner_seperator_middle)
    sheet.write(m, 7, "", cell_format_banner_seperator_middle)
    sheet.write(m, 8, "", cell_format_banner_seperator_right)
    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)): #check for invalid chars in a parsed array
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
                if (crc == 3):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        if((parsed_event[8] == "Catastrophic kill") & (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red") & (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue")):
            m = m + 1
            Num_counter = Num_counter + 1
            Catastrophic_kill_events_R_B = Catastrophic_kill_events_R_B + 1
            print("Found CK")
            sheet.write(m, 0, Num_counter, cell_format_detail_right_No)  # Hit Type Call arabic method
            # sheet.write(m, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
            sheet.write(m, 2, int(parsed_event[11]), cell_format_detail_middle)
            try:
                sheet.write(m, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle)
            except:
                sheet.write(m, 3, "Party", cell_format_detail_middle)
            if (parsed_event[5] == "No"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            elif (parsed_event[5] == "Yes"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
            else:
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            # sheet.write(m, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
            sheet.write(m, 6, int(parsed_event[3]), cell_format_detail_middle)
            try:
                sheet.write(m, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
            # sheet.write(m, 8, int(parsed_event[3]), cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 7, "Party", cell_format_detail_middle)
            except:
                sheet.write(m, 7, "N/A", cell_format_detail_middle)
            # sheet.write(m, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
            Event_Times = time_adjuster(parsed_event[1])
            sheet.write(m, 8, Event_Times, cell_format_detail_left)
            write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3], parsed_event[5])
    if (Catastrophic_kill_events_R_B < 1):
        m = m + 1
        sheet.write(m, 0, "", cell_format_no_events_GR_left)
        sheet.write(m, 1, "", cell_format_no_events_GR_middle)
        sheet.write(m, 2, "", cell_format_no_events_GR_middle)
        sheet.write(m, 3, "", cell_format_no_events_GR_middle)
        sheet.write(m, 4, "No Catastrophic Killed Events", cell_format_no_events_GR_middle)
        sheet.write(m, 5, "", cell_format_no_events_GR_middle)
        sheet.write(m, 6, "", cell_format_no_events_GR_middle)
        sheet.write(m, 7, "", cell_format_no_events_GR_middle)
        sheet.write(m, 8, "", cell_format_no_events_GR_right)
    print("CK_RB : " + str(Catastrophic_kill_events_R_B))

    # /////////////////////////Wounding (Blue > Red)
    # ////////Wounding (Blue > Red)
    m = m + 1
    sheet.write(m, 0, "", cell_format_banner_seperator_left)
    sheet.write(m, 1, "", cell_format_banner_seperator_middle)
    sheet.write(m, 2, "", cell_format_banner_seperator_middle)
    sheet.write(m, 3, "", cell_format_banner_seperator_middle)
    sheet.write(m, 4, "Wounding (Blue > Red)", cell_format_banner_seperator_middle)
    sheet.write(m, 5, "", cell_format_banner_seperator_middle)
    sheet.write(m, 6, "", cell_format_banner_seperator_middle)
    sheet.write(m, 7, "", cell_format_banner_seperator_middle)
    sheet.write(m, 8, "", cell_format_banner_seperator_right)
    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)): #check for invalid chars in a parsed array
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
                if (crc == 3):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        if(((parsed_event[8] == "Heavy damage") or (parsed_event[8] == "Medium damage") or (parsed_event[8] == "Light damage") ) & (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue") & (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red")):
            m = m + 1
            Num_counter = Num_counter + 1
            Wounding_events_B_R = Wounding_events_B_R + 1
            print("Found CK")
            sheet.write(m, 0, Num_counter, cell_format_detail_right_No)  # Hit Type Call arabic method
            # sheet.write(m, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
            sheet.write(m, 2, int(parsed_event[11]), cell_format_detail_middle)
            try:
                sheet.write(m, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle)
            except:
                sheet.write(m, 3, "Party", cell_format_detail_middle)
            if (parsed_event[5] == "No"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            elif (parsed_event[5] == "Yes"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
            else:
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            # sheet.write(m, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
            sheet.write(m, 6, int(parsed_event[3]), cell_format_detail_middle)
            try:
                sheet.write(m, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
            # sheet.write(m, 8, int(parsed_event[3]), cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 7, "Party", cell_format_detail_middle)
            except:
                sheet.write(m, 7, "N/A", cell_format_detail_middle)
            # sheet.write(m, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
            Event_Times = time_adjuster(parsed_event[1])
            sheet.write(m, 8, Event_Times, cell_format_detail_left)
            write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3], parsed_event[5])
    if (Wounding_events_B_R < 1):
        m = m + 1
        sheet.write(m, 0, "", cell_format_no_events_left)
        sheet.write(m, 1, "", cell_format_no_events_middle)
        sheet.write(m, 2, "", cell_format_no_events_middle)
        sheet.write(m, 3, "", cell_format_no_events_middle)
        sheet.write(m, 4, "No Wounding Events", cell_format_no_events_middle)
        sheet.write(m, 5, "", cell_format_no_events_middle)
        sheet.write(m, 6, "", cell_format_no_events_middle)
        sheet.write(m, 7, "", cell_format_no_events_middle)
        sheet.write(m, 8, "", cell_format_no_events_right)
    print("W_BR : " + str(Wounding_events_B_R))

    # /////////////////////////Wounding (Red > Blue)
    # ////////Wounding (Red > Blue)
    m = m + 1
    sheet.write(m, 0, "", cell_format_banner_seperator_left)
    sheet.write(m, 1, "", cell_format_banner_seperator_middle)
    sheet.write(m, 2, "", cell_format_banner_seperator_middle)
    sheet.write(m, 3, "", cell_format_banner_seperator_middle)
    sheet.write(m, 4, "Wounding (Red > Blue)", cell_format_banner_seperator_middle)
    sheet.write(m, 5, "", cell_format_banner_seperator_middle)
    sheet.write(m, 6, "", cell_format_banner_seperator_middle)
    sheet.write(m, 7, "", cell_format_banner_seperator_middle)
    sheet.write(m, 8, "", cell_format_banner_seperator_right)
    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)): #check for invalid chars in a parsed array
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
                if (crc == 3):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        if(((parsed_event[8] == "Heavy damage") or (parsed_event[8] == "Medium damage") or (parsed_event[8] == "Light damage") ) & (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red") & (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue")):
            m = m + 1
            Num_counter = Num_counter + 1
            Wounding_events_R_B = Wounding_events_R_B + 1
            print("Found CK")
            sheet.write(m, 0, Num_counter, cell_format_detail_right_No)  # Hit Type Call arabic method
            # sheet.write(m, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
            sheet.write(m, 2, int(parsed_event[11]), cell_format_detail_middle)
            try:
                sheet.write(m, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle)
            except:
                sheet.write(m, 3, "Party", cell_format_detail_middle)
            if (parsed_event[5] == "No"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            elif (parsed_event[5] == "Yes"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
            else:
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            # sheet.write(m, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
            sheet.write(m, 6, int(parsed_event[3]), cell_format_detail_middle)
            try:
                sheet.write(m, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
            # sheet.write(m, 8, int(parsed_event[3]), cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 7, "Party", cell_format_detail_middle)
            except:
                sheet.write(m, 7, "N/A", cell_format_detail_middle)
            # sheet.write(m, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
            Event_Times = time_adjuster(parsed_event[1])
            sheet.write(m, 8, Event_Times, cell_format_detail_left)
            write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3], parsed_event[5])
    if (Wounding_events_R_B < 1):
        m = m + 1
        sheet.write(m, 0, "", cell_format_no_events_GR_left)
        sheet.write(m, 1, "", cell_format_no_events_GR_middle)
        sheet.write(m, 2, "", cell_format_no_events_GR_middle)
        sheet.write(m, 3, "", cell_format_no_events_GR_middle)
        sheet.write(m, 4, "No Wounding Events", cell_format_no_events_GR_middle)
        sheet.write(m, 5, "", cell_format_no_events_GR_middle)
        sheet.write(m, 6, "", cell_format_no_events_GR_middle)
        sheet.write(m, 7, "", cell_format_no_events_GR_middle)
        sheet.write(m, 8, "", cell_format_no_events_GR_right)
    print("W_RB : " + str(Wounding_events_R_B))

    # /////////////////////////Near Miss (Blue > Red)
    # ////////Near Miss Banner
    m = m + 1
    sheet.write(m, 0, "", cell_format_banner_seperator_left)
    sheet.write(m, 1, "", cell_format_banner_seperator_middle)
    sheet.write(m, 2, "", cell_format_banner_seperator_middle)
    sheet.write(m, 3, "", cell_format_banner_seperator_middle)
    sheet.write(m, 4, "Near Miss (Blue > Red)", cell_format_banner_seperator_middle)
    sheet.write(m, 5, "", cell_format_banner_seperator_middle)
    sheet.write(m, 6, "", cell_format_banner_seperator_middle)
    sheet.write(m, 7, "", cell_format_banner_seperator_middle)
    sheet.write(m, 8, "", cell_format_banner_seperator_right)

    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)): #check for invalid chars in a parsed array
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
                if (crc == 3):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        if((parsed_event[8] == "Near miss") & (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue") & (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red")):
            m = m + 1
            Num_counter = Num_counter + 1
            Near_miss_events_B_R = Near_miss_events_B_R + 1
            print("Found CK")
            sheet.write(m, 0, Num_counter, cell_format_detail_right_No)  # Hit Type Call arabic method
            # sheet.write(m, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
            sheet.write(m, 2, int(parsed_event[11]), cell_format_detail_middle)
            try:
                sheet.write(m, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle)
            except:
                sheet.write(m, 3, "Party", cell_format_detail_middle)
            if (parsed_event[5] == "No"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            elif (parsed_event[5] == "Yes"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
            else:
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            # sheet.write(m, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
            sheet.write(m, 6, int(parsed_event[3]), cell_format_detail_middle)
            try:
                sheet.write(m, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
            # sheet.write(m, 8, int(parsed_event[3]), cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 7, "Party", cell_format_detail_middle)
            except:
                sheet.write(m, 7, "N/A", cell_format_detail_middle)
            # sheet.write(m, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
            Event_Times = time_adjuster(parsed_event[1])
            sheet.write(m, 8, Event_Times, cell_format_detail_left)
            write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3], parsed_event[5])
    # Freindly_fire_events = 0 # Test Code
    if(Near_miss_events_B_R < 1):
        m = m + 1
        sheet.write(m, 0, "", cell_format_no_events_left)
        sheet.write(m, 1, "", cell_format_no_events_middle)
        sheet.write(m, 2, "", cell_format_no_events_middle)
        sheet.write(m, 3, "", cell_format_no_events_middle)
        sheet.write(m, 4, "No Near Miss Events", cell_format_no_events_middle)
        sheet.write(m, 5, "", cell_format_no_events_middle)
        sheet.write(m, 6, "", cell_format_no_events_middle)
        sheet.write(m, 7, "", cell_format_no_events_middle)
        sheet.write(m, 8, "", cell_format_no_events_right)
    print("NM_BR : " + str(Near_miss_events_B_R))


    # /////////////////////////Near Miss (Red > Blue)
    # ////////Near Miss Banner
    m = m + 1
    sheet.write(m, 0, "", cell_format_banner_seperator_left)
    sheet.write(m, 1, "", cell_format_banner_seperator_middle)
    sheet.write(m, 2, "", cell_format_banner_seperator_middle)
    sheet.write(m, 3, "", cell_format_banner_seperator_middle)
    sheet.write(m, 4, "Near Miss (Red > Blue)", cell_format_banner_seperator_middle)
    sheet.write(m, 5, "", cell_format_banner_seperator_middle)
    sheet.write(m, 6, "", cell_format_banner_seperator_middle)
    sheet.write(m, 7, "", cell_format_banner_seperator_middle)
    sheet.write(m, 8, "", cell_format_banner_seperator_right)

    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)): #check for invalid chars in a parsed array
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
                if (crc == 3):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        if((parsed_event[8] == "Near miss") & (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red") & (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue")):
            m = m + 1
            Num_counter = Num_counter + 1
            Near_miss_events_R_B = Near_miss_events_R_B + 1
            print("Found CK")
            sheet.write(m, 0, Num_counter, cell_format_detail_right_No)  # Hit Type Call arabic method
            # sheet.write(m, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
            sheet.write(m, 2, int(parsed_event[11]), cell_format_detail_middle)
            try:
                sheet.write(m, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                                cell_format_detail_middle)
            except:
                sheet.write(m, 3, "Party", cell_format_detail_middle)
            if (parsed_event[5] == "No"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            elif (parsed_event[5] == "Yes"):
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
            else:
                sheet.write(m, 4, str(hit_typ), cell_format_detail_middle)
            # sheet.write(m, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
            sheet.write(m, 6, int(parsed_event[3]), cell_format_detail_middle)
            try:
                sheet.write(m, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                    get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
            except:
                sheet.write(m, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
            # sheet.write(m, 8, int(parsed_event[3]), cell_format_detail_middle)
            try:
                if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_blue)
                elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                        get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                    sheet.write(m, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                                cell_format_detail_middle_party_red)
                else:
                    sheet.write(m, 7, "Party", cell_format_detail_middle)
            except:
                sheet.write(m, 7, "N/A", cell_format_detail_middle)
            # sheet.write(m, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
            Event_Times = time_adjuster(parsed_event[1])
            sheet.write(m, 8, Event_Times, cell_format_detail_left)
            write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3], parsed_event[5])
    # Freindly_fire_events = 0 # Test Code
    if(Near_miss_events_R_B < 1):
        m = m + 1
        sheet.write(m, 0, "", cell_format_no_events_GR_left)
        sheet.write(m, 1, "", cell_format_no_events_GR_middle)
        sheet.write(m, 2, "", cell_format_no_events_GR_middle)
        sheet.write(m, 3, "", cell_format_no_events_GR_middle)
        sheet.write(m, 4, "No Near Miss Events", cell_format_no_events_GR_middle)
        sheet.write(m, 5, "", cell_format_no_events_GR_middle)
        sheet.write(m, 6, "", cell_format_no_events_GR_middle)
        sheet.write(m, 7, "", cell_format_no_events_GR_middle)
        sheet.write(m, 8, "", cell_format_no_events_GR_right)
    print("NM_RB : " + str(Near_miss_events_R_B))

    sheet.write((m + 6), 1, "Solider ", cell_format_top_right)
    sheet.write((m + 6), 3, "Hit events ", cell_format_top_left)
    sheet.write((m + 6), 2, "Total Hits ", cell_format_top_middle)
    line_st = m + 6
    print(line_st)
    for pan in range(len(Data_participants)):
        # print(Data_participants[pan].get_panid())
        if (Data_participants[pan].get_panid() != "PAN ID"):
            events = read_solider_events_files(directory, Data_participants[pan].get_panid())
            events_c = events.split(':')
            if(len(events_c) - 1 > 0):
                line_st = line_st + 1
                sheet.write(line_st, 1,
                            Data_participants[pan].get_en_name() + " (" + Data_participants[pan].get_ar_name() + ")",
                            cell_format_detail_Score_names)
                sheet.write(line_st, 3, events, cell_format_detail_middle_wrap)
                sheet.write(line_st, 2, len(events_c) - 1, cell_format_detail_Score_nums)
                print(line_st)


            # print(events)
    sheet.autofit()
    workbook.close()
    print("Excel File Generated..")
def exc_header(directory):
    xl_name = directory + "\Af_results_Header.xlsx"
    filtered_text_file = open(directory + "\Filtered_txt.txt", encoding='utf-8')
    filtered_text_file = filtered_text_file.read().split("\n")
    file_lines = len(filtered_text_file)
    workbook = xl.Workbook(xl_name)
    sheet = workbook.add_worksheet()
    sheet.set_column('A:A', 8)
    sheet.set_column('B:B', 25)
    sheet.set_column('C:C', 19)
    sheet.set_column('D:D', 23)
    sheet.set_column('E:E', 17)
    sheet.set_column('F:F', 26.5)
    sheet.set_column('G:G', 17)
    sheet.set_column('H:H', 20)
    sheet.set_column('I:I', 30)
    sheet.set_column('J:J', 15)
    sheet.set_column('K:K', 25)
    sheet.set_row(0, 71)
    sheet.set_row(1, 25)
    cell_format_top = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4'})
    cell_format_top_right = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 2,
         'right': 1, 'top': 2})
    cell_format_top_middle = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 1,
         'right': 1, 'top': 2})
    cell_format_top_left = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 1,
         'right': 2, 'top': 2})
    cell_format_detail_right = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'font_size': '12', 'bottom': 1, 'left': 2, 'right': 1})
    cell_format_detail_right_No = workbook.add_format(
        {'bold': False,'align': 'center', 'font_color': 'black', 'font_size': '12', 'bottom': 1, 'left': 2, 'right': 1})
    cell_format_detail_middle = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 1, 'left': 1,
         'right': 1})
    cell_format_detail_Score_names = workbook.add_format(
        {'bold': True, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
         'font_size': '14', 'bottom': 1, 'left': 2, 'right': 1, 'top': 2})
    cell_format_detail_Score_nums = workbook.add_format(
        {'bold': True, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
         'font_size': '18', 'bottom': 2, 'left': 1, 'right': 1, 'top': 1})
    cell_format_detail_middle_wrap = workbook.add_format(
        {'bold': False, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 2,
         'left': 1, 'right': 2, 'top': 1})
    cell_format_detail_left = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 1, 'left': 1,
         'right': 2})
    cell_format_detail_middle_party_blue = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'blue', 'bottom': 1,
         'left': 1, 'right': 1})
    cell_format_detail_middle_party_red = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'red', 'bottom': 1,
         'left': 1, 'right': 1})
    cell_format_detail_middle_FF_yellow = workbook.add_format(
        {'bold': False,'text_wrap': True, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'yellow', 'bottom': 1,
         'left': 1, 'right': 1})
    # sheet.write(0, 0, "Hit Type")
    Title = Exer_title + ": "+ Exer_day + "\n" + "        "+Exer_date
    sheet.insert_image('A1',"NSA_logo.png",  {"x_scale": NSA_logo_scale, "y_scale": NSA_logo_scale, 'x_offset': NSA_logo_offset_x, 'y_offset': NSA_logo_offset_y})
    sheet.insert_image('I1',"RBAT_logo.jpeg",  {"x_scale": RBAT_logo_scale, "y_scale": RBAT_logo_scale, 'x_offset': RBAT_logo_offset_x, 'y_offset': RBAT_logo_offset_y})
    text = Title   #"Excercise Title: Excercise \n        dd/mm/yy"
    options = {
        "x_offset": 15,
        "y_offset": 0,
        "width": 497,
        "height": 88,
        "fill": {"none": True},
        "font": {
            "bold": False,
            "italic": False,
            "name": "Calibri (Body)",
            "color": "black",
            "size": 24,
        },
        "align": {"vertical": "middle", "horizontal": "center"},
    }
    sheet.insert_textbox(0, 3, text, options)
    sheet.write(1, 0, "No.", cell_format_top_right)
    sheet.write(1, 1, "Shooter Name (الرامي)", cell_format_top_middle)
    sheet.write(1, 2, "Shooter PAN ID", cell_format_top_middle)
    sheet.write(1, 3, "Shooter Party", cell_format_top_middle)
    sheet.write(1, 4, "Status (حالة الإصابة)", cell_format_top_middle)
    sheet.write(1, 5, "Victim Name (المصاب)", cell_format_top_middle)
    sheet.write(1, 6, "Victim PAN ID", cell_format_top_middle)
    sheet.write(1, 7, "Victim Party", cell_format_top_middle)
    sheet.write(1, 8, "Time", cell_format_top_left)
    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)):
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        sheet.write(d + 2, 0, d+1, cell_format_detail_right_No)  # Hit Type Call arabic method
        # sheet.write(d + 2, 4, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method *****
        sheet.write(d + 2, 2, int(parsed_event[11]), cell_format_detail_middle)
        try:
            sheet.write(d + 2, 1, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
        except:
            sheet.write(d + 2, 1, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
        try:
            if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                    get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                sheet.write(d + 2, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                            cell_format_detail_middle_party_blue)
            elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                    get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                sheet.write(d + 2, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                            cell_format_detail_middle_party_red)
            else:
                sheet.write(d + 2, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                            cell_format_detail_middle)
        except:
            sheet.write(d + 2, 3, "Party", cell_format_detail_middle)
        if (parsed_event[5] == "No"):
            sheet.write(d + 2, 4, str(hit_typ), cell_format_detail_middle)
        elif (parsed_event[5] == "Yes"):
            sheet.write(d + 2, 4, str(hit_typ), cell_format_detail_middle_FF_yellow)
        else:
            sheet.write(d + 2, 4, str(hit_typ), cell_format_detail_middle)
        # sheet.write(d + 2, 11, parsed_event[4], cell_format_detail_middle) #re locate # Weapon data
        sheet.write(d + 2, 6, int(parsed_event[3]), cell_format_detail_middle)
        try:
            sheet.write(d + 2, 5, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
        except:
            sheet.write(d + 2, 5, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
        # sheet.write(d + 2, 8, int(parsed_event[3]), cell_format_detail_middle)
        try:
            if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                    get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                sheet.write(d + 2, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                            cell_format_detail_middle_party_blue)
            elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                    get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                sheet.write(d + 2, 7, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                            cell_format_detail_middle_party_red)
            else:
                sheet.write(d + 2, 7, "Party", cell_format_detail_middle)
        except:
            sheet.write(d + 2, 7, "N/A", cell_format_detail_middle)
        # sheet.write(d + 2, 9, parsed_event[5], cell_format_detail_middle) # Friendly fire Data
        Event_Times = time_adjuster(parsed_event[1])
        sheet.write(d + 2, 8, Event_Times, cell_format_detail_left)
        write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3])
    sheet.write((file_lines + 5), 1, "Solider ", cell_format_top_right)
    sheet.write((file_lines + 5), 3, "Hit events ", cell_format_top_left)
    sheet.write((file_lines + 5), 2, "Total Hits ", cell_format_top_middle)
    line_st = file_lines + 5
    print(line_st)
    for pan in range(len(Data_participants)):
        # print(Data_participants[pan].get_panid())
        if (Data_participants[pan].get_panid() != "PAN ID"):
            events = read_solider_events_files(directory, Data_participants[pan].get_panid())
            events_c = events.split(':')
            if(len(events_c) - 1 > 0):
                line_st = line_st + 1
                sheet.write(line_st, 1,
                            Data_participants[pan].get_en_name() + " (" + Data_participants[pan].get_ar_name() + ")",
                            cell_format_detail_Score_names)
                sheet.write(line_st, 3, events, cell_format_detail_middle_wrap)
                sheet.write(line_st, 2, len(events_c) - 1, cell_format_detail_Score_nums)
                print(line_st)


            # print(events)
    sheet.autofit()
    workbook.close()
    print("Excel File Generated..")


def excel_write_after_action(directory):
    xl_name = directory + "\Af_results.xlsx"
    # workbook = xl.Workbook("AF.xlsx")
    filtered_text_file = open(directory + "\Filtered_txt.txt", encoding='utf-8')
    filtered_text_file = filtered_text_file.read().split("\n")
    file_lines = len(filtered_text_file)
    # filtered_data = filtered_text_file[0]
    # print("Fild")
    # print(filtered_data)
    # sample_name = get_participant_data(index_file, "AR_NAME")
    workbook = xl.Workbook(xl_name)
    sheet = workbook.add_worksheet()
    sheet.set_column('A:A', 15)
    sheet.set_column('B:B', 25)
    sheet.set_column('C:C', 15)
    sheet.set_column('D:D', 14)
    sheet.set_column('E:E', 5)
    sheet.set_column('F:F', 26.5)
    sheet.set_column('G:G', 17)
    sheet.set_column('H:H', 16)
    sheet.set_column('I:I', 15)
    sheet.set_column('J:J', 15)
    sheet.set_column('K:K', 25)
    cell_format_top = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4'})
    cell_format_top_right = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 2,
         'right': 1, 'top': 2})
    cell_format_top_middle = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 1,
         'right': 1, 'top': 2})
    cell_format_top_left = workbook.add_format(
        {'bold': True, 'font_color': 'black', 'font_size': '14', 'bg_color': '#72E5F4', 'bottom': 2, 'left': 1,
         'right': 2, 'top': 2})
    cell_format_detail_right = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'font_size': '12', 'bottom': 1, 'left': 2, 'right': 1})
    cell_format_detail_middle = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 1, 'left': 1,
         'right': 1})
    cell_format_detail_Score_names = workbook.add_format(
        {'bold': True, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
         'font_size': '14', 'bottom': 1, 'left': 1, 'right': 1, 'top': 2})
    cell_format_detail_Score_nums = workbook.add_format(
        {'bold': True, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
         'font_size': '18', 'bottom': 2, 'left': 1, 'right': 1, 'top': 1})
    cell_format_detail_middle_wrap = workbook.add_format(
        {'bold': False, 'text_wrap': True, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 2,
         'left': 1, 'right': 1, 'top': 2})
    cell_format_detail_left = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bottom': 1, 'left': 1,
         'right': 2})
    cell_format_detail_middle_party_blue = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'blue', 'bottom': 1,
         'left': 1, 'right': 1})
    cell_format_detail_middle_party_red = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'red', 'bottom': 1,
         'left': 1, 'right': 1})
    cell_format_detail_middle_FF_yellow = workbook.add_format(
        {'bold': False, 'font_color': 'black', 'align': 'center', 'font_size': '12', 'bg_color': 'yellow', 'bottom': 1,
         'left': 1, 'right': 1})
    sheet.write(0, 0, "Hit Type", cell_format_top_right)
    sheet.write(0, 1, "Orgin PAN ID", cell_format_top_middle)
    sheet.write(0, 2, "Orgin Name", cell_format_top_middle)
    sheet.write(0, 3, "Orgin Party", cell_format_top_middle)
    sheet.write(0, 4, ">>>", cell_format_top_middle)
    sheet.write(0, 5, "Charge / Weapon type", cell_format_top_middle)
    sheet.write(0, 6, "Victim PAN ID", cell_format_top_middle)
    sheet.write(0, 7, "Victim NAME", cell_format_top_middle)
    sheet.write(0, 8, "Victim Party", cell_format_top_middle)
    sheet.write(0, 9, "Friendly Fire", cell_format_top_middle)
    sheet.write(0, 10, "Date Time", cell_format_top_left)
    for d in range(file_lines - 1):
        parsed_event = filtered_text_file[d].split(";")
        # print(parsed_event)
        for crc in range(len(parsed_event)):
            if (parsed_event[crc] == ""):
                if (crc == 11):
                    parsed_event[crc] = 0
                else:
                    parsed_event[crc] = "N/A"
            # print(parsed_event[crc])
            # print("parsed_event[crc]")
            hit_typ = translator(parsed_event[8])
        sheet.write(d + 1, 0, str(hit_typ), cell_format_detail_right)  # Hit Type Call arabic method
        sheet.write(d + 1, 1, int(parsed_event[11]), cell_format_detail_middle)
        try:
            sheet.write(d + 1, 2, str(str(get_participant_data(str(parsed_event[11]), "EN_NAME")) + " (" + str(
                get_participant_data(str(parsed_event[11]), "AR_NAME")) + ")"), cell_format_detail_middle)
        except:
            sheet.write(d + 1, 2, "Solider" + " (" + "Not in ORBAT" + ")", cell_format_detail_middle)
        try:
            if (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Blue" or str(
                    get_participant_data(str(parsed_event[11]), "TEAM")) == "BLUE"):
                sheet.write(d + 1, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                            cell_format_detail_middle_party_blue)
            elif (str(get_participant_data(str(parsed_event[11]), "TEAM")) == "Red" or str(
                    get_participant_data(str(parsed_event[11]), "TEAM")) == "RED"):
                sheet.write(d + 1, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                            cell_format_detail_middle_party_red)
            else:
                sheet.write(d + 1, 3, translator(str(get_participant_data(str(parsed_event[11]), "TEAM"))),
                            cell_format_detail_middle)
        except:
            sheet.write(d + 1, 3, "Party", cell_format_detail_middle)
        if (parsed_event[5] == "No"):
            sheet.write(d + 1, 4, ">>>", cell_format_detail_middle)
        elif (parsed_event[5] == "Yes"):
            sheet.write(d + 1, 4, ">>>", cell_format_detail_middle_FF_yellow)
        else:
            sheet.write(d + 1, 4, ">>>", cell_format_detail_middle)
        sheet.write(d + 1, 5, parsed_event[4], cell_format_detail_middle)
        sheet.write(d + 1, 6, int(parsed_event[3]), cell_format_detail_middle)
        try:
            sheet.write(d + 1, 7, str(str(get_participant_data(str(parsed_event[3]), "EN_NAME")) + " (" + str(
                get_participant_data(str(parsed_event[3]), "AR_NAME")) + ")"), cell_format_detail_middle)
        except:
            sheet.write(d + 1, 7, "Solider" + " (" + "N/A" + ")", cell_format_detail_middle)
        # sheet.write(d + 1, 8, int(parsed_event[3]), cell_format_detail_middle)
        try:
            if (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Blue" or str(
                    get_participant_data(str(parsed_event[3]), "TEAM")) == "BLUE"):
                sheet.write(d + 1, 8, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                            cell_format_detail_middle_party_blue)
            elif (str(get_participant_data(str(parsed_event[3]), "TEAM")) == "Red" or str(
                    get_participant_data(str(parsed_event[3]), "TEAM")) == "RED"):
                sheet.write(d + 1, 8, translator(str(get_participant_data(str(parsed_event[3]), "TEAM"))),
                            cell_format_detail_middle_party_red)
            else:
                sheet.write(d + 1, 8, "Party", cell_format_detail_middle)
        except:
            sheet.write(d + 1, 8, "N/A", cell_format_detail_middle)
        sheet.write(d + 1, 9, parsed_event[5], cell_format_detail_middle)
        sheet.write(d + 1, 10, parsed_event[1], cell_format_detail_left)
        write_solider_events_files(directory, parsed_event[11], parsed_event[8], parsed_event[3])
    sheet.write((file_lines + 5), 0, "Solider ", cell_format_top_right)
    sheet.write((file_lines + 5), 1, "Hit events ", cell_format_top_middle)
    sheet.write((file_lines + 5), 2, "Total Hits ", cell_format_top_left)
    for pan in range(len(Data_participants)):
        # print(Data_participants[pan].get_panid())
        if (Data_participants[pan].get_panid() != "PAN ID"):
            events = read_solider_events_files(directory, Data_participants[pan].get_panid())
            sheet.write((file_lines + 5) + pan, 0,
                        Data_participants[pan].get_en_name() + " (" + Data_participants[pan].get_ar_name() + ")",
                        cell_format_detail_Score_names)
            sheet.write((file_lines + 5) + pan, 1, events, cell_format_detail_middle_wrap)
            events_c = events.split(':')
            sheet.write((file_lines + 5) + pan, 2, len(events_c) - 1, cell_format_detail_Score_nums)
            # print(events)
    print("Excel File Generated..")
    sheet.autofit()
    workbook.close()


def get_Exercise_Title():
    global Exer_title

    try:
        Exer_title = str(input("Please input Exercise Name (e.g. ABMMC, Battlcamp, et. ) : "))
    except:
        print("Wrong input type..")
        Exer_title = "Exercise"
def get_Exercise_day():
    global Exer_day
    try:
        Exer_day = str(input("Please input Exercise Day (e.g. Day 1, Final , etc. ) : "))
    except:
        Exer_day = "Day 0"
        print("Wrong Input type...")

def get_csv_events(directory):
    csv_event_file_name = directory + "\CSV\csv_events.txt"
    csv_event_file = open(csv_event_file_name, encoding='utf-8')
    data_full = csv_event_file.read()
    # print(data_full)
    data_lines = data_full.split("\n")
    for l in range(len(data_lines)):
        # print(str(l) + ": " + str(data_lines[l]))
        data_elements = data_lines[l].split(";")
        for m in range(len(data_elements)):
            if (data_elements[m] == ''):
                data_elements[m] = "N/A"
        # print(str(l) + " : " + str(data_elements))
        print(str(l) + " : " + str(data_elements[1]) + " - " + str(data_elements[11]) + " >>> " + str(
            data_elements[8]) + ">>>" + str(data_elements[3]) + " with -> " + str(data_elements[4]) + " : " + str(
            data_elements[0]))
        # data_elements_date = data_elements[0].split(";")
        # print(data_elements_date)
    # for l in range(100):
    #     data = csv_event_file.readline()
    #     print(data)
    start_range = int(input("Enter the starting line (HIGHER): "))
    end_range = int(input("Enter the ending line (LOWER): "))
    for s in range(start_range - (end_range - 1)):
        element = start_range - s
        # print(element)
        selected_range_events.append(data_lines[element])
    print(50 * "/")
    # for m in range(len(selected_range_events)):
    #     print(selected_range_events[m])


def get_Participant_IDs(directory):
    invalid_flag = False
    participant_file_name = directory + "\ORBAT\Participants.txt"
    participant_file = open(participant_file_name, encoding='utf-8')
    data = participant_file.read().split('\n')
    # print(data)
    print(50 * "*" + "Participants" + 50 * "*")
    print("\n")
    for l in range(len(data) - 1):
        parsed = data[l].split('\t')
        parsed_panid = parsed[1]
        if (parsed_panid == ""):
            invalid_flag = True
        if (parsed_panid != "PAN ID"):
            write_solider_events_files_initial(directory, parsed_panid)
        PAN_IDs_Participants.append(parsed_panid)
        Data_participants.append(parsed_panid)
        parsed_Team = parsed[2]
        parsed_EN_name = parsed[3]
        parsed_AR_name = parsed[4]
        # PAN_IDs_Participants[l] = participant(parsed_panid, parsed_Team, parsed_EN_name, parsed_AR_name)
        Data_participants[l] = participant(parsed_panid, parsed_Team, parsed_EN_name, parsed_AR_name)
        print(data[l])
    print("\n")
    print(50 * "*" + "=============" + 50 * "*")
    if (invalid_flag == True):
        PAN_IDs_Participants.append("0")
        Data_participants.append("0")
        Data_participants[len(data) - 1] = participant("0", "Party", "N/A", "N/A")
    print(50 * "*" + "Mem Locs" + 50 * "*")
    print("\n")
    for rec in range(len(Data_participants)):
        print(Data_participants[rec])
    print("\n")
    print(50 * "*" + "============" + 50 * "*")
    # print(parsed_panid + parsed_Team + parsed_AR_name + parsed_EN_name)

    # print(Data_participants[PAN_IDs_Participants.index(index_file)].get_en_name())
    # print(Data_participants[PAN_IDs_Participants.index(index_file)].get_ar_name())
    # print(PAN_IDs_Participants)
    participant_file.close()


def get_file_directory():
    path = str(input("Paste Working Directory: "))
    return path


def get_participant_data(panid, data_type):
    if (data_type == "EN_NAME"):
        return Data_participants[PAN_IDs_Participants.index(panid)].get_en_name()
    elif (data_type == "AR_NAME"):
        return Data_participants[PAN_IDs_Participants.index(panid)].get_ar_name()
    elif (data_type == "TEAM"):
        if(panid == "N/A" ):
            return "Party"
        else:
            try:
                return Data_participants[PAN_IDs_Participants.index(panid)].get_team()
            except:
                return "Party"
    elif (data_type == "PAN_ID"):
        return Data_participants[PAN_IDs_Participants.index(panid)].get_panid()


def input_start_date():
    while (1):
        try:
            Input_year = int(input("Enter Start Year (YYYY): "))
            if ((len(str(Input_year)) != 4) or (Input_year < 2000)):
                print("Wrong Input")
            else:
                # print(Input_year)
                break
        except:
            print("Wrong Input")
    while (1):
        try:
            Input_month = int(input("Enter Start Month (MM): "))
            if ((Input_month < 1) or (Input_month > 12)):
                print("Wrong Input")
            else:
                # print(Input_month)
                break
        except:
            print("Wrong Input")
    while (1):
        try:
            Input_day = int(input("Enter Start day (DD): "))
            if ((Input_day < 1) or (Input_day > 31)):
                print("Wrong Input")
            else:
                # print(Input_day)
                break
        except:
            print("Wrong Input")
    # print(f"Year: {Input_year} Month: {Input_month} Day: {Input_day}")
    return Input_year, Input_month, Input_day


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    visuals_init()
    file_directory = get_file_directory()
    get_Exercise_Title()
    get_Exercise_day()
    get_Participant_IDs(file_directory)
    get_csv_events(file_directory)
    write_range_events_txt(file_directory)
    #excel_write_after_action(file_directory) #Enable main excel gen
    exc_header_V2(file_directory)

    # year_st, month_st, day_st = input_start_date()
    print("//////////////////////////////////////////")
    # print(get_participant_data(index_file, "TEAM"))
    # print(get_participant_data(index_file, "AR_NAME"))

    # open_txt()
    # open_txt_AR()
    # write_txt()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
