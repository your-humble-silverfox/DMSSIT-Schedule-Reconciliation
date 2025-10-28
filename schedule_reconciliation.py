from openpyxl import load_workbook
import re
import pandas
import sys



def startup():
    print("Добро пожаловать в программу сверки расписания")
    path_schedule = input("Введите путь к таблице с расписание в формате .xlxs: ")
    path_workload = input("Введите путь к таблице с нагрузкой в формате .xls: ")
    schedule_parser(path_schedule, workload_parser(path_workload))

def workload_parser(path: str):
    workload = pandas.read_excel(path, sheet_name="По ППС", usecols=["Мероприятие реестра, норма времени", "Вид потока", "План. поток", "ППС"], skiprows=1)
    #print(workload)
    workload["Мероприятие реестра, норма времени"] = workload["Мероприятие реестра, норма времени"].astype(str).str.strip()
    workload["Вид потока"] = workload["Вид потока"].str.strip()
    return workload

def num_to_column(num: int):
    column = ""
    while (num > 0):
        num -= 1
        column = chr( num % 26 + ord('A')) + column
        num //= 26
    return column

def column_to_num(column: str) -> int:
    num = 0
    for c in column:
        num = num * 26 + (ord(c.upper()) - ord('A') + 1)
    return num

# ? Should it be rewritten with pandas ?
def schedule_parser(path: str, workload: pandas.DataFrame):
    # Загружаем таблицу и выбираем первый лист
    wb = load_workbook(path)
    sheet = wb[wb.sheetnames[0]]
    # При помощи вложенных циклов и условий определяем столбцы групп и начинаем их парсить
    # Шагаем по ячейкам где располагаются названия групп
    for i in range(6, 10000 ,5):
        if sheet[num_to_column(i)+"2"].value == None:
            break
        elif sheet[num_to_column(i)+"2"].value != "День недели":
            group = sheet[num_to_column(i)+"2"].value
            for j in range(4,86,14):
                for k in range(0,14):
                    #print(sheet[num_to_column(i)+str(j)].value, i, j)
                    subject = sheet[num_to_column(i)+str(k+j)].value
                    if subject is None:
                        pass
                    else:
                        subject = str(subject)
                        subject = re.sub(r'\s+', ' ', subject.strip())
                        result = re.search(r"н.\s+([^\(]+)\s+\(",subject)
                        if result is not None and len(result.group(1)) > 0:
                            subject = result.group(1)
                        second_result = re.search(r"^([\d,]+ н\.)\s+(.*)$",subject)
                        if second_result is not None and len(second_result.group(2)) > 0:
                            subject = second_result.group(2)
                        professor = sheet[num_to_column(i+2)+str(k+j)].value
                        class_type = sheet[num_to_column(i+1)+str(k+j)].value
                        class_type = str(class_type)[:3]
                        if class_type == "Non":
                            pass
                        elif professor == None and class_type != None:
                            no_prof_message = subject + " " + group + " " + class_type +  " НЕ УКАЗАН ПРЕПОДАВАТЕЛЬ"
                            no_prof_message = no_prof_message.replace('\n', ' ').replace('\r', ' ')
                            #no_prof.append(no_prof_message)
                            print(no_prof_message)
                        else:
                            result = re.search(r'^([А-Яа-яЁё]+)\s([А-Яа-яЁё])[а-яё]*\s([А-Яа-яЁё])\.?$',professor)
                            if result is not None:
                                professor = result.group(1) + " " + result.group(2)+"."+result.group(3)+"."
                            subject_workload = workload[workload["Мероприятие реестра, норма времени"].str.contains(subject, na=False, case=False, regex=False) & workload["План. поток"].str.contains(group,na=False,case=False,regex=False) & workload["Вид потока"].str.contains(class_type, na=False)]
                            if subject_workload.empty == True:
                                pass
                            else:
                                expected_professor = subject_workload["ППС"].iloc[0]
                                result = re.search(r"^([А-Яа-яЁё]+)\s([А-Яа-яЁё])\w*\s([А-Яа-яЁё])\w",expected_professor)
                                if result is not None:
                                    expected_professor = result.group(1) + " " + result.group(2)+"."+result.group(3)+"."
                                    if expected_professor != professor:
                                        wrong_prof_message = subject + " " + group + " " + class_type + " ДОЛЖЕН БЫТЬ: " + expected_professor + " СТОИТ: " + professor
                                        print(wrong_prof_message)
                        
        else:
            pass

print("Добро пожаловать в программу сверки расписания")
path_schedule = sys.argv[1]
path_workload = sys.argv[2]
schedule_parser(path_schedule, workload_parser(path_workload))

#workload_parser("Учебная_нагрузка_читающего_подразделения_форма_A_02_09_2025.xls")
# startup()

#schedule_parser("ИИТ_2 курс_25-26_осень.xlsx",workload_parser("test_workload.xls"))