import pandas as pd
import os
import shutil
import zipfile
import re


def format_as_date(value):
    try:
        parsed = pd.to_datetime(value)
        return parsed.strftime('%d.%m.%Y')
    except:
        return value

def format_as_time(value):
    try:
        parsed = pd.to_datetime(value)
        return parsed.strftime('%H:%M:%S')
    except:
        return value

# Читаем файлы
meeting = pd.read_excel("meeting.xlsx", dtype=str)

# присваиваем номер
meetingbase_df = pd.read_excel('meetingbase.xlsx', dtype=str)
for column in meeting.columns:
    if column not in meetingbase_df.columns:
        meetingbase_df[column] = None
row_exists = (meetingbase_df[meeting.columns] == meeting.iloc[0]).all(axis=1)
if not row_exists.any():# Проверяем, существует ли уже такая строка в meetingbase_df
    meetingbase_df['номер'] = pd.to_numeric(meetingbase_df['номер'], errors='coerce')
    next_number = meetingbase_df['номер'].max() + 1
    new_row = meeting.iloc[0].to_dict()# Добавление новой строки
    new_row['номер'] = next_number
    meetingbase_df = pd.concat([meetingbase_df, pd.DataFrame([new_row])], ignore_index=True)
    meetingbase_df.to_excel('meetingbase.xlsx', index=False)
else:
    next_number = meetingbase_df.loc[row_exists, 'номер'].values[0]
# Форматируем столбцы с датами
for column in meeting.columns:
    if "Дата" in column:
        meeting[column] = meeting[column].apply(format_as_date)
    elif "Время" in column:
        meeting[column] = meeting[column].apply(format_as_time)
print(meeting)
combined_data = pd.read_excel("combined_data_with_votes.xlsx", dtype=str)
for column in combined_data.columns:
    if "Дата" in column:
        combined_data[column] = combined_data[column].apply(format_as_date)

# Объединение данных на основе логина
df = pd.merge(meeting, combined_data, left_on="User Login", right_on="Login", how="inner")
df1 = pd.DataFrame({
    'metka': '{' + df.columns + '}',
    'chenge': df.iloc[0].values
})

agenda_rows = df1[df1['metka'].str.contains('Пункт повестки')]

# Объединяем данные "Пункт повестки" в одну строку
formatted_agenda = ' '.join([
    f"{i + 1}. {row.chenge} </w:t><w:br/><w:t>"
    for i, row in enumerate(agenda_rows.itertuples())
])

df1 = df1.fillna("")

# Удаляем старые строки "Пункт повестки" и добавляем новую объединенную строку
df1 = df1[~df1['metka'].str.contains('Пункт повестки')]
print(df1)
new_row = pd.DataFrame([{'metka': '{Пункт повестки}', 'chenge': formatted_agenda}])
df1 = pd.concat([df1, new_row], ignore_index=True)

df1.to_excel("asd.xlsx")

pathword = "Уведомление_шаблон_для_заполнения.docx"
pathzip = "B.zip"
df = df1  # для совместимости с копией замены по меткам
mem = 0
isch = 0
usch = 0
found = ""



try:
    os.remove("B.zip")
except:
    asd = 1
shutil.copy(pathword, pathzip)
fantasy_zip = zipfile.ZipFile(pathzip)  # extract zip (+need rename docx to zip +need raname vise versa
fantasy_zip.extractall("/B")
fantasy_zip.close()


with open("/B/word/document.xml", 'r', encoding='utf-8') as f:
    content = f.read()
# Применяем замену с использованием регулярного выражения
content = re.sub(r"(\{[^}{]*?)</w:t>[^}{]*?<w:t>([^}{]*?})", r"\1\2", content) #чтоб каждый раз не удалять лишнее
# Записываем обновленное содержимое обратно в файл
with open("/B/word/document.xml", 'w', encoding='utf-8') as f:
    f.write(content)


with open("/B/word/document.xml", 'r', encoding='utf-8') as f:  # save before chenge
    get_all = f.readlines()
print("xml opened")
with open("/B/word/document.xml", 'w', encoding='utf-8') as f:  # look for { and chenge it
    for i in get_all:         # STARTS THE NUMBERING FROM 1 (by default it begins with 0)
        usch = len(i)-1
        for u in i:
            try:
                if get_all[isch][usch] == "}":
                    mem = 1
            except:
                print(isch, usch, u, i, get_all)
                print(get_all[isch][usch])
            if mem == 1:
                found = get_all[isch][usch] + found
            if get_all[isch][usch] == "{":
                mem = 0
                print(found)
                #found = re.sub(r"(\{[^}{]*?)</w:t>[^}{]*?<w:t>([^}{]*?})", r"\1\2", found)
                #found = re.sub(r"</w:t>[^}{]*?(?=})", "", found)
                print(found)
                tx = df[df["metka"] == found]["chenge"].values[0]#header=False,
                try:
                    float(tx)
                    tx = str(tx).replace(".", ",")
                except:
                    tx = str(tx)
                    asd = 0
                get_all[isch] = get_all[isch][:usch] + tx + get_all[isch][usch + len(found):]
                found = ""
            usch = usch - 1
        isch = isch + 1
    f.writelines(get_all)
print("XML chanched")
try:
    os.remove("B.zip")
except:
    asd = 1
fantasy_zip = zipfile.ZipFile("B.zip", 'w')
for folder, subfolders, files in os.walk("/B"):
    for file in files:
        fantasy_zip.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder, file), "/B"))
fantasy_zip.close()  # transform it to zip
print("zip saved")
name = "Uvedomlenie_" + str(next_number) + "_" + str(df1.loc[df1['metka'] == '{User Login}', 'chenge'].iloc[0])
try:
    os.remove(name + ".docx")
    print(name, "removed")
except:
    asd = 1
os.rename("B.zip", name + ".docx")
shutil.rmtree("/B/")
print("FINISH")
