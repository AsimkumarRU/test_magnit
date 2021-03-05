import sqlite3
import pyexcel
import fitz
import unicodedata
import re
from fpdf import FPDF


# Создание базы и заполнение таблиц при первом запуске
def init ():
    con = sqlite3.connect("test.db")
    cur = con.cursor()
    cur.executescript("""
        CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY,
            second_name TEXT,
            firts_name TEXT,
            patronymic TEXT,
            region_id INT,
            city_id INT,
            phone TEXT,
            email TEXT,
            FOREIGN KEY(region_id) REFERENCES regions(id),
            FOREIGN KEY(city_id) REFERENCES cities(id)
        );

        CREATE TABLE IF NOT EXISTS regions(
            id INTEGER PRIMARY KEY,
            region_name TEXT
        );

        INSERT OR IGNORE INTO regions (id, region_name)
            VALUES (0, 'Краснодарский край'); 

        INSERT OR IGNORE INTO regions (id, region_name)
            VALUES (1, 'Ростовская область');

        INSERT OR IGNORE INTO regions (id, region_name)
            VALUES (2, 'Ставропольский край');          

        CREATE TABLE IF NOT EXISTS cities(
            id INTEGER PRIMARY KEY,
            region_id INT,
            city_name TEXT,
            FOREIGN KEY(region_id) REFERENCES regions(id)
        );

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (0, 0, 'Краснодар');

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (1, 0, 'Кропоткин');

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (2, 0, 'Славянск');  

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (3, 1, 'Ростов');

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (4, 1, 'Шахты');

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (5, 1, 'Батайск');

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (6, 2, 'Ставрополь');

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (7, 2, 'Пятигорск');

        INSERT OR IGNORE INTO cities (id, region_id, city_name)
            VALUES (8, 2, 'Кисловодск');          

    """)

# Импорт данных из xls/xslx
def read_excel(name_xls):
    my_dict = pyexcel.get_array(file_name=name_xls, name_columns_by_row=0)

    for i in range(len(my_dict)):
        for j in range(len(my_dict[i])):
            if (my_dict[i][j] == 'Краснодарский край') or (my_dict[i][j] == 'Краснодар'):
                my_dict[i][j] = 0
            elif (my_dict[i][j] == 'Ростовская область') or (my_dict[i][j] == 'Кропоткин'):
                my_dict[i][j] = 1
            elif (my_dict[i][j] == 'Ставропольский край') or (my_dict[i][j] == 'Славянск'):
                my_dict[i][j] = 2
            elif my_dict[i][j] == 'Ростов':
                my_dict[i][j] = 3
            elif my_dict[i][j] == 'Шахты':
                my_dict[i][j] = 4
            elif my_dict[i][j] == 'Батайск':
                my_dict[i][j] = 5
            elif my_dict[i][j] == 'Ставрополь':
                my_dict[i][j] = 6
            elif my_dict[i][j] == 'Пятигорск':
                my_dict[i][j] = 7
            elif my_dict[i][j] == 'Кисловодск':
                my_dict[i][j] = 8
    return my_dict

# Запись в базу SQLite
def write_sql(my_dict):
    con = sqlite3.connect("test.db")
    cur = con.cursor()
    for elem in my_dict:
        tpl = tuple(elem)
        cur.execute(
            "INSERT INTO users VALUES(NULL, ?, ?, ?, ?, ?, ?, ?);", tpl)
    con.commit()
    cur.close()

# Чтение из базы SQLite
def read_sql():
    con = sqlite3.connect("test.db")
    cur = con.cursor()
    cur.execute("""
            SELECT 
                users.second_name,
                users.firts_name, 
                users.patronymic,
                regions.region_name,
                cities.city_name,
                users.phone,
                users.email
            FROM 
                users 
            JOIN 
                regions 
            ON 
                users.region_id = regions.id 
            JOIN 
                cities 
            ON 
                users.city_id = cities.id;""")
    all_results = cur.fetchall()
    cur.close()
    my_list = [list(ele) for ele in all_results]
    return my_list

# Экспорт данных в xls/xslx
def write_excel(my_list, name_xls):
    pyexcel.save_as(array=my_list, dest_file_name=name_xls)

# Импорт данных из pdf
def import_pdf(pdf_document):  
    doc = fitz.open(pdf_document)
    page1 = doc.loadPage(0)  
    page1text = page1.getText("text")  
    text = unicodedata.normalize("NFKD", page1text)
    text = list(text.split("\n"))

    my_list = text[0].split(' ')
    if (text[4].split(' '))[1] == 'Краснодар':
        my_list.append('0')
        my_list.append('0')
    elif (text[4].split(' '))[1] == 'Кропоткин':
        my_list.append('0')
        my_list.append('1')
    elif (text[4].split(' '))[1] == 'Славянск':
        my_list.append('0')
        my_list.append('2')
    elif (text[4].split(' '))[1] == 'Ростов':
        my_list.append('1')
        my_list.append('3')
    elif (text[4].split(' '))[1] == 'Шахты':
        my_list.append('1')
        my_list.append('4')
    elif (text[4].split(' '))[1] == 'Батайск':
        my_list.append('1')
        my_list.append('5')
    elif (text[4].split(' '))[1] == 'Ставрополь':
        my_list.append('2')
        my_list.append('6')
    elif (text[4].split(' '))[1] == 'Пятигорск':
        my_list.append('2')
        my_list.append('7')
    elif (text[4].split(' '))[1] == 'Пятигорск':
        my_list.append('2')
        my_list.append('8')

    pattern = re.findall(r'(\+7|8).*?(\d{3}).*?(\d{3}).*?(\d{2}).*?(\d{2})', text[2])
    my_list.append("". join(pattern[0]))
    my_list.append(text[3])

    con = sqlite3.connect("test.db")
    cur = con.cursor()
    cur.execute(
        "INSERT INTO users VALUES(NULL, ?, ?, ?, ?, ?, ?, ?);", my_list)
    con.commit()
    cur.close()    

# Экспорт данных в pdf
def export_pdf(my_list, path_export):
    pdf = FPDF()
    pdf.add_font('ArialUnicode',fname='Arial Unicode MS.ttf',uni=True)

    for i in range(len(my_list)):
        pdf.add_page()
        pdf.set_font("ArialUnicode", size=16)
        pdf.cell(20, 10, txt=f"{my_list[i][0]} {my_list[i][1]} {my_list[i][2]}", ln=1)
        pdf.set_font("ArialUnicode", size=12)
        pdf.cell(20, 7, txt=f"Регион: {my_list[i][3]}", ln=1)
        pdf.cell(20, 7, txt=f"Город: {my_list[i][4]}", ln=1)
        pdf.cell(20, 7, txt=f"Номер телефона: {my_list[i][5]}", ln=1)
        pdf.cell(20, 7, txt=f"Email: {my_list[i][6]}", ln=1)
    pdf.output(path_export)

if __name__ == "__main__":
    init ()
    while True:
        try:
            func = int(input (
                f"Выберите операцию: "
                f"1 - Импорт данных из xls/xslx"
                f"2 - Экспорт данных в xls/xslx"
                f"3 - Импорт данных из pdf"
                f"4 - Экспорт данных в pdf "
                ))
        except ValueError:
            print ("Это не число, попробуйте снова.")

        else:       
            if func == 1:
                name = input (
                    f"Введите полный путь до файла в формате: "
                    f"/ваша/папка/имя_файла.xlsx "
                    )
                my_dict = read_excel(name)
                write_sql(my_dict)
                break
            elif func == 2:
                name = input (
                    f"Введите полный путь до файла и его имя в формате: "
                    f"/ваша/папка/имя_файла.xlsx "
                    )
                my_list = read_sql()
                write_excel(my_list, name)
                break
            elif func == 3:
                name = input (
                    f"Введите полный путь до файла summary.pdf в формате: "
                    f"/ваша/папка/summary.pdf "
                    )
                import_pdf(name)
                break
            elif func == 4:
                name = input (
                    f"Введите полный путь до файла в формате: "
                    f"/ваша/папка/имя_файла.pdf "
                    )
                my_list = read_sql()
                export_pdf(my_list, name)
                break

