from flask import Flask,url_for
from colorhash import ColorHash
import sqlite3
app = Flask(__name__)

table_file = 'time_tables.db'

years = {
        "1a" : "1STYEARA",
        "1b": "1STYEARB",
        "2a": "2NDYEARA",
        "2b": "2NDYEARB",
        "3a": "3RDYEARA",
        "3b": "3RDYEARB",
        "4": "4THYEAR",
        "mca": "MEM.TECHMSCMCA"
    }

@app.route('/')
def list_all_years():
    response = ""
    for year in years:
        response += "<a href='{}'>{}</a><br>".format(url_for('class_in_year',year=year),years[year])
    return response

@app.route('/<year>/<batch>/')
def give_table(year:str,batch:str):
    return get_html_table(year,batch)

@app.route('/<year>/')
def class_in_year(year:str):
    conn = sqlite3.connect(table_file)
    res = conn.execute("SELECT name FROM sqlite_master WHERE type='table' and name like '%{}%';".format(years[year]))
    response = ""
    for name in res:
        batch_name = name[0].split("_")[-1]
        response += "<a href='{}'>{}</a><br>".format(url_for('give_table',year=year,batch=batch_name),batch_name)
    return response

def get_html_table(year:str, batch:str):
    day = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    table_name = "t{}_{}".format(years[year],batch.upper())
    conn = sqlite3.connect(table_file)
    table_string = '''<table>
        {}
        </table>'''
    row = '''<tr>{}</tr>'''
    rows = ''
    time = 8
    data = "<td width='100'>{}</td>"
    while time < 19:
        data += "<td width='100'>{}</td>".format(time)
        time += 1
    rows += row.format(data)
    for i in range(5):
        data = '<td><b>{}</b></td>'.format(day[i])
        time = 8
        while time < 19:
            period = conn.execute('SELECT * FROM {} WHERE DAY = {} AND START_TIME = {};'.format(table_name, i, time)).fetchone()
            if period is not None:
                #TODO: fix this nonsense
                color = ColorHash(str(period[5][:-1]).strip().upper()).hex
                if period[3] == 'P':
                        time+=2
                        data = data + "<td width='100' colspan='2' bgcolor='{}'>{}</td>".format(color,period[5])
                else:
                    time+=1
                    data = data + "<td width='100' bgcolor='{}'>{}</td>".format(color,period[5])
            else:
                time+=1
                data = data + "<td width='100'>&nbsp;</td>"
        rows += row.format(data)
    return table_string.format(rows)
    

if __name__ == '__main__':
    app.run(debug = True, port = 80)