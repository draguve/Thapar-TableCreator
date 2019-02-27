from flask import Flask
import sqlite3
app = Flask(__name__)

@app.route('/<file_name>/<table_name>')
def give_table(file_name:str,table_name:str):
    return get_html_table(file_name,table_name)

def get_html_table(year:str, batch:str):
    day = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
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
    table_name = "t{}_{}".format(years[year],batch.upper())
    conn = sqlite3.connect('time_tables.db')
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
            period = conn.execute('SELECT * FROM {} WHERE DAY = {} AND START_TIME = {}'.format(table_name, i, time)).fetchone()
            if period is not None:
                if period[3] == 'P':
                        time+=2
                        data = data + "<td width='100' colspan='2'>{}</td>".format(period[5])
                else:
                    time+=1
                    data = data + "<td width='100'>{}</td>".format(period[5])
            else:
                time+=1
                data = data + "<td width='100'>&nbsp;</td>"
        rows += row.format(data)
    return table_string.format(rows)
    

if __name__ == '__main__':
    app.run(debug = True, port = 80)