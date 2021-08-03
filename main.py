import cx_Oracle
from datetime import datetime
from config import USERNAME, PASSWORD, TNSNAMES, SQL_STRING, HEADERS
import csv
from os import system

OUTPUT_FILE_NAME = "oFile.csv"

def main():

    print('start time: ' + str(datetime.now()))
    
    with cx_Oracle.connect(user=USERNAME, password=PASSWORD, dsn=TNSNAMES['dictionary'], encoding='UTF-8') as session: 
        cur = session.cursor()
        rows = cur.execute(SQL_STRING)
        with open(OUTPUT_FILE_NAME, 'w', newline='') as report:
            csv_writer=csv.writer(report)
            csv_writer.writerow(HEADERS)
            for row in rows:
                csv_writer.writerow(row)
            

    print('done!')

    answer=input('\nopen now? (y/n)')

    if answer=='y' or answer=='Y':
        system('start excel ' + OUTPUT_FILE_NAME)


if __name__=='__main__':
    main()
