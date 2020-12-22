"""
This file contains the main function for the timesheet scrapper.  The scrapper
is for taking collected timesheets, scrapping them, and placing that data
into the production database
"""

# imported from standard library
import os
import logging
import pandas as pd
import sqlalchemy as sa
import xlrd

# imported from third party repos

# imported from local directories
import config as cfg
import databaseFunctions as dbfnc
import myFunctions as myfnc
from timesheet import TS2015, TS2011, TSCasual

# TODO: make a generator loop

# logging info
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO) # change to DEBUG when required

formatter = logging.Formatter('%(asctime)s:%(name)s:%(levelname)s:%(message)s')

file_handler = logging.FileHandler(cfg.log_files + '\\' + 'scrape.log')
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)

def grab_01_head_alpha(sheet, dummy1, name_row, name_col):
    data = sheet.cell_value(name_row, name_col)
    head_alpha = dbfnc.find_head_alpha_number2(data)
    return head_alpha

def grab_02_head_id(sheet, dummy1, name_row, name_col):
    data = sheet.cell_value(name_row, name_col)
    head_id = dbfnc.find_head_number2(data)
    return head_id

def grab_03_date(sheet, row, col):
    data = ts15.ts_grab_date(sheet, row, col)
    return data

#no 4 or 5 yet

def grab_06_evnt_yr(sheet, row, col):
    my_date = ts15.ts_grab_date(sheet, row, col)
    evnt_yr = dbfnc.grabeventYR2(my_date)
    return evnt_yr

def grab_07_evnt_id(sheet, row, accnt_col, date_col):
    head_acct = ts15.ts_15_write_acct(sheet, row, accnt_col)
    my_date = ts15.ts_grab_date(sheet, row, date_col)
    evnt_id = ts15.ts_event_id(head_acct, my_date)
    return evnt_id

# no 8, 9, 10 yet

def grab_11_acct(sheet, row, col):
    data = ts15.ts_15_write_acct(sheet, row, col)
    return data

def grab_12_blacks(sheet, row, col):
    data = ts15.ts_blacks_call(sheet, row, col)
    return data

def grab_13_MP(sheet, row, col):
    data = ts15.ts_mp(sheet, row, col)
    return data

def create_null(dummy, dummy1):
    data = ""
    return data

def create_eight(dummy, dummy1):
    data = 8
    return data

def create_zero(dummy, dummy1):
    data = 0
    return data


# and this is the app...
def main():
    # set up the db column headers in lists to receive the scraped data
    # The order of these comes form teh db and is important, don't change it.
    crew_keys = ["Shift", "CrewIDLetter", "CrewIDNumber",
                 "Date", "InTime", "OutTime",
                 "EventYrID", "EventID", "Reg",
                 "OT", "Double", "Acct",
                 "Blackscall", "MP", "ShiftType"]
    head_keys = ["Shift", "HeadIDLetter", "HeadIDNumber",
                 "Date", "InTime", "OutTime",
                 "EventYrID", "EventID", "Reg",
                 "OT", "Double", "Acct",
                 "Blackscall", "MP"]

    # # Creating empty dataframes with column names only
    df_head = pd.DataFrame(columns=head_keys)
    df_crew = pd.DataFrame(columns=crew_keys)

    # find the max number in the Shift number lists so as to correctly
    # number your ShiftIDs
    # this info is pulled directly from the db
    query = "SELECT * FROM HeadShiftWorkedTable"
    with dbfnc.connection(cfg.my_driver, cfg.my_server, cfg.my_db) as conn:
        df_h_shift = pd.read_sql(query, conn)
    last_head_record = df_h_shift['HeadShiftWorkedID'].max()
    next_head_num = last_head_record + 1

    query = "SELECT * FROM CrewShiftWorkedTable"
    with dbfnc.connection(cfg.my_driver, cfg.my_server, cfg.my_db) as conn:
        df_c_shift = pd.read_sql(query, conn)
    last_crew_record = df_c_shift['ShiftWorkedID'].max()
    next_crew_num = last_crew_record + 1

    # create a list of the read_books this is a list of the files which
    # we will need to read so we need a list to iterate over
    read_list = os.listdir(cfg.my_dir)
    print(f"We need to read approximately {len(read_list)} files")
    print("Are you ready to read? RETURN for yes, CTRL+C for no.")
    logger.info('Waiting for input from user')
    input()
    logger.info('Input from user received')

    # here is the loop to iterate through the read_list
    for i in range(len(read_list)):
        print()
        print()
        print("-----------------------------------------------------------------------------------")
        read_file = (cfg.my_dir + '\\' + read_list[i])
        print(read_file)
        read_book = xlrd.open_workbook(read_file) # TODO: write a try bock here
        read_sheet = read_book.sheet_by_index(0)

        # is this actually a timesheet? And which one is it?
        # the timesheets class arguments are as follows
        # (version, name_row, name_column, start_data_row,
        # start_data_col, spaces_per_day)
        if read_sheet.cell_value(7, 0) == 'SUNDAY':
            print("This timesheet was designed in 2011. Begin data scrape")
            ts11 = TS2011('ts11', 3, 1, 7, 1, 55, 7)
            # r_row = ts11.start_data_row

            print("This 2011 loop was deleted during the python2 to python3 upgrade")
            print("You'll probably want to re-write this 2011 loop someday")
            # TODO :write the 2011 loop

        elif read_sheet.cell_value(14, 0) == 'SUNDAY':
            print("This timesheet belongs to a casual. Begin data scrape")
            ts_cas = TSCasual('ts_cas', 4, 1, 14, 1, 55, 6)
            ts_cas.ts_print_my_name(read_sheet)

            # A loop to iterate through the time slots one at a time
            for r_row in range(ts_cas.start_data_row, ts_cas.end_data_row):
                r_col = ts_cas.start_data_col + 1
                if read_sheet.cell_type(r_row, r_col) != 0:

                    # write shift number
                    crew_data_list = [next_crew_num]
                    next_crew_num += 1

                    # Grab a casual Alpha number from the db
                    data = read_sheet.cell_value(ts_cas.name_row, ts_cas.name_column)
                    myfnc.from_func_2_db(crew_data_list, dbfnc.find_crew_Alpha_number2, data)

                    # Grab a casual id number from the db
                    data = read_sheet.cell_value(ts_cas.name_row, ts_cas.name_column)
                    cas_id = dbfnc.find_crew_number2(data)
                    crew_data_list.append(cas_id)

                    # Grab ts date
                    myfnc.from_func_2_db(crew_data_list, ts_cas.ts_grab_date, read_sheet, r_row, 3)

                    # Grab in time
                    myfnc.from_func_2_db(crew_data_list, ts_cas.ts_write_time, read_sheet, r_row, 0)

                    # Grab out time
                    myfnc.from_func_2_db(crew_data_list, ts_cas.ts_write_time, read_sheet, r_row, 1)

                    # Grab event year
                    data = ts_cas.ts_grab_date(read_sheet, r_row, 3)
                    evnt_yr = dbfnc.grabeventYR2(data)
                    crew_data_list.append(evnt_yr)

                    # Grab Event ID
                    my_date = ts_cas.ts_grab_date(read_sheet, r_row, 3)
                    crew_accnt_code = ts_cas.ts_cas_write_acct()
                    show = ts_cas.ts_event_id(crew_accnt_code, my_date)
                    crew_data_list.append(show)

                    # grab reg time, ot, dt
                    data = ts_cas.ts_grab_hrs(read_sheet, r_row, 2)
                    crew_data_list.append(data)

                    data = ts_cas.ts_grab_hrs(read_sheet, r_row, 3)
                    crew_data_list.append(data)

                    data = ts_cas.ts_grab_hrs(read_sheet, r_row, 4)
                    crew_data_list.append(data)

                    # write accounting code
                    crew_data_list.append(crew_accnt_code)

                    # blackscall true/false
                    myfnc.from_func_2_db(crew_data_list,ts_cas.ts_blacks_call, read_sheet, r_row, 5)

                    # Grab MP
                    myfnc.from_func_2_db(crew_data_list, ts_cas.ts_mp, read_sheet, r_row, 6)

                    # Grab Shiftype
                    data = 9 # default to 9, general hand
                    crew_data_list.append(data)

                    print(crew_data_list)

                    # add this row to the df
                    print("adding to crew df")
                    my_dict = dict(zip(crew_keys, crew_data_list))
                    df_crew = df_crew.append(my_dict, ignore_index=True)

                else:
                    print("no data in cel B" + str((r_row) + 1))  # move on to the next time slot
                    # TODO: what is this pycharm error above?

        # now we move onto the salaried head loops
        elif read_sheet.cell_value(19, 0) == 'SUNDAY':
            print("This timesheet was designed in 2015. Begin data scrape")
            ts15 = TS2015('ts15', 15, 2, 19, 2, 69, 7)
            ts15.ts_print_my_name(read_sheet)

            # first, lets loop over the 8 hr flags and write those to the list
            r_row = ts15.start_data_row + ts15.spaces_per_day - 1
            while r_row < ts15.end_data_row:
                if read_sheet.cell_value(r_row, 0) == "":
                    r_row+=ts15.spaces_per_day
                else:
                    # grab the data you want for more than one action
                    my_date = ts15.ts_grab_date(read_sheet, r_row, 1)
                    head_acct = ts15.ts_15_write_acct(read_sheet, r_row, 6)

                    # write shift number
                    head_data_list = [next_head_num]
                    next_head_num += 1

                    # Grab a salaried head Alpha number from the db
                    data = read_sheet.cell_value(ts15.name_row, ts15.name_column)
                    myfnc.from_func_2_db(head_data_list, dbfnc.find_head_alpha_number2, data)

                    # Grab a head id number from the db
                    data = read_sheet.cell_value(ts15.name_row, ts15.name_column)
                    head_id = dbfnc.find_head_number2(data)
                    head_data_list.append(head_id)

                    # write ts date
                    head_data_list.append(my_date)

                    # in time
                    data = ""
                    head_data_list.append(data)

                    # out time
                    data = ""
                    head_data_list.append(data)

                    # Grab event year
                    evnt_yr = dbfnc.grabeventYR2(my_date)
                    head_data_list.append(evnt_yr)

                    # Grab Event ID
                    data = ts15.ts_event_id(head_acct, my_date)
                    head_data_list.append(data)

                    # reg time, ot, dt
                    data = 8
                    head_data_list.append(data)

                    data = 0
                    head_data_list.append(data)

                    data = 0
                    head_data_list.append(data)

                    # write accounting code
                    in_out(head_acct)
                    head_data_list.append(head_acct)

                    # showcall true/false
                    myfnc.from_func_2_db(head_data_list, ts15.ts_blacks_call, read_sheet, r_row, 7)

                    # Grab MP
                    myfnc.from_func_2_db(head_data_list, ts15.ts_mp, read_sheet, r_row, 5)

                    print(head_data_list)

                    # add this row to the df
                    print("adding to head df")
                    my_dict = dict(zip(head_keys, head_data_list))
                    df_head = df_head.append(my_dict, ignore_index=True)

                    r_row += ts15.spaces_per_day

            # second, lets loop over the stat flags and write those to the list
            r_row = ts15.start_data_row + ts15.spaces_per_day - 4
            while r_row < ts15.end_data_row:
                if read_sheet.cell_value(r_row, 0) == "":
                    r_row+=ts15.spaces_per_day
                else:
                    # grab the data you want for more than one action
                    my_date = ts15.ts_grab_date(read_sheet, r_row, 1)
                    head_acct = ts15.ts_15_write_acct(read_sheet, r_row, 6)

                    # write shift number
                    head_data_list = [next_head_num]
                    next_head_num += 1

                    # Grab a salaried head Alpha number from the db
                    data = read_sheet.cell_value(ts15.name_row, ts15.name_column)
                    myfnc.from_func_2_db(head_data_list, dbfnc.find_head_alpha_number2, data)

                    # Grab a head id number from the db
                    data = read_sheet.cell_value(ts15.name_row, ts15.name_column)
                    head_id = dbfnc.find_head_number2(data)
                    head_data_list.append(head_id)

                    # write ts date
                    head_data_list.append(my_date)

                    # in time
                    data = ""
                    head_data_list.append(data)

                    # out time
                    data = ""
                    head_data_list.append(data)

                    # Grab event year
                    evnt_yr = dbfnc.grabeventYR2(my_date)
                    head_data_list.append(evnt_yr)

                    # Grab Event ID
                    data = ts15.ts_event_id(head_acct, my_date)
                    head_data_list.append(data)

                    # reg time, ot, dt
                    data = 8
                    head_data_list.append(data)

                    data = 0
                    head_data_list.append(data)

                    data = 0
                    head_data_list.append(data)

                    # write accounting code
                    head_data_list.append(head_acct)

                    # showcall true/false
                    myfnc.from_func_2_db(head_data_list, ts15.ts_blacks_call, read_sheet, r_row, 7)

                    # Grab MP
                    myfnc.from_func_2_db(head_data_list, ts15.ts_mp, read_sheet, r_row, 5)

                    print(head_data_list)

                    # add this row to the df
                    print("adding to head df")
                    my_dict = dict(zip(head_keys, head_data_list))
                    df_head = df_head.append(my_dict, ignore_index=True)

                    r_row += ts15.spaces_per_day

            # A loop to iterate through the time slots one at a time
            for r_row in range(ts15.start_data_row, ts15.end_data_row):
                # Find the first slot with data
                r_col = ts15.start_data_col

                if read_sheet.cell_type(r_row, r_col) != 0:
                    # grab the data you want for more than one action
                    my_date = ts15.ts_grab_date(read_sheet, r_row, 1)
                    head_acct = ts15.ts_15_write_acct(read_sheet, r_row, 6)

                    # write shift number
                    head_data_list = [next_head_num]
                    next_head_num += 1

                    # Grab a salaried head Alpha number from the db
                    data = read_sheet.cell_value(ts15.name_row, ts15.name_column)
                    myfnc.from_func_2_db(head_data_list, dbfnc.find_head_alpha_number2, data)

                    # Grab a head id number from the db
                    data = read_sheet.cell_value(ts15.name_row, ts15.name_column)
                    head_id = dbfnc.find_head_number2(data)
                    head_data_list.append(head_id)

                    # write ts date
                    head_data_list.append(my_date)

                    # Grab in time w the kris fix
                    if head_id == 3:
                        myfnc.from_func_2_db(head_data_list, ts15.ts_15_kf_format, read_sheet, r_row, 0)
                    else:
                        myfnc.from_func_2_db(head_data_list, ts15.ts_write_time, read_sheet, r_row,0)

                    # Grab out time w the kris fix
                    if head_id == 3:
                        myfnc.from_func_2_db(head_data_list, ts15.ts_15_kf_format, read_sheet, r_row, 1)
                    else:
                        myfnc.from_func_2_db(head_data_list, ts15.ts_write_time, read_sheet, r_row, 1)

                    # Grab event year
                    evnt_yr = dbfnc.grabeventYR2(my_date)
                    head_data_list.append(evnt_yr)

                    # Grab Event ID
                    data = ts15.ts_event_id(head_acct, my_date)
                    head_data_list.append(data)

                    # grab reg time, ot, dt
                    data = ts15.ts_grab_hrs(read_sheet, r_row, 2)
                    head_data_list.append(data)

                    data = ts15.ts_grab_hrs(read_sheet, r_row, 3)
                    head_data_list.append(data)

                    data = ts15.ts_grab_hrs(read_sheet, r_row, 4)
                    head_data_list.append(data)

                    # write accounting code
                    head_data_list.append(head_acct)

                    # showcall true/false
                    myfnc.from_func_2_db(head_data_list, ts15.ts_blacks_call, read_sheet, r_row, 7)

                    # Grab MP
                    myfnc.from_func_2_db(head_data_list, ts15.ts_mp, read_sheet, r_row, 5)

                    print(head_data_list)

                    # add this row to the df
                    print("adding to head df")
                    my_dict = dict(zip(head_keys, head_data_list))
                    df_head = df_head.append(my_dict, ignore_index=True)

                else:
                    print("no data in cel E" + str((r_row) + 1))  # move on to the next time slot
        print("-----------------------------------------------------------------------------------")
    else:
        print()
        print("We are done")
    print()

    # send the df to the db
    tbl = 'TMPtblWeeklyHeadsData'
    with dbfnc.connection(cfg.my_driver, cfg.my_server, cfg.my_db) as conn:
        tbl_exist = dbfnc.checkTableExists(conn, tbl)

    if tbl_exist:
        with dbfnc.connection(cfg.my_driver, cfg.my_server, cfg.my_db) as conn:
            print("let's empty Heads the table")
            cur = conn.cursor()
            cur.execute("DROP Table TMPtblWeeklyHeadsData")

    else:
        print("OOPS! IF YOU ARE READING THIS WE HAVE A PROBLEM!")
        print("NEXT COMES THE ERROR MESSAGE!")

    print("Table path cleared, let's write to the db")
    engine = sa.create_engine(cfg.alc_str)
    df_head.to_sql(tbl,
                   con=engine,
                   if_exists='append',
                   index=False,
                   dtype={'Shift': sa.types.INT,
                          'HeadIDLetter': sa.types.NVARCHAR(length=255),
                          'HeadIDNumber': sa.types.INT,  # notice this is an INT
                          'Date': sa.dialects.mssql.DATETIME2(0),
                          'InTime': sa.dialects.mssql.DATETIME2(0),
                          'OutTime': sa.dialects.mssql.DATETIME2(0),
                          'EventYrID': sa.types.NVARCHAR(length=255),
                          'EventID': sa.types.NVARCHAR(length=255),
                          'Reg': sa.types.FLOAT,
                          'OT': sa.types.FLOAT,
                          'Double': sa.types.FLOAT,
                          'Acct': sa.types.NVARCHAR(length=255),
                          'Blackscall': sa.dialects.mssql.BIT,
                          'MP': sa.dialects.mssql.BIT}
                   )
    print()

    # send the df to the db
    tbl = 'TMPtblWeeklyCrewData'
    with dbfnc.connection(cfg.my_driver, cfg.my_server, cfg.my_db) as conn:
        tbl_exist = dbfnc.checkTableExists(conn, tbl)

    if tbl_exist:
        with dbfnc.connection(cfg.my_driver, cfg.my_server, cfg.my_db) as conn:
            print("let's empty the Crew table")
            cur = conn.cursor()
            cur.execute("DROP Table TMPtblWeeklyCrewData")

    else:
        print("OOPS! IF YOU ARE READING THIS WE HAVE A PROBLEM!")
        print("NEXT COMES THE ERROR MESSAGE!")

    print("Table path cleared, let's write to the db")
    engine = sa.create_engine(cfg.alc_str)
    df_crew.to_sql(tbl,
                   con=engine,
                   if_exists='append',
                   index=False,
                   dtype={'Shift': sa.types.INT,
                          'CrewIDLetter': sa.types.NVARCHAR(length=255),
                          'CrewIDNumber': sa.types.NVARCHAR(length=255),  # notice this is a CHAR
                          'Date': sa.dialects.mssql.DATETIME2(0),
                          'InTime': sa.dialects.mssql.DATETIME2(0),
                          'OutTime': sa.dialects.mssql.DATETIME2(0),
                          'EventYrID': sa.types.NVARCHAR(length=255),
                          'EventID': sa.types.NVARCHAR(length=255),
                          'Reg': sa.types.FLOAT,
                          'OT': sa.types.FLOAT,
                          'Double': sa.types.FLOAT,
                          'Acct': sa.types.NVARCHAR(length=255),
                          'Blackscall': sa.dialects.mssql.BIT,
                          'MP': sa.dialects.mssql.BIT,
                          'ShiftType': sa.types.INT}
                   )
    print()
    print("All done! Time to bailout")
    print()



if __name__ == '__main__':
    logger.info('~~~The fiile scrape.py has started~~~')
    print()
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("         timesheetscrapper_python3 package launched")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print()
    main()
    print()
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("                        VICTORY!")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print()
    logger.info('~~~~The fiile scrape.py has finished OK~~~')
