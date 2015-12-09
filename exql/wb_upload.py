"""read in a workbook and upload it to SQL"""
from openpyxl import Workbook
from sqlalchemy import Column, Integer, Unicode, String, Float, MetaData, Table
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import mapper, create_session
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.cell import column_index_from_string
import click
import os
import re
from Sheet import Sheet
from Upload import Upload
import utils
import logging


logging.basicConfig(level=logging.DEBUG)

def get_conn():
    return create_engine(os.environ['dbk'])
    #m = MetaData()
    #m.reflect(engine)


def append_sht(sht, tbl_nm, col_nms):
    """append ws values to a table"""
    cur = conn.cursor()

    fmt = ','.join(['%s'] * len(sht.values))
    insert_query = 'insert into {0} {1} values {2}'.format(tbl_nm, sht.col_nms, fmt)
    cur.execute(insert_query, sht.values)
    cur.commit()

def create_with_sheet(sht, tbl_nm):
    """insert ws into a new table"""
    cur = conn.cursor()


def get_sht(wb, sht_nm):
    """decide which sheet to pull from wb (default to first if None)"""
    if sht_nm == None:
        return wb.sheetnames[0]
    else:
        return sht_nm


#Do we want to have option for multiple wbs? how to handle that logic? this would be
#better in utils file i think
def import_wbs(src, path, sheet_name):
    """decide if we're returning either a single or list of workbook objects"""

    wbs = []
    if os.path.isfile(path):
        try:
            wbs = etl.pull_wb(path, src)
        except:
            Exception('Cant pull this workbook!')

    else:
        #read in WS by WS and note any invalid notebooks. if we find >=1, Exception
        file_list = etl.get_file_list(path, src)
        bad_wb = []

        for f in file_list:
            try:
                wbs.append[etl.pull_wb(path, src)]
            except:
                bad_wb.append(path)

        if len(bad_wb) > 0:
            Exception('Workbook directory has an invalid workbook!' + path)

    return get_sheets(wbs, sheet_name)

def get_sheets(wbs, sheet_name):
    """return specified worksheets from given wbs"""
    ws = []

    for wb in wbs:
        for s in wb.worksheets:
            if s.title in sheet_name:
                ws.append(s)

    return ws

def create_with_sheet(sht, engine):
    """create a SQL table and populate with data"""
    session = create_session(bind=engine, autocommit=False, autoflush=True)
    metadata = MetaData(bind=engine)
    tbl_met = Table(sht.tbl_nm, metadata, Column('pk', Integer, primary_key=True),
                    *(Column(col, String()) for col in sht.col_nms))

    metadata.create_all()
    logging.info("Table %s has been created" % sht.tbl_nm)
    mapper(Upload, tbl_met)

    for r_i, row in enumerate(sht.values):
        up = Upload()
        assert(len(sht.col_nms) == len(row))
        for c_i, v in enumerate(row):
            setattr(up, sht.col_nms[c_i], v)

        if r_i%100 == 0 :
            logging.info("%i entries have been parsed" % r_i)
            session.commit()
        session.add(up)

    session.commit()
    logging.info("Values have been uploaded to table %s" % sht.tbl_nm)

@click.command()
@click.option('--src', help='local files or on dropbox?', type = click.Choice(['db','local']))
@click.option('--sheet_name', help='which sheet should we pull? Defaults to first if blank')
@click.option('--path', help='file path to xlsx or directory')
@click.option('--test', help='are we testing?', is_flag = True)
@click.option('--append', help='we are appending to SQL table or insterting?', is_flag = True, default=False)
@click.option('--table_name', help='name of table appending or creating to')
@click.option('--col_names', help='optional, list of column names in SQL table. defaults to header', required=False)
def ingest(src, path, test, append, table_name, col_names, sheet_name):
    """iterate through wbs and send to sql"""
    logging.info("Extracting Worksheet. Stand by.")
    wb = utils.pull_wb(path, src, True)
    engine = get_conn()
    sht_raw = wb.get_sheet_by_name(get_sht(wb, sheet_name))
    sht = Sheet(ws = sht_raw, name = table_name, col_nms = col_names)

    if not engine.has_table(table_name) and append:
        raise Exception('Table does not exist to append to')

    elif append:
        append_sht(sht, engine)

    elif not append:
        create_with_sheet(sht, engine)

if __name__ == '__main__':
    ingest()

