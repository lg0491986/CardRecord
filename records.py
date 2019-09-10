# -*- coding: utf-8 -*-
import xlrd
import datetime
import json


class Records:
    def __init__(self, file):
        holidays_conf = "holidays.json"
        self.file = file
        try:
            with open(holidays_conf, "r") as f:
                self.holiday_cfg = json.load(f)
            self.records = xlrd.open_workbook(file)
        except Exception as e:
            print(type(e))
            print(e)

    def find_records_by_username(self, username):
        # first sheet
        table = self.records.sheets()[0]

        nrows = table.nrows
        ncols = table.ncols
        colnames = table.row_values(0)  # first row is title

        results = []
        for row_num in range(1, nrows):
            row = table.row_values(row_num)
            if row and row[0] == username:
                result = []

                for i in range(len(row)):
                    elem_ctype = table.cell(row_num, i).ctype
                    elem = table.cell_value(row_num, i)
                    if elem_ctype == 3:
                        d_dt = xlrd.xldate_as_datetime(elem, 0)
                        result.append(d_dt.strftime("%H:%M:%S"))
                    else:
                        if type(elem) == float:
                            elem = round(elem, 2)
                        result.append(elem)
                length = len(result)
                result.append(0)

                is_weekend = valid_weekend(result)
                if not is_weekend:
                    ret = self.valid_holidays(result)
                    # not holiday
                    if ret == 3:
                        if not valid_start_time(result):
                            result[length] = 1
                        elif not valid_end_time(result):
                            result[length] = 1
                        elif valid_over_time(result, False):
                            result[length] = 2
                        else:
                            result[length] = 0
                    elif ret == 2:
                        result[length] = 2
                        if not valid_start_time(result):
                            result[length] = 1
                        elif not valid_end_time(result):
                            result[length] = 1
                    else:
                        result[length] = ret
                else:
                    ret = self.valid_workdays(result)
                    # not workday
                    if ret == 3 and valid_over_time(result, True):
                        result[length] = 2
                    else:
                        result[length] = ret

                results.append(result)
        return results

    def valid_workdays(self, record):
        r_date = record[4]
        r_start = record[5]
        r_end = record[6]

        workdays = self.holiday_cfg["workdays"]
        for workday in workdays:
            w_date = workday['date']
            if w_date != r_date:
                continue

            if 'start' not in workday or 'end' not in workday:
                # work all days
                if valid_over_time(record, False):
                    return 2
                return 0
            w_start = workday['start']
            w_end = workday['end']

            if w_start < r_start or w_end > r_end:
                return 1
            else:
                # workday, but no overtime
                return 0
        # not match, not workday
        return 3

    def valid_holidays(self, record):
        r_date = record[4]
        r_start = record[5]
        r_end = record[6]

        holidays = self.holiday_cfg["holidays"]
        for holiday in holidays:
            h_date = holiday['date']
            if h_date != r_date:
                continue

            if 'start' not in holiday or 'end' not in holiday:
                if valid_over_time(record, True):
                    return 2
                return 0
            h_start = holiday['start']
            h_end = holiday['end']
            if h_start < r_start or h_end > r_end:
                return 1
            else:
                # holiday, but no overtime
                return 0
        # not match, not holiday
        return 3


def valid_start_time(record):
    start_time = record[5]
    stand_time = "09:00:00"
    if start_time > stand_time:
        return False
    return True


def valid_end_time(record):
    end_time = record[6]
    stand_time = "18:00:00"
    if end_time < stand_time:
        return False
    return True


def valid_over_time(record, weekend=False):
    if weekend:
        stand_delta = 4 * 60 * 60
    else:
        stand_delta = 12 * 60 * 60
    start_time = record[5]
    end_time = record[6]

    if len(start_time.split(":")) < 2 or len(end_time.split(":")) < 2:
        return False

    start = datetime.datetime.strptime(start_time, '%H:%M:%S')
    end = datetime.datetime.strptime(end_time, '%H:%M:%S')
    delta = (end - start).total_seconds()
    if delta >= stand_delta:
        return True
    return False


def valid_weekend(record):
    day = record[4]
    week = datetime.datetime.strptime(day, '%Y/%m/%d').weekday()
    if week == 5 or week == 6:
        return True
    return False


def display_by_name(file, username):
    records = Records(file)
    result = records.find_records_by_username(username)
    for row in result:
        if row:
            pass
            print(row)
    pass


if __name__ == "__main__":
    display_by_name("201908AP_card.xlsx", "周洪")
