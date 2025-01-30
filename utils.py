import datetime

def get_date() -> datetime.date:
        return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"

def months_by_quarter(quarter: int) -> list:
    if quarter == 1:
        return [1, 2, 3]
    elif quarter == 2:
        return [4, 5, 6]
    elif quarter == 3:
        return [7, 8, 9]
    elif quarter == 4:
        return [10, 11, 12]
    else:
        raise ValueError("Quarter must be between 1 and 4")