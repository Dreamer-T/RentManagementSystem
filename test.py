from dateutil.relativedelta import relativedelta
import datetime
start = datetime.datetime(year=2021, month=7, day=1)
print(start+relativedelta(months=+7, days=-1))
