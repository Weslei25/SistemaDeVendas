import datetime as dt
import pytz
import pendulum as pl
from datetime import date

dthoje = dt.datetime.now()
print(dthoje.strftime('%Y-%m-%d'))
# print(dthoje.strftime('%d/%m/%Y %H:%M:%S %Z'))



cr_date = '31/10/2013 18:23:29.000227'
cr_date = dt.datetime.strptime(cr_date, '%d/%m/%Y %H:%M:%S.%f')
cr_date = cr_date.strftime('%Y-%m-%d %H:%M:%S.%f')

print(cr_date)






'''formatted_date = dt.datetime.strptime(transaction_date[:-6], '%d/%m/%Y:%H:%M:%S')
dthoje = formatted_date.replace(tzinfo=pytz.UTC)
my_timezone = pytz.timezone('America/Sao_Paulo')
local_date = dthoje.astimezone(my_timezone)
print(local_date.strftime('%d/%m/%Y %H:%M:%S %Z'))'''