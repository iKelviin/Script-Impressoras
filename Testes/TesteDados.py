import time
from datetime import datetime
import datetime as dt


Hoje = dt.datetime.now()
Hoje1 = Hoje.strftime("%d/%m/%Y %H:%M")
Ontem = Hoje - dt.timedelta(1)
OntemdeOntem = Ontem - dt.timedelta(1)
atual = dt.datetime.today()

print(Hoje1)
