from datetime import datetime
from calendar import monthrange

now = datetime.now()
cls_month = now.month-1 if now.month > 1 else 12
cls_year = now.year if cls_month != 12 else now.year-1

companyCode = "H903"
thisYear = cls_year
thisMonth = cls_month
yearANDmonth = str(cls_year)+"."+str(cls_month)
sttDate = str(cls_year)+"."+str(cls_month)+".01"
endDate = str(cls_year)+"."+str(cls_month)+"."+str(monthrange(cls_year,cls_month)[1])

"""
companyCode = "H903"
thisYear = "2021"
thisMonth = "2"
yearANDmonth = "2021.02"
sttDate = "2021.02.01"
endDate = "2021.02.28"
"""

"맞는지 프린트해봐라"

if __name__ == "__main__":

    print(thisYear)
    print(thisMonth)
    print(yearANDmonth)
    print(sttDate)
    print(endDate)
