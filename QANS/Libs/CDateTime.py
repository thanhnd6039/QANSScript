from datetime import datetime
from robot.api import logger
class CDateTime(object):
    def get_current_quarter(self):
        currentMonth = datetime.now().month
        if 1 <= currentMonth <= 3:
            return 1
        elif 4 <= currentMonth <= 6:
            return 2
        elif 7 <= currentMonth <= 9:
            return 3
        elif 10 <= currentMonth <= 12:
            return 4