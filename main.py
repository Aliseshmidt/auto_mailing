import time
import schedule
from ots import get_xls_file_from_ots
from mail import send_email_with_attachment
def job():
    get_xls_file_from_ots()
    send_email_with_attachment()

job()

# Планирование задач
# schedule.every().monday.at("15:20").do(job)
# schedule.every().monday.at("15:11").do(job)
# schedule.every().thursday.at("10:00").do(job)
#
# while True:
#     schedule.run_pending()
#     time.sleep(1)


