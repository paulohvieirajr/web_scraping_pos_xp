from datetime import datetime
import time
import os

from apscheduler.schedulers.background import BackgroundScheduler
from movimento_falimentar.movimento_falimentar import MovimentoFalimentar
from email_arquivos import email

if __name__ == '__main__':
    movimento_fa_instance = MovimentoFalimentar()

    scheduler = BackgroundScheduler()
    scheduler.add_job(movimento_fa_instance.execute(), day_of_week='mon-fri', hour=8)
    scheduler.add_job(email.dispara_email(), day_of_week='mon-fri', hour=9)
    scheduler.start()

    try:
        while True:
            time.sleep(2)

    except (KeyboardInterrupt, SystemExit):
        scheduler.shutdown()
