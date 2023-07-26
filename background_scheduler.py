import atexit
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger
from commissionSales import generate_report  # This should be the path to your generate_report function

scheduler = BackgroundScheduler()
scheduler.start()
scheduler.add_job(func=generate_report, trigger=IntervalTrigger(weeks=1))

# Don't forget to gracefully shut down the scheduler when you're done with it.
atexit.register(lambda: scheduler.shutdown())

# Keep the script running.
while True:
    pass
