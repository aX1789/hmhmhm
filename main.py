import yagmail
import pandas as pd
import datetime
import time

from config import ADMIN_EMAIL_PASSWORD, target_hour, target_minute
from news import NewsFeed

isemailsend = None
while True:
    smt = False
    if datetime.datetime.now().hour == target_hour and datetime.datetime.now().minute == target_minute:
        df = pd.read_excel("people.xlsx")

        for index, row in df.iterrows():
            # print(row, "\n")
            # print(row["interest"], "\n")

            today = datetime.date.today().isoformat()
            yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
            # print(today)
            # print(yesterday)
            # print("JA JA")
            user_interest = row["interest"]

            news_feed = NewsFeed(interests=user_interest, from_date=yesterday, to_date=today)
            email_body = news_feed.get_news()
            mail_contects = f"Hi {row['name']}\n\n See what's on about {user_interest} today!\n\n {email_body}"

            email = yagmail.SMTP(user="olegimbarig@gmail.com", password=ADMIN_EMAIL_PASSWORD)
            email.send(
                to=row["email"],
                subject=f"Your {user_interest} news are ahead for today!",
                contents=mail_contects,
            )
        smt = True
    if smt:
        time.sleep(60)