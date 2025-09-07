import csv
from datetime import datetime, timedelta
import random

with open('test_large.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f, delimiter=';')
    for i in range(1_000_000):
        date = datetime(2000,1,1) + timedelta(days=random.randint(0,8000))
        first = f"Name{random.randint(1,1000000)}"
        last = f"Surname{random.randint(1,1000000)}"
        sur = f"Patronymic{random.randint(1,1000000)}"
        city = f"City{random.randint(1,1000)}"
        country = f"Country{random.randint(1,100)}"
        writer.writerow([date.strftime("%Y-%m-%d"), first, last, sur, city, country])

print("CSV файл создан.")