from datetime import datetime
from time import sleep
# 每n秒执行一次
def timer(n):
        print(datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))
        sleep(n)
timer(5)