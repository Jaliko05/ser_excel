import os 
from datetime import datetime

def log(routeLog, nameAplication, message):
    route_log = routeLog + nameAplication + "_" + datetime.today().strftime('%Y-%m-%d') +".txt"
    print(  "ruta log: ", route_log)
    try:
        if not os.path.exists(route_log):
            with open(route_log, 'w') as f:
                f.write("--------------------------------------------------------------------------------\n")
                f.write("--------------"+datetime.today().strftime('%Y-%m-%d %H:%M:%S')+"----------------\n")
                f.write(message+"\n")
        else:
            with open(route_log, 'a') as f:
                f.write("--------------------------------------------------------------------------------\n")
                f.write("--------------"+datetime.today().strftime('%Y-%m-%d %H:%M:%S')+"----------------\n")
                f.write(message+"\n")
    except Exception as e:
        print(e)