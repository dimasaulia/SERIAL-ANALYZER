import serial
import serial.tools.list_ports
from datetime import datetime
import xlsxwriter
import json

ports = serial.tools.list_ports.comports()
portsList = []
for port in ports:
    portsList.append(str(port))
    print(str(port))

selectedPort = input("Plase select available port. COM:")
watchSring = input("String to analyze (starts with):")
fileName = input("Output File:")

workbook = xlsxwriter.Workbook(f'Serial Recorder {fileName}.xlsx')
date_format = workbook.add_format({'num_format': 'hh:mm:ss.000', 'align': 'left'})
worksheet = workbook.add_worksheet("ESP MESH")
worksheet.write("A1", "NO")
worksheet.write("B1", "PAYLOAD")
worksheet.write("C1", "TIME STAMP")
index = 2
serialDebug = serial.Serial(port=f'COM{selectedPort}', baudrate=115200,
                            bytesize=8, parity="N", stopbits=serial.STOPBITS_TWO, timeout=1)
while True:
    try:
        serialString = serialDebug.readline().decode("utf").rstrip("\n")
        if (len(serialString) > 0):
            print(serialString)
            if(serialString.startswith(watchSring)):
                worksheet.write(f"A{int(index)}", f"{index-1}")
                worksheet.write(f"B{int(index)}", f'{serialString.replace("_x000D_"," ")}')
                worksheet.write(f"C{int(index)}", f"{datetime.strftime(datetime.now(),'%H:%M:%S.%f')}", date_format)
                index += 1
    except KeyboardInterrupt:
        print("Handling interrupt...")
        print("Saving File...")
        workbook.close()
        break

print("Process Done")
