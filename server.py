import xlwings as xw
import requests

wb = xw.Book('App.xlsx')
weather = wb.sheets['Weather']
get_token = wb.sheets['Additional Details']
token = get_token.range('B2').value
print("Server running for file App.xlsx")

def give_data(city,unit):
    data = requests.get('http://api.openweathermap.org/data/2.5/weather?q='+city+'&appid='+token).json()
    temp = round(float(data['main']['temp']) - 273.15,2) 
    if unit == 'F':
        temp = temp * ( 9 / 5 ) + 32
    humid = data['main']['humidity']
    return [temp, humid]

while 1:
    for i in range(2, weather.range('A2').end('down').row+1):
        if weather.range('A'+str(i)+':E'+str(i)).value[4] == 1.0:
            city = weather.range('A'+str(i)+':E'+str(i)).value[0]
            unit = weather.range('A'+str(i)+':E'+str(i)).value[3]
            data = give_data(city,unit)
            weather.range('B'+str(i)).value = data[0]
            weather.range('C'+str(i)).value = data[1]