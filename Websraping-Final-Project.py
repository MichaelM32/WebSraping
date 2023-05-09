import requests
from openpyxl import Workbook
import keys
from twilio.rest import Client


client = Client(keys.realaccount_sid, keys.auth_token)

TWnumber = "+18885744257"
myphone = "+19564564656"

url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest"
parameters = {
    'start': '1',
    'limit': '5',
    'convert': 'USD'
}
headers = {
    'Accepts': 'application/json',
    'X-CMC_PRO_API_KEY': keys.cmc_api_key,
}

response = requests.get(url, headers=headers, params=parameters)
data = response.json()['data']

workbook = Workbook()
worksheet = workbook.active
worksheet.title = "Cryptocurrencies"
worksheet["A1"] = "Name"
worksheet["B1"] = "Symbol"
worksheet["C1"] = "Price"
worksheet["D1"] = "% Change (24h)"

for i, currency in enumerate(data):
    name = currency['name']
    symbol = currency['symbol']
    price = currency['quote']['USD']['price']
    percent_change_24h = currency['quote']['USD']['percent_change_24h']
    if symbol in ["BTC", "ETH"] and abs(percent_change_24h) >= 5:
        if percent_change_24h > 0:
            message = f"{name} ({symbol}) has increased by $5"
        else:
            message = f"{name} ({symbol}) has decreased by $5"
        message += f" ({round(price - price * abs(percent_change_24h) / 100, 2):.2f} to {price:.2f})"
        client.messages.create(to=myphone, from_=TWnumber, body=message)
    worksheet[f"A{i+2}"] = name
    worksheet[f"B{i+2}"] = symbol
    worksheet[f"C{i+2}"] = price
    worksheet[f"D{i+2}"] = percent_change_24h

workbook.save("cryptocurrencies.xlsx")