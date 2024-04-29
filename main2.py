import time
import webbrowser
from kiteconnect import KiteConnect
from kiteconnect import KiteTicker
import configparser
import requests
from datetime import datetime,date
import os
import pandas as pd

api_key =None
api_secret = None
def read_api_details():
    global api_key, api_secret
    
    config = configparser.ConfigParser()
    config.read('config.ini')
    api_key = config.get('system', 'api_key')
    api_secret = config.get('system', 'api_secret')

read_api_details()

loggedIn = False
while not loggedIn:
    try:
        kite = KiteConnect(api_key=api_key) 
        webbrowser.open(kite.login_url())
        request_token = input(f"[+] Paste your request token : ")

        data = kite.generate_session(request_token, api_secret=api_secret)
        access_token = data.get('access_token')
        kite.set_access_token(data["access_token"])

        print(f"\t[+] Welcome : {kite.profile()['user_name']}")
        loggedIn = True
    except Exception as e:
        print("Error while login : ",e)

# Login Done



# Initialise
kws = KiteTicker(api_key=api_key, access_token=access_token)
ltp_dict = {}
def on_ticks(ws, ticks):
    for tick in ticks:

        instrument_token = tick['instrument_token']
        ltp_dict[instrument_token] = tick

def on_connect(ws, response):
    
    ws.subscribe([738561, 5633])
    ws.set_mode(ws.MODE_FULL, [738561, 5633])

def on_close(ws, code, reason):
    pass

# Assign the callbacks.
kws.on_ticks = on_ticks
kws.on_connect = on_connect
kws.on_close = on_close
kws.connect(threaded=True)


def get_tick(instrument_token):
    """ returns stored ticks """
    if instrument_token is None:
        return None
    
    try:
        instrument_token = int(instrument_token)
    except:
        return None


    try:    
        return ltp_dict[instrument_token]
    except:

        kws.subscribe([instrument_token])
        kws.set_mode(kws.MODE_FULL, [instrument_token])

    return None




###------------------------------------------- EXCEL SHEET LOGIC -------------------------------------------##
import xlwings as xw    
wb = xw.Book('Dashboard.xlsx')
sheet = home_sheet = wb.sheets['home']
sheet.range('H2').value='OFF'


def is_new_day():
    today = date.today()
    instruments_file = f"instruments_{today.strftime('%Y%m%d')}.csv"
    return not os.path.isfile(instruments_file), instruments_file

# Check if it's a new day and get the file name
new_day, instruments_file = is_new_day()

if new_day:
    url = "https://api.kite.trade/instruments"
    response = requests.get(url)
    
    if response.status_code == 200:
        # Write the CSV data to a file
        with open(instruments_file, "w") as file:
            file.write(response.text)
        print("CSV file downloaded successfully.")
        df = pd.read_csv(instruments_file)  # Read CSV into DataFrame
    else:
        print("Failed to retrieve CSV data from the website")
else:
    print("CSV file already exists for today. Skipping download.")
    df = pd.read_csv(instruments_file) 

def update_row_data(row, tick):
    global sheet
    try:
        sheet.range(f"E{row}").value = tick['last_price']
        sheet.range(f"H{row}").value = tick['volume_traded']    

        sheet.range(f"C{row}").value = tick['depth']['buy'][0]['price']
        sheet.range(f"D{row}").value = tick['depth']['buy'][0]['quantity']      

        sheet.range(f"F{row}").value = tick['depth']['sell'][0]['price']
        sheet.range(f"G{row}").value = tick['depth']['sell'][0]['quantity']    
    except:
        pass

trading_symbol_dict = {}

def find_trading_symbol(instrument_token, delimiter=', '):
    try:
        return trading_symbol_dict[instrument_token]
    except Exception as e:
        new_df = df[df['instrument_token'] == instrument_token]
        contract = list(new_df.to_dict(orient='index').values())[0]

        trading_symbol_dict[instrument_token] = contract.copy()
    
    return trading_symbol_dict[instrument_token]

# def find_exchange(instrument_token, delimiter=', '):
#     new_df = df[df['instrument_token'] == instrument_token]
#     new_data = new_df.set_index('instrument_token')['exchange'].to_dict()
#     return delimiter.join(new_data.values())


def to_check_engine_status():
    start_value = sheet.range('H2').value
    if start_value == "ON":
        return True 
    return False


while True:
    try:
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        for row in range(4, last_row+1):
            instrument_token = sheet.range(f'A{row}').value
            tick = get_tick(instrument_token)
            if tick is None:
                continue
            update_row_data(row=row, tick=tick) 

            if not to_check_engine_status():
                continue
            ##----------------------------- START : BUSINESS LOGIC -----------------------------

            volume_traded = tick['volume_traded']    
            last_price = tick['last_price']
            ask_qty_available = tick['depth']['sell'][0]['quantity']
            ask_qty_price = tick['depth']['sell'][0]['price']

            contract = find_trading_symbol(int(instrument_token))
            trading_symbol = contract['tradingsymbol']
            status = sheet.range(f"N{row}").value
            sheet.range(f"B{row}").value = contract['tradingsymbol']

            entry_volume = entry_price = quantity = transaction_type = execution_status = entry_ask_qty = None
            try:
                entry_volume = sheet.range(f"I{row}").value
                entry_price = sheet.range(f"K{row}").value
                quantity = sheet.range(f"L{row}").value
                transaction_type = sheet.range(f"M{row}").value
                execution_status = sheet.range(f"N{row}").value

                entry_ask_qty = sheet.range(f"J{row}").value
            except Exception as e:
                print('Error while reading data ', e)
            
            # print(f"entry_volume = {entry_volume}, entry_price={entry_price}, quantity={quantity}, transaction_type={transaction_type}, execution_status={execution_status}, entry_ask_qty={entry_ask_qty}")
            # print(f"volume_traded = {volume_traded}")

            # Check multiple conditions here
            condition1 = entry_volume is not None and entry_price is not None and quantity is not None and transaction_type is not None and volume_traded >= entry_volume 
            condition2 = entry_volume is None and entry_price is not None and transaction_type is not None and  quantity is not None 
            condition3= entry_ask_qty is not None and transaction_type is  not None and ask_qty_available <= entry_ask_qty 
            # or execution_status in [None, 'pending'] and condition2
            if (execution_status in [None, 'pending'] and condition1)  or (execution_status in [None, 'pending'] and condition2) or (execution_status in [None, 'pending'] and condition3) :

                if entry_volume is not None and entry_price is not None and quantity is not None and transaction_type is not None and volume_traded >= entry_volume:
                #     quantity = int(ask_qty_available)
                    entry_price += 0.10

                    sheet.range(f"N{row}").value = "executed"
                    try:
                        order_id = kite.place_order(tradingsymbol=trading_symbol,
                                    exchange=contract['exchange'],
                                    transaction_type=transaction_type,
                                    quantity=int(quantity),
                                    variety=kite.VARIETY_REGULAR,
                                    order_type=kite.ORDER_TYPE_LIMIT,
                                    product="CNC",
                                    price=float(entry_price),
                                    validity=kite.VALIDITY_DAY)
                        
                        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Order placed #{order_id}, Trading_symbol => {trading_symbol}, last_price_when_order_get_executed => {last_price}, entry_price_when_order_get_executed => {entry_price}")
                    except Exception as e:
                        print('Error occurred while placing order:', e)
                elif entry_volume is None and entry_price is not None and transaction_type is not None and  quantity is not None :    
                    entry_price=entry_price 

                    sheet.range(f"N{row}").value = "executed"
                    try:
                        order_id = kite.place_order(tradingsymbol=trading_symbol,
                                    exchange=contract['exchange'],
                                    transaction_type=transaction_type,
                                    quantity=int(quantity),
                                    variety=kite.VARIETY_REGULAR,
                                    order_type=kite.ORDER_TYPE_LIMIT,
                                    product="CNC",
                                    price=float(entry_price),
                                    validity=kite.VALIDITY_DAY)
                        
                        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Order placed #{order_id}, Trading_symbol => {trading_symbol}, last_price_when_order_get_executed => {last_price}, entry_price_when_order_get_executed => {entry_price}")
                    except Exception as e:
                        print('Error occurred while placing order:', e)

                elif entry_ask_qty is not None and transaction_type is  not None and ask_qty_available <= entry_ask_qty :
                    entry_price=float(entry_price)
                    print(type(ask_qty_available))
                    quantity=ask_qty_available
                    
                    try:
                        order_id = kite.place_order(tradingsymbol=trading_symbol,
                                    exchange=contract['exchange'],
                                    transaction_type=transaction_type,
                                    quantity=int(quantity),
                                    variety=kite.VARIETY_REGULAR,
                                    order_type=kite.ORDER_TYPE_LIMIT,
                                    product="CNC",
                                    price=float(entry_price),
                                    validity=kite.VALIDITY_DAY)
                        sheet.range(f"N{row}").value = "executed"
                        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Order placed #{order_id}, Trading_symbol => {trading_symbol}, last_price_when_order_get_executed => {last_price}, entry_price_when_order_get_executed => {entry_price}")
                    except Exception as e:
                        print('Error occurred while placing order:', e)










            # if entry_volume is not None and buy_sell_status is not None:
            #     if status is None and volume_traded >= entry_volume:
            #         print(f"The volume got exceeded for {trading_symbol}")
            #         if to_check_engine_status():
            #             try:
            #                 order_id = kite.place_order(tradingsymbol=trading_symbol,
            #                             exchange=exchange,
            #                             transaction_type=buy_sell_status,
            #                             quantity=entry_qty,
            #                             variety=kite.VARIETY_REGULAR,
            #                             order_type=kite.ORDER_TYPE_LIMIT,
            #                             product="CNC",
            #                             price=entry_price,
            #                             validity=kite.VALIDITY_DAY)
            #                 sheet.range(f"N{row}").value = "Executed...!"
            #                 print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Order placed #{order_id}, Trading_symbol => {trading_symbol}, last_price_when_order_get_executed => {last_price}, entry_price_when_order_get_executed => {entry_price}")
            #             except Exception as e:
            #                 print('Error occurred while placing order:', e)
            ##----------------------------- END : BUSINESS LOGIC  -----------------------------
    except Exception as e:
        pass

    time.sleep(0.5)





# {'tradable': True, 'mode': 'full', 'instrument_token': 4454401, 'last_price': 91.2, 'last_traded_quantity': 1, 'average_traded_price': 90.78, 'volume_traded': 18650302, 'total_buy_quantity': 2489253, 'total_sell_quantity': 8290221, 'ohlc': {'open': 90.4, 'high': 91.5, 'low': 90.0, 'close': 90.25}, 'change': 1.0526315789473715, 'last_trade_time': datetime.datetime(2024, 4, 25, 11, 46, 32), 'oi': 0, 'oi_day_high': 0, 'oi_day_low': 0, 'exchange_timestamp': datetime.datetime(2024, 4, 25, 11, 46, 36), 'depth': {'buy': [{'quantity': 1380, 'price': 91.15, 'orders': 8}, {'quantity': 18551, 'price': 91.1, 'orders': 36}, {'quantity': 23540, 'price': 91.05, 'orders': 33}, {'quantity': 69081, 'price': 91.0, 'orders': 72}, {'quantity': 17108, 'price': 90.95, 'orders': 25}], 'sell': [{'quantity': 19891, 'price': 91.2, 'orders': 36}, {'quantity': 54374, 'price': 91.25, 'orders': 76}, {'quantity': 102443, 'price': 91.3, 'orders': 161}, {'quantity': 47937, 'price': 91.35, 'orders': 52}, {'quantity': 97881, 'price': 91.4, 'orders': 138}]}}




# ask_qty_available = tick['depth']['sell'][0]['price']

