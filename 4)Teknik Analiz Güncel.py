import pandas as pd
import yfinance as yf
import ta
import numpy as np
import os
import requests
from io import StringIO
import warnings
from tqdm import tqdm 

# TimedeltaIndex uyarılarını devre dışı bırak
warnings.filterwarnings('ignore', message="The 'unit' keyword in TimedeltaIndex construction is deprecated", category=FutureWarning)

# Excel dosyasını oku
excel_dosya_adi = "C:/Users/q/Desktop/bt/finansal_veriler_hesaplanmis.xlsx"
veriler = pd.read_excel(excel_dosya_adi)

# Dosya yolunu belirtin
dosya_klasoru = "C:/Users/q/Desktop/bılancolar/"

# Dosya klasöründeki tüm dosya isimlerini al
dosya_isimleri = os.listdir(dosya_klasoru)

# Tickers listesini oluştur
tickers = []

for dosya_adi in dosya_isimleri:
    # Dosya adının sonuna ".IS" ekleyerek ticker oluştur
    ticker = dosya_adi.split('.')[0] + '.IS'
    tickers.append(ticker)

# Create an empty DataFrame to store the signals for each stock
signals = pd.DataFrame(columns=['Ticker', 'CCI', 'RSI', 'MOM', 'STO', 'MACD', 'Histogram', 'SMA_20', 'BBand'])

def map_to_signal(values):
    """
    Map indicator values to "Buy", "Sell", or "Hold" signals.
    """
    buy_threshold = 0.6
    sell_threshold = 0.4
    return np.where(values > buy_threshold, 'Buy',
                   np.where(values < sell_threshold, 'Sell', 'Hold'))

# Loop through each stock ticker and calculate the technical indicators and signals
for ticker in tqdm(tickers, desc="Hisse Senetleri İşleniyor"):
    # Get the current date
    current_date = pd.Timestamp.now().strftime('%Y-%m-%d')

    # Download the data for the ticker
    data = yf.download(ticker, start="2022-01-01", end=current_date)

    # Calculate CCI
    data['CCI'] = ta.trend.cci(data['High'], data['Low'], data['Close'])

    # Calculate RSI
    data['RSI'] = ta.momentum.rsi(data['Close'])

    # Calculate MOM
    data['MOM'] = data['Close'].diff()

    # Calculate STO
    data['STO'] = ta.momentum.StochasticOscillator(data['High'], data['Low'], data['Close']).stoch()

    # Calculate MACD
    exp12, exp26 = ta.trend.ema_indicator(data['Close'], 12), ta.trend.ema_indicator(data['Close'], 26)
    data['MACD'] = exp12 - exp26
    data['Signal Line'] = ta.trend.ema_indicator(data['MACD'], 9)
    data['Histogram'] = data['MACD'] - data['Signal Line']

    # Calculate SMA
    data['SMA_20'] = ta.trend.sma_indicator(data['Close'], 20)

    # Calculate BBand
    data['BBand'] = ta.volatility.BollingerBands(data['Close']).bollinger_hband() - ta.volatility.BollingerBands(data['Close']).bollinger_lband()

    # Generate buy/sell signals based on technical indicators
    cci_signal = map_to_signal(data['CCI'])
    rsi_signal = map_to_signal(data['RSI'])
    mom_signal = map_to_signal(data['MOM'])
    sto_signal = map_to_signal(data['STO'])
    macd_signal = map_to_signal(data['Histogram'])
    sma_signal = map_to_signal(data['Close'] - data['SMA_20'])
    bband_signal = map_to_signal(data['Close'] - data['BBand'])

    # Add the stock ticker and signals to the signals DataFrame
    stock_signals = pd.DataFrame(index=data.index)
    stock_signals['Ticker'] = ticker[:-3] # Removethe '.IS' extension from the ticker
    stock_signals['CCI'] = data['CCI']
    stock_signals['RSI'] = data['RSI']
    stock_signals['MOM'] = data['MOM']
    stock_signals['STO'] = data['STO']
    stock_signals['MACD'] = data['MACD']
    stock_signals['Histogram'] = data['Histogram']
    stock_signals['SMA_20'] = data['SMA_20']
    stock_signals['BBand'] = data['BBand']
    stock_signals['CCI_Signal'] = cci_signal
    stock_signals['RSI_Signal'] = rsi_signal
    stock_signals['MOM_Signal'] = mom_signal
    stock_signals['STO_Signal'] = sto_signal
    stock_signals['MACD_Signal'] = macd_signal
    stock_signals['SMA_Signal'] = sma_signal
    stock_signals['BBand_Signal'] = bband_signal

    # Add the stock signals to the signalsDataFrame
    signals = pd.concat([signals, stock_signals.tail(1)], ignore_index=True)

# Sort the signals DataFrame in a descending order based on the CCI_Signal column
signals = signals.sort_values('CCI_Signal', ascending=False)

# Save the signals DataFrame to an Excel file on the desktop
signals.to_excel(r'C:/Users/q/Desktop/bt/BIST_Signals.xlsx',index=False)

for index, row in veriler.iterrows():
    veriler.at[index, 'CCI'] = pd.Series(cci_signal)[index]
    veriler.at[index, 'RSI'] = pd.Series(rsi_signal)[index]
    veriler.at[index, 'MOM'] = pd.Series(mom_signal)[index]
    veriler.at[index, 'STO'] = pd.Series(sto_signal)[index]
    veriler.at[index, 'MACD'] = pd.Series(macd_signal)[index]
    veriler.at[index, 'SMA'] = pd.Series(sma_signal)[index]
    veriler.at[index, 'BBand'] = pd.Series(bband_signal)[index]


import pandas as pd

# Dosya yollarını tanımlayalım
bist_signals_path = "C:/Users/q/Desktop/bt/BIST_Signals.xlsx"
finansal_veriler_path = "C:/Users/q/Desktop/bt/finansal_veriler_hesaplanmis.xlsx"
neva_data_path = "C:/Users/q/Desktop/bt/NEVAData.xlsx"

# Excel dosyalarını yükle
try:
    bist_signals = pd.read_excel(bist_signals_path)
    finansal_veriler = pd.read_excel(finansal_veriler_path)
except FileNotFoundError:
    print("Excel dosyaları bulunamadı.")
    exit()

# 'BIST_Signals' ve 'finansal_veriler' DataFrame'lerini birleştirme işlemi
merged_data = pd.merge(finansal_veriler, bist_signals, left_on='Hisse', right_on='Ticker', how='inner')

# Gereksiz sütunları kaldır
merged_data.drop(['Ticker'], axis=1, inplace=True)

# Yeni DataFrame'i 'NEVAData.xlsx' adında bir Excel dosyasına kaydet
merged_data.to_excel(neva_data_path, index=False)

print(f"Veriler başarıyla '{neva_data_path}' adlı Excel dosyasına kaydedildi.")
