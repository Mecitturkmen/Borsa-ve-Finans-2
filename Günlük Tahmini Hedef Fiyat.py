import yfinance as yf
import datetime
from datetime import date
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score, mean_absolute_error


def download_data(op, start_date, end_date):
    df = yf.download(op, start=start_date, end=end_date, progress=False)
    return df

def model_engine(model, num):
    # yalnızca kapanış fiyatını alma
    df = data[['Close']]
    # kapanış fiyatı tahmin edilen gün sayısına göre değiştiriliyor
    df['preds'] = df.Close.shift(-num)
    # verileri ölçeklendirme
    x = df.drop(['preds'], axis=1).values
    x = scaler.fit_transform(x)
    # son gün sayısı verilerinin saklanması
    x_forecast = x[-num:]
    # tahmin için gerekli değerleri seçme
    x = x[:-num]
    # preds sütununu alma
    y = df.preds.values
    # tahmin için gerekli değerleri seçme
    y = y[:-num]

    #verileri ayırma
    x_train, x_test, y_train, y_test = train_test_split(x, y, test_size=.2, random_state=7)
    # modeli tahmin etme
    model.fit(x_train, y_train)
    preds = model.predict(x_test)
    print(f'Doğruluk Oranı Tahmini : {r2_score(y_test, preds)}')
    # gün sayısına göre hisse senedi fiyatını tahmin etmek
    forecast_pred = model.predict(x_forecast)
    day = 10
    for i in forecast_pred:
        print(f'{day}. Gün İçin Tahmini Kapanış Fiyatı : {i}')
        day += 1


stock = "CCOLA.IS"
today = datetime.date.today()
duration = 3000
before = today - datetime.timedelta(days=duration)
start_date = before
# YARIN İÇİN -1 BUGÜN İÇİN 0 YAZILMALI
end_date = today-datetime.timedelta(days=0) 
print(end_date)
scaler = StandardScaler()

data = download_data(stock,start_date,end_date)

num = 1

engine = LinearRegression()
model_engine(engine, num)
