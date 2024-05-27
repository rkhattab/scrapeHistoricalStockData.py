import yfinance as yf
import pandas as pd
from ta import add_all_ta_features

# Fetch historical data for a specific ticker
data = yf.download('AAPL', start='2020-01-01', end='2024-05-24', interval='1d', auto_adjust=True, actions=True)

# Calculate RSI
delta = data['Close'].diff()
gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
rs = gain / loss
rsi = 100 - (100 / (1 + rs))

# Calculate MACD
short_ema = data['Close'].ewm(span=12, adjust=False).mean()
long_ema = data['Close'].ewm(span=26, adjust=False).mean()
macd = short_ema - long_ema
signal_line = macd.ewm(span=9, adjust=False).mean()

# Add RSI and MACD columns to the DataFrame
data['RSI'] = rsi
data['MACD'] = macd
data['Signal Line'] = signal_line

# Extract date from index and add as a column
data['Date'] = data.index

# Reorder columns to have Date, Open, High, Low, Close, Volume, RSI, MACD, and Signal Line
data = data[['Date', 'Open', 'High', 'Low', 'Close', 'Volume', 'RSI', 'MACD', 'Signal Line']]

# Save DataFrame to an Excel File
data.to_excel('aapl20_data_daily.xlsx', index=False)

# Read the Excel File into a DataFrame
df_read = pd.read_excel('aapl20_data_daily.xlsx')

# Print the DataFrame
print(df_read)
