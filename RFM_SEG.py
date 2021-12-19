import pandas as pd
import datetime as dt

pd.set_option('display.max_columns', None)
pd.set_option('display.float_format', lambda x: '%.5f' % x)

DIR = 'online_retail_II.xlsx'  # read data file path
SHEET_NAME = 'Year 2010-2011'
LABELS = [1, 2, 3, 4, 5]
SEG_NUM = 5
SAVE = False
SEGMENT = 'loyal_customers'
FILE_NAME = 'loyal_customers_ids.xlsx'  # save file name as an excel

raw_df = pd.read_excel(DIR, sheet_name=SHEET_NAME)
df = raw_df.copy()

df.dropna(inplace=True)
df = df[~df['Invoice'].str.startswith('C', na=False)]
df['TotalPrice'] = df['Quantity'] * df['Price']

today_date = dt.datetime(2011, 12, 11)
rfm = df.groupby('Customer ID').agg(
    {'InvoiceDate': lambda date: (today_date - date.max()).days,  # calculates recency, frequency, monetary
     'Invoice': lambda invoice: invoice.nunique(),
     'TotalPrice': lambda price: price.sum()})

rfm.columns = ['recency', 'frequency', 'monetary']
rfm = rfm[rfm['monetary'] > 0]  # should be over 0

# convert to scores
rfm['recency_score'] = pd.qcut(rfm['recency'], SEG_NUM, labels=LABELS[::-1])
rfm['frequency_score'] = pd.qcut(rfm['frequency'].rank(method='first'), SEG_NUM, labels=LABELS)
rfm['monetary_score'] = pd.qcut(rfm['monetary'], SEG_NUM, labels=LABELS)
rfm['RFM_SCORE'] = rfm['recency_score'].astype('str') + rfm['frequency_score'].astype('str') # calculates rfm score by recency & frequency

# segmentation map
seg_map = {
    r'[1-2][1-2]': 'hibernating',
    r'[1-2][3-4]': 'at_Risk',
    r'[1-2]5': 'cant_loose',
    r'3[1-2]': 'about_to_sleep',
    r'33': 'need_attention',
    r'[3-4][4-5]': 'loyal_customers',
    r'41': 'promising',
    r'51': 'new_customers',
    r'[4-5][2-3]': 'potential_loyalists',
    r'5[4-5]': 'champions'
}
rfm['segment'] = rfm['RFM_SCORE'].replace(seg_map, regex=True)

# statistical approach
# rfm[["segment", "recency", "frequency", "monetary"]].groupby("segment").agg(["mean", "count"])

# save as a file for loyal customers
if SAVE:
    loyal_customers_ids = pd.DataFrame(rfm[rfm['segment'] == SEGMENT].index)
    loyal_customers_ids.to_excel(FILE_NAME, index=False)
