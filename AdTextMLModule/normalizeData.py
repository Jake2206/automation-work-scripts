import pandas as pd
from sklearn import preprocessing

myData = pd.read_excel('myData.xlsx', keep_default_na=False)
myData = myData.replace('', 0)

for i in range(len(myData['Amount Spent (USD)'])):
    totalSpent = myData['Amount Spent (USD)'][i]
    myData['Reach'][i] /= totalSpent
    myData['Impressions'][i] /= totalSpent
    myData['Frequency'][i] /= totalSpent
    myData['Clicks (All)'][i] /= totalSpent
    myData['Link Clicks'][i] /= totalSpent
    myData['Unique Link Clicks'][i] /= totalSpent
    myData['Amount Spent (USD)'][i] /= totalSpent

scaler = preprocessing.MinMaxScaler()
names = myData.columns
d = scaler.fit_transform(myData)
scaled_df = pd.DataFrame(d, columns = names)
scaled_df.head()

with pd.ExcelWriter('myNormalizedData.xlsx', engine='xlsxwriter') as writer:
    scaled_df.to_excel(writer, sheet_name='Sheet1')
