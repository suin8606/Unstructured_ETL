```python
# pip install python-barcode
# pip install "python-barcode[images]"
import pandas as pd
import numpy as np
import pyodbc
import openpyxl
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from pathlib import Path
from openpyxl import load_workbook, styles, formatting
import sys
import os
```


```python
# import barcode
# from barcode.writer import ImageWriter
# code='725272730706'
# sample_barcode=barcode.get('upca',code,writer=ImageWriter())
# generated_filename=sample_barcode.save('uuu')
# print('upc-a'+generated_filename)
```


```python
server = '10.1.3.25' 
database = 'KIRA' 
username = 'kiradba' 
password = 'Kiss!234!' 
connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)
print("Connection Established:")
```

    Connection Established:
    


```python
df=pd.read_sql('''
WITH T1 as (
SELECT material, description, ip, ct, nsp, srp, upc
FROM [ivy.mm.dim.mtrl]
WHERE ivykiss = 'X' and ms in ('01','41','91','N1','D1'))
SELECT *
FROM T1
WHERE nsp is not null
''',con=engine)
df=df.astype({"material":"str", "description":"str","ip":"int","ct":"int","nsp":"float","srp":"float","upc":"str"}).sort_values(by=["upc"],ascending=True)
```


```python
df.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>material</th>
      <th>description</th>
      <th>ip</th>
      <th>ct</th>
      <th>nsp</th>
      <th>srp</th>
      <th>upc</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>1511</th>
      <td>ALL02</td>
      <td>RK Auto Lip Liner-Brown</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002738</td>
    </tr>
    <tr>
      <th>1506</th>
      <td>ALL05</td>
      <td>RK Auto Lip Liner-Cappuccino</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002769</td>
    </tr>
    <tr>
      <th>1507</th>
      <td>ALL06</td>
      <td>RK Auto Lip Liner-Cocoa</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002776</td>
    </tr>
    <tr>
      <th>1508</th>
      <td>ALL07</td>
      <td>RK Auto Lip Liner-Black</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002783</td>
    </tr>
    <tr>
      <th>1509</th>
      <td>ALL11</td>
      <td>RK Auto Lip Liner-Plum</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002820</td>
    </tr>
  </tbody>
</table>
</div>




```python
df.tail()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>material</th>
      <th>description</th>
      <th>ip</th>
      <th>ct</th>
      <th>nsp</th>
      <th>srp</th>
      <th>upc</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>1385</th>
      <td>KID03Y1</td>
      <td>IEK 3D Collection Double 03</td>
      <td>3</td>
      <td>288</td>
      <td>2.0</td>
      <td>3.99</td>
      <td>731509870169</td>
    </tr>
    <tr>
      <th>1751</th>
      <td>FCT01</td>
      <td>KS Gel Fantasy Toenails- This is Classic</td>
      <td>2</td>
      <td>36</td>
      <td>4.8</td>
      <td>7.99</td>
      <td>731509876055</td>
    </tr>
    <tr>
      <th>1330</th>
      <td>KNAR02</td>
      <td>KS Nail Art Rhinestones - Multi Shapes</td>
      <td>2</td>
      <td>36</td>
      <td>4.5</td>
      <td>9.00</td>
      <td>731509877960</td>
    </tr>
    <tr>
      <th>1473</th>
      <td>BN05</td>
      <td>KS Bare-But-Better Nails - Berry Nude</td>
      <td>2</td>
      <td>36</td>
      <td>4.5</td>
      <td>8.99</td>
      <td>731509880106</td>
    </tr>
    <tr>
      <th>2881</th>
      <td>LGC01</td>
      <td>KC Temp Blend-away - Jet Black</td>
      <td>3</td>
      <td>120</td>
      <td>2.5</td>
      <td>4.99</td>
      <td>731509999952</td>
    </tr>
  </tbody>
</table>
</div>




```python
df.info()
```

    <class 'pandas.core.frame.DataFrame'>
    Int64Index: 2935 entries, 1511 to 2881
    Data columns (total 7 columns):
     #   Column       Non-Null Count  Dtype  
    ---  ------       --------------  -----  
     0   material     2935 non-null   object 
     1   description  2935 non-null   object 
     2   ip           2935 non-null   int32  
     3   ct           2935 non-null   int32  
     4   nsp          2935 non-null   float64
     5   srp          2916 non-null   float64
     6   upc          2935 non-null   object 
    dtypes: float64(2), int32(2), object(3)
    memory usage: 160.5+ KB
    


```python
# convert it as excel file.
df.to_excel(r"Y:\OM ONLY_Shared Documents\5 Reports with Power Query\Reports\IVYKISS UPC Report\2. UPC_LIST\UPC_LIST.xlsx",index=False)
```


```python
previous_df=pd.read_excel(r"Y:\OM ONLY_Shared Documents\5 Reports with Power Query\Reports\IVYKISS UPC Report\2. UPC_LIST\UPC_LIST_091222.xlsx")
previous_df.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>material</th>
      <th>description</th>
      <th>ip</th>
      <th>ct</th>
      <th>nsp</th>
      <th>srp</th>
      <th>upc</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>ALL02</td>
      <td>RK Auto Lip Liner-Brown</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002738</td>
    </tr>
    <tr>
      <th>1</th>
      <td>ALL05</td>
      <td>RK Auto Lip Liner-Cappuccino</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002769</td>
    </tr>
    <tr>
      <th>2</th>
      <td>ALL06</td>
      <td>RK Auto Lip Liner-Cocoa</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002776</td>
    </tr>
    <tr>
      <th>3</th>
      <td>ALL07</td>
      <td>RK Auto Lip Liner-Black</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002783</td>
    </tr>
    <tr>
      <th>4</th>
      <td>ALL11</td>
      <td>RK Auto Lip Liner-Plum</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002820</td>
    </tr>
  </tbody>
</table>
</div>




```python
thisweek_df=pd.read_excel(r"Y:\OM ONLY_Shared Documents\5 Reports with Power Query\Reports\IVYKISS UPC Report\2. UPC_LIST\UPC_LIST_091922.xlsx")
thisweek_df.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>material</th>
      <th>description</th>
      <th>ip</th>
      <th>ct</th>
      <th>nsp</th>
      <th>srp</th>
      <th>upc</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>ALL02</td>
      <td>RK Auto Lip Liner-Brown</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002738</td>
    </tr>
    <tr>
      <th>1</th>
      <td>ALL05</td>
      <td>RK Auto Lip Liner-Cappuccino</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002769</td>
    </tr>
    <tr>
      <th>2</th>
      <td>ALL06</td>
      <td>RK Auto Lip Liner-Cocoa</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002776</td>
    </tr>
    <tr>
      <th>3</th>
      <td>ALL07</td>
      <td>RK Auto Lip Liner-Black</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002783</td>
    </tr>
    <tr>
      <th>4</th>
      <td>ALL11</td>
      <td>RK Auto Lip Liner-Plum</td>
      <td>12</td>
      <td>288</td>
      <td>0.55</td>
      <td>0.99</td>
      <td>649674002820</td>
    </tr>
  </tbody>
</table>
</div>




```python
difference=list(set(previous_df["material"]) - set(thisweek_df["material"])) # a way to inspect the differnces in material between previous week and this week.
print('The Different Material List between last & this week: {}'.format(difference))
```

    The Different Material List between last & this week: ['KEG100H', 'KEG025F', 'KEG100C', 'KEG100D', 'KEG025E', 'KEG100L', 'KEG100E']
    


```python
os.getcwd() # currently directory location
newpath = r'Y:\OM ONLY_Shared Documents\5 Reports with Power Query\Reports\IVYKISS UPC Report\3. Barcode_by_Python'
os.chdir(newpath) # save it as current directory.
```


```python
import barcode
from barcode import UPCA
from barcode.writer import ImageWriter
for i,y in zip(df["material"], df["upc"]):
    with open(str(y) + ".png", "wb") as f:
        UPCA(str(y), writer=ImageWriter()).write(f)
```
