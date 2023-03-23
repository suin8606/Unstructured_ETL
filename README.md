```python
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import datetime as dt
import re
from dateutil.relativedelta import relativedelta
from string import ascii_uppercase
```


```python
# Loading the raw data. Specified a column name as numbers in object to see easily.
df=pd.read_excel("/Users/suinkim/Downloads/TEST.xlsx",usecols=list(range(1,25,+1)),header=None)
```


```python
df3 = df.fillna("")
df3
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
      <th>1</th>
      <th>2</th>
      <th>3</th>
      <th>4</th>
      <th>5</th>
      <th>6</th>
      <th>7</th>
      <th>8</th>
      <th>9</th>
      <th>10</th>
      <th>...</th>
      <th>15</th>
      <th>16</th>
      <th>17</th>
      <th>18</th>
      <th>19</th>
      <th>20</th>
      <th>21</th>
      <th>22</th>
      <th>23</th>
      <th>24</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>1</th>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>2</th>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>3</th>
      <td></td>
      <td></td>
      <td>YEAR 1</td>
      <td></td>
      <td></td>
      <td>FULL YEAR</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>YEAR 3</td>
      <td></td>
      <td></td>
      <td>FULL YEAR</td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>4</th>
      <td></td>
      <td></td>
      <td>Fincial Metric</td>
      <td></td>
      <td></td>
      <td>Weight</td>
      <td>Times</td>
      <td>outcome</td>
      <td></td>
      <td></td>
      <td>...</td>
      <td>Times</td>
      <td>outcome</td>
      <td></td>
      <td></td>
      <td>Fincial Metric</td>
      <td></td>
      <td></td>
      <td>Weight</td>
      <td>Times</td>
      <td>outcome</td>
    </tr>
    <tr>
      <th>5</th>
      <td></td>
      <td></td>
      <td>Fincial Metric 1</td>
      <td></td>
      <td></td>
      <td>100</td>
      <td>1</td>
      <td>Y</td>
      <td></td>
      <td></td>
      <td>...</td>
      <td>10</td>
      <td>Y</td>
      <td></td>
      <td></td>
      <td>Fincial Metric 1</td>
      <td></td>
      <td></td>
      <td>21213</td>
      <td>50</td>
      <td>Y</td>
    </tr>
    <tr>
      <th>6</th>
      <td></td>
      <td></td>
      <td>Fincial Metric 2</td>
      <td></td>
      <td></td>
      <td>200</td>
      <td>2</td>
      <td>N</td>
      <td></td>
      <td></td>
      <td>...</td>
      <td>20</td>
      <td>N</td>
      <td></td>
      <td></td>
      <td>Fincial Metric 2</td>
      <td></td>
      <td></td>
      <td>55567</td>
      <td>60</td>
      <td>N</td>
    </tr>
    <tr>
      <th>7</th>
      <td></td>
      <td></td>
      <td>Fincial Metric 3</td>
      <td></td>
      <td></td>
      <td>300</td>
      <td>3</td>
      <td>Y</td>
      <td></td>
      <td></td>
      <td>...</td>
      <td>30</td>
      <td>Y</td>
      <td></td>
      <td></td>
      <td>Fincial Metric 3</td>
      <td></td>
      <td></td>
      <td>89890</td>
      <td>70</td>
      <td>Y</td>
    </tr>
    <tr>
      <th>8</th>
      <td></td>
      <td></td>
      <td>Fincial Metric 4</td>
      <td></td>
      <td></td>
      <td>400</td>
      <td>4</td>
      <td>N</td>
      <td></td>
      <td></td>
      <td>...</td>
      <td>40</td>
      <td>N</td>
      <td></td>
      <td></td>
      <td>Fincial Metric 4</td>
      <td></td>
      <td></td>
      <td>40524</td>
      <td>80</td>
      <td>N</td>
    </tr>
    <tr>
      <th>9</th>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>10</th>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>11</th>
      <td></td>
      <td></td>
      <td>Comment</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>12</th>
      <td></td>
      <td></td>
      <td>Please review something</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>13</th>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>14</th>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>Score</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>15</th>
      <td></td>
      <td></td>
      <td>Result</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>A-2</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <th>16</th>
      <td></td>
      <td></td>
      <td>Passed</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>...</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
  </tbody>
</table>
<p>17 rows × 24 columns</p>
</div>




```python
df2=df.rename(columns={3:'Comment',19:'Score'})
```


```python
# Filter only relevent data
df2=df2[["Comment","Score"]]
```


```python
df2=df2[11:]
```


```python
# Getting raw data as a set of list since columns and raw are unstructured 
lst = df2["Comment"].tolist() + df2["Score"].tolist()
```


```python
# Filtering NAN out
cleanedlst=[x for x in lst if str(x) != 'nan']
```


```python
# Setting columns
col_names=cleanedlst[::2]
```


```python
# Setting values
val = cleanedlst[1::2]
```


```python
# Expressing in DataFrame
data={
    'Comment':val[col_names.index('Comment')],
    'Result':val[col_names.index('Result')],
    "Score":val[col_names.index("Score")]
}
```


```python
df2=pd.DataFrame(data,index=[0])
```


```python
# Arrange an ID to join with Years table. 
df2["id"]='1'
```


```python
df2
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
      <th>Comment</th>
      <th>Result</th>
      <th>Score</th>
      <th>id</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>1</td>
    </tr>
  </tbody>
</table>
</div>




```python
# Full Year 1
#df = df.fillna("")
df.dropna(how='all')
df_year1=df.loc[4:8,3:8]
df_year1=df_year1[[3,6,7,8]]
df_year1.columns=df_year1.iloc[0]
df_year1["Year"]='1'
df_year1=df_year1[["Year","Fincial Metric","Weight","Times","outcome"]]
df_year1=df_year1[1:]
```


```python
# Full Year 2
df_year2=df.loc[4:8,10:16]
df_year2=df_year2[[11,14,15,16]]
df_year2.columns=df_year2.iloc[0]
df_year2["Year"]='2'
df_year2=df_year2[["Year","Fincial Metric","Weight","Times","outcome"]]
df_year2=df_year2[1:]
```


```python
# Full Year 3
df_year3=df.loc[4:8,18:24]
df_year3=df_year3[[19,22,23,24]]
df_year3.columns=df_year3.iloc[0]
df_year3["Year"]='3'
df_year3=df_year3[["Year","Fincial Metric","Weight","Times","outcome"]]
df_year3=df_year3[1:]
```


```python
# Concat year1, year2 and year3. Can express year1,year2,year3 seperately as your favor.
df_final=pd.concat([df_year1,df_year2,df_year3],ignore_index=True)
```


```python
# Arrange an ID to join with other table. This can be also a sheet number.
df_final["id"]='1'
```


```python
df_final
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
      <th>4</th>
      <th>Year</th>
      <th>Fincial Metric</th>
      <th>Weight</th>
      <th>Times</th>
      <th>outcome</th>
      <th>id</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>1</td>
      <td>Fincial Metric 1</td>
      <td>100</td>
      <td>1</td>
      <td>Y</td>
      <td>1</td>
    </tr>
    <tr>
      <th>1</th>
      <td>1</td>
      <td>Fincial Metric 2</td>
      <td>200</td>
      <td>2</td>
      <td>N</td>
      <td>1</td>
    </tr>
    <tr>
      <th>2</th>
      <td>1</td>
      <td>Fincial Metric 3</td>
      <td>300</td>
      <td>3</td>
      <td>Y</td>
      <td>1</td>
    </tr>
    <tr>
      <th>3</th>
      <td>1</td>
      <td>Fincial Metric 4</td>
      <td>400</td>
      <td>4</td>
      <td>N</td>
      <td>1</td>
    </tr>
    <tr>
      <th>4</th>
      <td>2</td>
      <td>Fincial Metric 1</td>
      <td>6565</td>
      <td>10</td>
      <td>Y</td>
      <td>1</td>
    </tr>
    <tr>
      <th>5</th>
      <td>2</td>
      <td>Fincial Metric 2</td>
      <td>3900</td>
      <td>20</td>
      <td>N</td>
      <td>1</td>
    </tr>
    <tr>
      <th>6</th>
      <td>2</td>
      <td>Fincial Metric 3</td>
      <td>5621</td>
      <td>30</td>
      <td>Y</td>
      <td>1</td>
    </tr>
    <tr>
      <th>7</th>
      <td>2</td>
      <td>Fincial Metric 4</td>
      <td>6049</td>
      <td>40</td>
      <td>N</td>
      <td>1</td>
    </tr>
    <tr>
      <th>8</th>
      <td>3</td>
      <td>Fincial Metric 1</td>
      <td>21213</td>
      <td>50</td>
      <td>Y</td>
      <td>1</td>
    </tr>
    <tr>
      <th>9</th>
      <td>3</td>
      <td>Fincial Metric 2</td>
      <td>55567</td>
      <td>60</td>
      <td>N</td>
      <td>1</td>
    </tr>
    <tr>
      <th>10</th>
      <td>3</td>
      <td>Fincial Metric 3</td>
      <td>89890</td>
      <td>70</td>
      <td>Y</td>
      <td>1</td>
    </tr>
    <tr>
      <th>11</th>
      <td>3</td>
      <td>Fincial Metric 4</td>
      <td>40524</td>
      <td>80</td>
      <td>N</td>
      <td>1</td>
    </tr>
  </tbody>
</table>
</div>




```python
# Create a new table by merging a common table, ID.
df_final=df_final.merge(df2,left_on="id", right_on='id')
```


```python
# Setting Primary Key
df_final["index"]=range(1,len(df_final)+1)
```


```python
df_final
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
      <th>Year</th>
      <th>Fincial Metric</th>
      <th>Weight</th>
      <th>Times</th>
      <th>outcome</th>
      <th>id</th>
      <th>Comment</th>
      <th>Result</th>
      <th>Score</th>
      <th>index</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>1</td>
      <td>Fincial Metric 1</td>
      <td>100</td>
      <td>1</td>
      <td>Y</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>1</td>
    </tr>
    <tr>
      <th>1</th>
      <td>1</td>
      <td>Fincial Metric 2</td>
      <td>200</td>
      <td>2</td>
      <td>N</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>2</td>
    </tr>
    <tr>
      <th>2</th>
      <td>1</td>
      <td>Fincial Metric 3</td>
      <td>300</td>
      <td>3</td>
      <td>Y</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>3</td>
    </tr>
    <tr>
      <th>3</th>
      <td>1</td>
      <td>Fincial Metric 4</td>
      <td>400</td>
      <td>4</td>
      <td>N</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>4</td>
    </tr>
    <tr>
      <th>4</th>
      <td>2</td>
      <td>Fincial Metric 1</td>
      <td>6565</td>
      <td>10</td>
      <td>Y</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>5</td>
    </tr>
    <tr>
      <th>5</th>
      <td>2</td>
      <td>Fincial Metric 2</td>
      <td>3900</td>
      <td>20</td>
      <td>N</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>6</td>
    </tr>
    <tr>
      <th>6</th>
      <td>2</td>
      <td>Fincial Metric 3</td>
      <td>5621</td>
      <td>30</td>
      <td>Y</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>7</td>
    </tr>
    <tr>
      <th>7</th>
      <td>2</td>
      <td>Fincial Metric 4</td>
      <td>6049</td>
      <td>40</td>
      <td>N</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>8</td>
    </tr>
    <tr>
      <th>8</th>
      <td>3</td>
      <td>Fincial Metric 1</td>
      <td>21213</td>
      <td>50</td>
      <td>Y</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>9</td>
    </tr>
    <tr>
      <th>9</th>
      <td>3</td>
      <td>Fincial Metric 2</td>
      <td>55567</td>
      <td>60</td>
      <td>N</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>10</td>
    </tr>
    <tr>
      <th>10</th>
      <td>3</td>
      <td>Fincial Metric 3</td>
      <td>89890</td>
      <td>70</td>
      <td>Y</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>11</td>
    </tr>
    <tr>
      <th>11</th>
      <td>3</td>
      <td>Fincial Metric 4</td>
      <td>40524</td>
      <td>80</td>
      <td>N</td>
      <td>1</td>
      <td>Please review something</td>
      <td>Passed</td>
      <td>A-2</td>
      <td>12</td>
    </tr>
  </tbody>
</table>
</div>




```python
# Connect to SQL 
server = '10.1.3.25' 
database = '**' 
username = '**' 
password = '**' 
connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)
print("connected")
```


```python
# Transfer into Database
df_final.to_sql('fact.ETL', engine, schema = "dbo", if_exists='append', index=False, chunksize=10000)
print("\n Successfully Transported")
```
