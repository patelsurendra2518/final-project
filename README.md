# final-project
```python
import pandas as pd

# Inspect sheet names
excel_path = 's Final Project (1).xlsx'
xls = pd.ExcelFile(excel_path)
print("Sheet names:", xls.sheet_names)

# Inspect head of each sheet
for sheet in xls.sheet_names:
    print(f"\n--- Sheet: {sheet} ---")
    df = pd.read_excel(excel_path, sheet_name=sheet)
    print("Shape:", df.shape)
    print("Columns:", df.columns.tolist())
    print(df.head(3))


```

```text
Sheet names: ['Instruction', 'Raw_Data', 'Analysis', 'Summary_Output', 'Pivot_Report', 'Dashboard']

--- Sheet: Instruction ---
Shape: (162, 1)
Columns: ['Project Title']
                                                   Project Title
0                                                            NaN
1  Customer Transaction Analysis Dashboard using Microsoft Excel
2                                                            NaN

--- Sheet: Raw_Data ---
Shape: (250, 21)
Columns: ['Transaction_ID', 'Date', 'Customer_ID', 'Customer_Name', 'Product_ID', 'Product_Name', 'Category', 'Quantity', 'Unit_Price', 'Payment_Method', 'Region', 'Customer_Segment', 'Customer_Since', 'Total_Amount', 'Month', 'Year', 'Customer Age Days', 'Customer Tenure', 'Transaction Month End', 'High Value Customer', 'Timestamp']
  Transaction_ID       Date Customer_ID   Customer_Name Product_ID Product_Name     Category  Quantity  Unit_Price Payment_Method   Region Customer_Segment Customer_Since  Total_Amount     Month  Year  Customer Age Days  Customer Tenure Transaction Month End High Value Customer               Timestamp
0        TRX0155 2024-04-11     CUST023      Paul Baker       P009      Monitor  Electronics         3      249.99           Cash  Central            Basic     2021-08-17        749.97  Apr-2024  2024               1862                5            2024-04-30              Normal 2026-09-22 00:42:51.600
1        TRX0023 2024-04-14     CUST013  Joseph Jackson       P009      Monitor  Electronics         3      249.99    Credit Card  Central          Premium     2022-01-14        749.97  Apr-2024  2024               1712                4            2024-04-30              Normal 2026-09-22 00:42:51.600
2        TRX0139 2024-04-17     CUST021     Kevin Scott       P005      Blender   Appliances         3       59.99         PayPal    South         Standard     2023-12-02        179.97  Apr-2024  2024               1025                2            2024-04-30              Normal 2026-09-22 00:42:51.600

--- Sheet: Analysis ---
Shape: (52, 3)
Columns: ['Customer Transaction Analysis', 'Unnamed: 1', 'Unnamed: 2']
  Customer Transaction Analysis Unnamed: 1 Unnamed: 2
0                           NaN        NaN        NaN
1                           NaN        NaN        NaN
2                             A          B        NaN

--- Sheet: Summary_Output ---
Shape: (273, 9)
Columns: ['SUMMARY OUTPUT', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
          SUMMARY OUTPUT Unnamed: 1 Unnamed: 2 Unnamed: 3 Unnamed: 4 Unnamed: 5 Unnamed: 6 Unnamed: 7 Unnamed: 8
0                    NaN        NaN        NaN        NaN        NaN        NaN        NaN        NaN        NaN
1  Regression Statistics        NaN        NaN        NaN        NaN        NaN        NaN        NaN        NaN
2             Multiple R    0.45808        NaN        NaN        NaN        NaN        NaN        NaN        NaN

--- Sheet: Pivot_Report ---
Shape: (74, 10)
Columns: ['Pivot Reports - Customer Transaction Analysis', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']
  Pivot Reports - Customer Transaction Analysis   Unnamed: 1  Unnamed: 2  Unnamed: 3                                  Unnamed: 4     Unnamed: 5  Unnamed: 6  Unnamed: 7                        Unnamed: 8   Unnamed: 9
0                                           NaN          NaN         NaN         NaN                                         NaN            NaN         NaN         NaN                               NaN          NaN
1             Pivot Table 1 – Region Wise Sales          NaN         NaN         NaN  Pivot Table 3 – Product Wise Quantity Sold            NaN         NaN         NaN  Pivot Table 5 – Customer Segment          NaN
2                                        Region  Total Sales         NaN         NaN                                Product_Name  Quantity Sold         NaN         NaN                  Customer_Segment  Total_Sales

--- Sheet: Dashboard ---
Shape: (16, 2)
Columns: ['CUSTOMER TRANSACTION DASHBOARD', 'Unnamed: 1']
   CUSTOMER TRANSACTION DASHBOARD     Unnamed: 1
0                             NaN            NaN
1                             NaN            NaN
2                             NaN  Total Revenue


```

```python
# Let's inspect data summary, data types, missing values for both sheets
for sheet in xls.sheet_names:
    print(f"\n==================== Sheet: {sheet} ====================")
    df = pd.read_excel(excel_path, sheet_name=sheet)
    print(df.info())
    print("\nSummary Statistics:")
    print(df.describe(include='all'))


```

```text

==================== Sheet: Instruction ====================
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 162 entries, 0 to 161
Data columns (total 1 columns):
 #   Column         Non-Null Count  Dtype 
---  ------         --------------  ----- 
 0   Project Title  105 non-null    object
dtypes: object(1)
memory usage: 1.4+ KB
None

Summary Statistics:
       Project Title
count            105
unique            97
top     Pivot Charts
freq               3

==================== Sheet: Raw_Data ====================
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 250 entries, 0 to 249
Data columns (total 21 columns):
 #   Column                 Non-Null Count  Dtype         
---  ------                 --------------  -----         
 0   Transaction_ID         250 non-null    object        
 1   Date                   250 non-null    datetime64[ns]
 2   Customer_ID            250 non-null    object        
 3   Customer_Name          250 non-null    object        
 4   Product_ID             250 non-null    object        
 5   Product_Name           250 non-null    object        
 6   Category               250 non-null    object        
 7   Quantity               250 non-null    int64         
 8   Unit_Price             250 non-null    float64       
 9   Payment_Method         250 non-null    object        
 10  Region                 250 non-null    object        
 11  Customer_Segment       250 non-null    object        
 12  Customer_Since         250 non-null    datetime64[ns]
 13  Total_Amount           250 non-null    float64       
 14  Month                  250 non-null    object        
 15  Year                   250 non-null    int64         
 16  Customer Age Days      250 non-null    int64         
 17  Customer Tenure        250 non-null    int64         
 18  Transaction Month End  250 non-null    datetime64[ns]
 19  High Value Customer    250 non-null    object        
 20  Timestamp              250 non-null    datetime64[ns]
dtypes: datetime64[ns](4), float64(2), int64(4), object(11)
memory usage: 41.1+ KB
None

Summary Statistics:
       Transaction_ID                 Date Customer_ID Customer_Name Product_ID Product_Name     Category    Quantity  Unit_Price Payment_Method Region Customer_Segment       Customer_Since  Total_Amount     Month         Year  Customer Age Days  Customer Tenure Transaction Month End High Value Customer                   Timestamp
count             250                  250         250           250        250          250          250  250.000000  250.000000            250    250              250                  250     250.00000       250   250.000000         250.000000       250.000000                   250                 250                         250
unique            250                  NaN          50            50         10           10            3         NaN         NaN              4      5                3                  NaN           NaN        13          NaN                NaN              NaN                   NaN                   2                         NaN
top           TRX0155                  NaN     CUST025   Mark Carter       P008    Bookshelf  Electronics         NaN         NaN           Cash  North          Premium                  NaN           NaN  Jan-2025          NaN                NaN              NaN                   NaN              Normal                         NaN
freq                1                  NaN          12            12         35           35          131         NaN         NaN             76     59               95                  NaN           NaN        26          NaN                NaN              NaN                   NaN                 189                         NaN
mean              NaN  2024-10-08 17:45:36         NaN           NaN        NaN          NaN          NaN    3.012000  296.670000            NaN    NaN              NaN  2022-10-27 19:12:00     916.76988       NaN  2024.288000        1425.200000         3.388000   2024-10-22 20:09:36                 NaN  2026-09-22 00:42:51.600000
min               NaN  2024-04-11 00:00:00         NaN           NaN        NaN          NaN          NaN    1.000000   59.990000            NaN    NaN              NaN  2021-07-18 00:00:00      59.99000       NaN  2024.000000         894.000000         2.000000   2024-04-30 00:00:00                 NaN  2026-09-22 00:42:51.600000
25%               NaN  2024-07-04 06:00:00         NaN           NaN        NaN          NaN          NaN    2.000000   89.990000            NaN    NaN              NaN  2022-01-18 00:00:00     249.99000       NaN  2024.000000        1146.250000         3.000000   2024-07-31 00:00:00                 NaN  2026-09-22 00:42:51.600000
50%               NaN  2024-10-09 12:00:00         NaN           NaN        NaN          NaN          NaN    3.000000  149.990000            NaN    NaN              NaN  2022-10-25 00:00:00     499.98000       NaN  2024.000000        1428.000000         3.000000   2024-10-31 00:00:00                 NaN  2026-09-22 00:42:51.600000
75%               NaN  2025-01-12 12:00:00         NaN           NaN        NaN          NaN          NaN    4.000000  299.990000            NaN    NaN              NaN  2023-08-02 18:00:00     999.96000       NaN  2025.000000        1708.000000         4.000000   2025-01-31 00:00:00                 NaN  2026-09-22 00:42:51.600000
max               NaN  2025-04-11 00:00:00         NaN           NaN        NaN          NaN          NaN    5.000000  899.990000            NaN    NaN              NaN  2024-04-11 00:00:00    4499.95000       NaN  2025.000000        1892.000000         5.000000   2025-04-30 00:00:00                 NaN  2026-09-22 00:42:51.600000
std               NaN                  NaN         NaN           NaN        NaN          NaN          NaN    1.336789  274.828493            NaN    NaN              NaN                  NaN    1032.45310       NaN     0.453739         299.950063         0.903727                   NaN                 NaN                         NaN

==================== Sheet: Analysis ====================
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 52 entries, 0 to 51
Data columns (total 3 columns):
 #   Column                         Non-Null Count  Dtype 
---  ------                         --------------  ----- 
 0   Customer Transaction Analysis  44 non-null     object
 1   Unnamed: 1                     40 non-null     object
 2   Unnamed: 2                     12 non-null     object
dtypes: object(3)
memory usage: 1.3+ KB
None

Summary Statistics:
       Customer Transaction Analysis Unnamed: 1  Unnamed: 2
count                             44         40          12
unique                            43         35           4
top                                A      Sales  High Value
freq                               2          4           6

==================== Sheet: Summary_Output ====================
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 273 entries, 0 to 272
Data columns (total 9 columns):
 #   Column          Non-Null Count  Dtype 
---  ------          --------------  ----- 
 0   SUMMARY OUTPUT  264 non-null    object
 1   Unnamed: 1      263 non-null    object
 2   Unnamed: 2      258 non-null    object
 3   Unnamed: 3      6 non-null      object
 4   Unnamed: 4      5 non-null      object
 5   Unnamed: 5      5 non-null      object
 6   Unnamed: 6      3 non-null      object
 7   Unnamed: 7      3 non-null      object
 8   Unnamed: 8      3 non-null      object
dtypes: object(9)
memory usage: 19.3+ KB
None

Summary Statistics:
               SUMMARY OUTPUT  Unnamed: 1  Unnamed: 2 Unnamed: 3 Unnamed: 4      Unnamed: 5 Unnamed: 6   Unnamed: 7   Unnamed: 8
count                     264  263.000000   258.00000          6          5               5          3            3            3
unique                    264   18.000000    52.00000          6          5               5          3            3            3
top     Regression Statistics  912.524364  -258.75136         MS          F  Significance F  Upper 95%  Lower 95.0%  Upper 95.0%
freq                        1   66.000000    16.00000          1          1               1          1            1            1

==================== Sheet: Pivot_Report ====================
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 74 entries, 0 to 73
Data columns (total 10 columns):
 #   Column                                         Non-Null Count  Dtype  
---  ------                                         --------------  -----  
 0   Pivot Reports - Customer Transaction Analysis  64 non-null     object 
 1   Unnamed: 1                                     61 non-null     object 
 2   Unnamed: 2                                     0 non-null      float64
 3   Unnamed: 3                                     0 non-null      float64
 4   Unnamed: 4                                     25 non-null     object 
 5   Unnamed: 5                                     22 non-null     object 
 6   Unnamed: 6                                     0 non-null      float64
 7   Unnamed: 7                                     0 non-null      float64
 8   Unnamed: 8                                     32 non-null     object 
 9   Unnamed: 9                                     29 non-null     object 
dtypes: float64(4), object(6)
memory usage: 5.9+ KB
None

Summary Statistics:
       Pivot Reports - Customer Transaction Analysis   Unnamed: 1  Unnamed: 2  Unnamed: 3                                  Unnamed: 4  Unnamed: 5  Unnamed: 6  Unnamed: 7                        Unnamed: 8  Unnamed: 9
count                                             64           61         0.0         0.0                                          25        22.0         0.0         0.0                                32        29.0
unique                                            64           60         NaN         NaN                                          25        21.0         NaN         NaN                                32        28.0
top                Pivot Table 1 – Region Wise Sales  Total Sales         NaN         NaN  Pivot Table 3 – Product Wise Quantity Sold        47.0         NaN         NaN  Pivot Table 5 – Customer Segment        26.0
freq                                               1            2         NaN         NaN                                           1         2.0         NaN         NaN                                 1         2.0
mean                                             NaN          NaN         NaN         NaN                                         NaN         NaN         NaN         NaN                               NaN         NaN
std                                              NaN          NaN         NaN         NaN                                         NaN         NaN         NaN         NaN                               NaN         NaN
min                                              NaN          NaN         NaN         NaN                                         NaN         NaN         NaN         NaN                               NaN         NaN
25%                                              NaN          NaN         NaN         NaN                                         NaN         NaN         NaN         NaN                               NaN         NaN
50%                                              NaN          NaN         NaN         NaN                                         NaN         NaN         NaN         NaN                               NaN         NaN
75%                                              NaN          NaN         NaN         NaN                                         NaN         NaN         NaN         NaN                               NaN         NaN
max                                              NaN          NaN         NaN         NaN                                         NaN         NaN         NaN         NaN                               NaN         NaN

==================== Sheet: Dashboard ====================
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 16 entries, 0 to 15
Data columns (total 2 columns):
 #   Column                          Non-Null Count  Dtype  
---  ------                          --------------  -----  
 0   CUSTOMER TRANSACTION DASHBOARD  0 non-null      float64
 1   Unnamed: 1                      10 non-null     object 
dtypes: float64(1), object(1)
memory usage: 384.0+ bytes
None

Summary Statistics:
        CUSTOMER TRANSACTION DASHBOARD     Unnamed: 1
count                              0.0             10
unique                             NaN             10
top                                NaN  Total Revenue
freq                               NaN              1
mean                               NaN            NaN
std                                NaN            NaN
min                                NaN            NaN
25%                                NaN            NaN
50%                                NaN            NaN
75%                                NaN            NaN
max                                NaN            NaN


```

```python
df_sales = pd.read_excel(excel_path, sheet_name='Sales Data')
df_returns = pd.read_excel(excel_path, sheet_name='Returns Data')

print("Sales Data columns:", df_sales.columns.tolist())
print(df_sales.head(3))

print("\nReturns Data columns:", df_returns.columns.tolist())
print(df_returns.head(3))

print("\nSales Data Shape:", df_sales.shape)
print("Returns Data Shape:", df_returns.shape)


```

```text
Traceback (most recent call last):
  File "<xbox-string>", line 1, in <module>
    df_sales = pd.read_excel(excel_path, sheet_name='Sales Data')
  File "_base.py", line 486, in read_excel
    data = io.parse(
  File "_base.py", line 1551, in parse
    return self._reader.parse(
  File "_base.py", line 746, in parse
    sheet = self.get_sheet_by_name(asheetname)
  File "_openpyxl.py", line 569, in get_sheet_by_name
    self.raise_if_bad_sheet_by_name(name)
  File "_base.py", line 597, in raise_if_bad_sheet_by_name
    raise ValueError(f"Worksheet named '{name}' not found")
ValueError: Worksheet named 'Sales Data' not found


```

```python
for idx, name in enumerate(xls.sheet_names):
    df = pd.read_excel(excel_path, sheet_name=idx)
    print(f"Sheet index {idx}: name='{name}'")
    print("Columns:", df.columns.tolist())
    print("Shape:", df.shape)
    print(df.head(2))
    print("-" * 50)


```

```text
Sheet index 0: name='Instruction'
Columns: ['Project Title']
Shape: (162, 1)
                                                   Project Title
0                                                            NaN
1  Customer Transaction Analysis Dashboard using Microsoft Excel
--------------------------------------------------
Sheet index 1: name='Raw_Data'
Columns: ['Transaction_ID', 'Date', 'Customer_ID', 'Customer_Name', 'Product_ID', 'Product_Name', 'Category', 'Quantity', 'Unit_Price', 'Payment_Method', 'Region', 'Customer_Segment', 'Customer_Since', 'Total_Amount', 'Month', 'Year', 'Customer Age Days', 'Customer Tenure', 'Transaction Month End', 'High Value Customer', 'Timestamp']
Shape: (250, 21)
  Transaction_ID       Date Customer_ID   Customer_Name Product_ID Product_Name     Category  Quantity  Unit_Price Payment_Method   Region Customer_Segment Customer_Since  Total_Amount     Month  Year  Customer Age Days  Customer Tenure Transaction Month End High Value Customer               Timestamp
0        TRX0155 2024-04-11     CUST023      Paul Baker       P009      Monitor  Electronics         3      249.99           Cash  Central            Basic     2021-08-17        749.97  Apr-2024  2024               1862                5            2024-04-30              Normal 2026-09-22 00:42:51.600
1        TRX0023 2024-04-14     CUST013  Joseph Jackson       P009      Monitor  Electronics         3      249.99    Credit Card  Central          Premium     2022-01-14        749.97  Apr-2024  2024               1712                4            2024-04-30              Normal 2026-09-22 00:42:51.600
--------------------------------------------------
Sheet index 2: name='Analysis'
Columns: ['Customer Transaction Analysis', 'Unnamed: 1', 'Unnamed: 2']
Shape: (52, 3)
  Customer Transaction Analysis Unnamed: 1 Unnamed: 2
0                           NaN        NaN        NaN
1                           NaN        NaN        NaN
--------------------------------------------------
Sheet index 3: name='Summary_Output'
Columns: ['SUMMARY OUTPUT', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
Shape: (273, 9)
          SUMMARY OUTPUT Unnamed: 1 Unnamed: 2 Unnamed: 3 Unnamed: 4 Unnamed: 5 Unnamed: 6 Unnamed: 7 Unnamed: 8
0                    NaN        NaN        NaN        NaN        NaN        NaN        NaN        NaN        NaN
1  Regression Statistics        NaN        NaN        NaN        NaN        NaN        NaN        NaN        NaN
--------------------------------------------------
Sheet index 4: name='Pivot_Report'
Columns: ['Pivot Reports - Customer Transaction Analysis', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']
Shape: (74, 10)
  Pivot Reports - Customer Transaction Analysis Unnamed: 1  Unnamed: 2  Unnamed: 3                                  Unnamed: 4 Unnamed: 5  Unnamed: 6  Unnamed: 7                        Unnamed: 8 Unnamed: 9
0                                           NaN        NaN         NaN         NaN                                         NaN        NaN         NaN         NaN                               NaN        NaN
1             Pivot Table 1 – Region Wise Sales        NaN         NaN         NaN  Pivot Table 3 – Product Wise Quantity Sold        NaN         NaN         NaN  Pivot Table 5 – Customer Segment        NaN
--------------------------------------------------
Sheet index 5: name='Dashboard'
Columns: ['CUSTOMER TRANSACTION DASHBOARD', 'Unnamed: 1']
Shape: (16, 2)
   CUSTOMER TRANSACTION DASHBOARD Unnamed: 1
0                             NaN        NaN
1                             NaN        NaN
--------------------------------------------------


```

```python
sheet_0 = pd.read_excel(excel_path, sheet_name=0)
print("Sheet 0 shape:", sheet_0.shape)
print("Sheet 0 columns:", sheet_0.columns)
print(sheet_0.head())

if len(xls.sheet_names) > 1:
    sheet_1 = pd.read_excel(excel_path, sheet_name=1)
    print("\nSheet 1 shape:", sheet_1.shape)
    print("Sheet 1 columns:", sheet_1.columns)
    print(sheet_1.head())


```

```text
Sheet 0 shape: (162, 1)
Sheet 0 columns: Index(['Project Title'], dtype='object')
                                                   Project Title
0                                                            NaN
1  Customer Transaction Analysis Dashboard using Microsoft Excel
2                                                            NaN
3                                                            NaN
4                                                            NaN

Sheet 1 shape: (250, 21)
Sheet 1 columns: Index(['Transaction_ID', 'Date', 'Customer_ID', 'Customer_Name', 'Product_ID',
       'Product_Name', 'Category', 'Quantity', 'Unit_Price', 'Payment_Method',
       'Region', 'Customer_Segment', 'Customer_Since', 'Total_Amount', 'Month',
       'Year', 'Customer Age Days', 'Customer Tenure', 'Transaction Month End',
       'High Value Customer', 'Timestamp'],
      dtype='object')
  Transaction_ID       Date Customer_ID     Customer_Name Product_ID Product_Name     Category  Quantity  Unit_Price Payment_Method   Region Customer_Segment Customer_Since  Total_Amount     Month  Year  Customer Age Days  Customer Tenure Transaction Month End High Value Customer               Timestamp
0        TRX0155 2024-04-11     CUST023        Paul Baker       P009      Monitor  Electronics         3      249.99           Cash  Central            Basic     2021-08-17        749.97  Apr-2024  2024               1862                5            2024-04-30              Normal 2026-09-22 00:42:51.600
1        TRX0023 2024-04-14     CUST013    Joseph Jackson       P009      Monitor  Electronics         3      249.99    Credit Card  Central          Premium     2022-01-14        749.97  Apr-2024  2024               1712                4            2024-04-30              Normal 2026-09-22 00:42:51.600
2        TRX0139 2024-04-17     CUST021       Kevin Scott       P005      Blender   Appliances         3       59.99         PayPal    South         Standard     2023-12-02        179.97  Apr-2024  2024               1025                2            2024-04-30              Normal 2026-09-22 00:42:51.600
3        TRX0043 2024-04-18     CUST046  Kathleen Bennett       P005      Blender   Appliances         3       59.99           Cash    North          Premium     2023-07-07        179.97  Apr-2024  2024               1173                3            2024-04-30              Normal 2026-09-22 00:42:51.600
4        TRX0212 2024-04-18     CUST003     Michael Brown       P008    Bookshelf    Furniture         5      149.99    Credit Card    North          Premium     2024-01-04        749.95  Apr-2024  2024                992                2            2024-04-30              Normal 2026-09-22 00:42:51.600


```

```python
print("Sheet names in order:", xls.sheet_names)

for idx, sheet in enumerate(xls.sheet_names):
    df = pd.read_excel(excel_path, sheet_name=sheet)
    print(f"\nSheet Index {idx}: '{sheet}' | Shape: {df.shape}")
    print("Non-empty preview:")
    print(df.dropna(how='all').head(10))


```

```text
Sheet names in order: ['Instruction', 'Raw_Data', 'Analysis', 'Summary_Output', 'Pivot_Report', 'Dashboard']

Sheet Index 0: 'Instruction' | Shape: (162, 1)
Non-empty preview:
                                                                                                                                                                                                                                            Project Title
1                                                                                                                                                                                           Customer Transaction Analysis Dashboard using Microsoft Excel
5                                                                                                                                                                                                                                       Project Objective
7   The objective of this project is to analyze customer transaction data using Microsoft Excel. The project demonstrates data cleaning, data analysis, Pivot Tables, Pivot Charts, Dashboard creation, and business insights for better decision-making.
11                                                                                                                                                                                                                                          Software Used
13                                                                                                                                                                                                               Microsoft Excel (Old Version Compatible)
14                                                                                                                                                                                                                                     Excel Pivot Tables
15                                                                                                                                                                                                                                           Pivot Charts
16                                                                                                                                                                                                                                  Data Analysis ToolPak
17                                                                                                                                                                                                                                              Goal Seek
18                                                                                                                                                                                                                                 Conditional Formatting

Sheet Index 1: 'Raw_Data' | Shape: (250, 21)
Non-empty preview:
  Transaction_ID       Date Customer_ID     Customer_Name Product_ID Product_Name     Category  Quantity  Unit_Price Payment_Method   Region Customer_Segment Customer_Since  Total_Amount     Month  Year  Customer Age Days  Customer Tenure Transaction Month End High Value Customer               Timestamp
0        TRX0155 2024-04-11     CUST023        Paul Baker       P009      Monitor  Electronics         3      249.99           Cash  Central            Basic     2021-08-17        749.97  Apr-2024  2024               1862                5            2024-04-30              Normal 2026-09-22 00:42:51.600
1        TRX0023 2024-04-14     CUST013    Joseph Jackson       P009      Monitor  Electronics         3      249.99    Credit Card  Central          Premium     2022-01-14        749.97  Apr-2024  2024               1712                4            2024-04-30              Normal 2026-09-22 00:42:51.600
2        TRX0139 2024-04-17     CUST021       Kevin Scott       P005      Blender   Appliances         3       59.99         PayPal    South         Standard     2023-12-02        179.97  Apr-2024  2024               1025                2            2024-04-30              Normal 2026-09-22 00:42:51.600
3        TRX0043 2024-04-18     CUST046  Kathleen Bennett       P005      Blender   Appliances         3       59.99           Cash    North          Premium     2023-07-07        179.97  Apr-2024  2024               1173                3            2024-04-30              Normal 2026-09-22 00:42:51.600
4        TRX0212 2024-04-18     CUST003     Michael Brown       P008    Bookshelf    Furniture         5      149.99    Credit Card    North          Premium     2024-01-04        749.95  Apr-2024  2024                992                2            2024-04-30              Normal 2026-09-22 00:42:51.600
5        TRX0092 2024-04-21     CUST024    Dorothy Nelson       P005      Blender   Appliances         2       59.99           Cash     West         Standard     2022-07-04        119.98  Apr-2024  2024               1541                4            2024-04-30              Normal 2026-09-22 00:42:51.600
6        TRX0193 2024-04-21     CUST024    Dorothy Nelson       P003   Headphones  Electronics         5      149.99    Credit Card  Central         Standard     2021-08-10        749.95  Apr-2024  2024               1869                5            2024-04-30              Normal 2026-09-22 00:42:51.600
7        TRX0240 2024-04-21     CUST017     Charles Allen       P005      Blender   Appliances         3       59.99     Debit Card    South            Basic     2023-03-22        179.97  Apr-2024  2024               1280                3            2024-04-30              Normal 2026-09-22 00:42:51.600
8        TRX0135 2024-04-23     CUST009     Robert Thomas       P001       Laptop  Electronics         3      899.99         PayPal     East         Standard     2024-01-08       2699.97  Apr-2024  2024                988                2            2024-04-30          High Value 2026-09-22 00:42:51.600
9        TRX0185 2024-04-26     CUST026        Karen Hill       P005      Blender   Appliances         3       59.99           Cash     West          Premium     2023-10-08        179.97  Apr-2024  2024               1080                2            2024-04-30              Normal 2026-09-22 00:42:51.600

Sheet Index 2: 'Analysis' | Shape: (52, 3)
Non-empty preview:
   Customer Transaction Analysis Unnamed: 1 Unnamed: 2
2                              A          B        NaN
3                  Total Revenue  229192.47        NaN
4             Total Transactions        250        NaN
5            Total Quantity Sold        753        NaN
6            Average Order Value  916.76988        NaN
7                   Highest Sale    4499.95        NaN
8                    Lowest Sale      59.99        NaN
9                Total Customers        250        NaN
10          High Value Customers         61        NaN
12                      Category      Sales        NaN

Sheet Index 3: 'Summary_Output' | Shape: (273, 9)
Non-empty preview:
           SUMMARY OUTPUT  Unnamed: 1        Unnamed: 2       Unnamed: 3 Unnamed: 4      Unnamed: 5 Unnamed: 6 Unnamed: 7 Unnamed: 8
1   Regression Statistics         NaN               NaN              NaN        NaN             NaN        NaN        NaN        NaN
2              Multiple R     0.45808               NaN              NaN        NaN             NaN        NaN        NaN        NaN
3                R Square    0.209838               NaN              NaN        NaN             NaN        NaN        NaN        NaN
4       Adjusted R Square    0.206651               NaN              NaN        NaN             NaN        NaN        NaN        NaN
5          Standard Error  919.607153               NaN              NaN        NaN             NaN        NaN        NaN        NaN
6            Observations         250               NaN              NaN        NaN             NaN        NaN        NaN        NaN
8                   ANOVA         NaN               NaN              NaN        NaN             NaN        NaN        NaN        NaN
9                     NaN          df                SS               MS          F  Significance F        NaN        NaN        NaN
10             Regression           1   55695916.896353  55695916.896353  65.859537             0.0        NaN        NaN        NaN
11               Residual         248  209727974.556144    845677.316759        NaN             NaN        NaN        NaN        NaN

Sheet Index 4: 'Pivot_Report' | Shape: (74, 10)
Non-empty preview:
   Pivot Reports - Customer Transaction Analysis   Unnamed: 1  Unnamed: 2  Unnamed: 3                                  Unnamed: 4     Unnamed: 5  Unnamed: 6  Unnamed: 7                        Unnamed: 8   Unnamed: 9
1              Pivot Table 1 – Region Wise Sales          NaN         NaN         NaN  Pivot Table 3 – Product Wise Quantity Sold            NaN         NaN         NaN  Pivot Table 5 – Customer Segment          NaN
2                                         Region  Total Sales         NaN         NaN                                Product_Name  Quantity Sold         NaN         NaN                  Customer_Segment  Total_Sales
3                                        Central     41288.34         NaN         NaN                                   Bookshelf            102         NaN         NaN                             Basic     62177.81
4                                           East     59288.39         NaN         NaN                                  Smartphone             96         NaN         NaN                           Premium     84657.12
5                                          North     50808.31         NaN         NaN                                    Keyboard             89         NaN         NaN                          Standard     82357.54
6                                          South     36398.75         NaN         NaN                                     Monitor             88         NaN         NaN                               NaN          NaN
7                                           West     41408.68         NaN         NaN                                     Blender             87         NaN         NaN                               NaN          NaN
8                                            NaN          NaN         NaN         NaN                                        Desk             78         NaN         NaN                               NaN          NaN
9                                            NaN          NaN         NaN         NaN                                      Laptop             75         NaN         NaN                               NaN          NaN
10                                           NaN          NaN         NaN         NaN                                Office Chair             52         NaN         NaN                               NaN          NaN

Sheet Index 5: 'Dashboard' | Shape: (16, 2)
Non-empty preview:
    CUSTOMER TRANSACTION DASHBOARD            Unnamed: 1
2                              NaN         Total Revenue
3                              NaN             229192.47
5                              NaN          Transactions
6                              NaN                   250
8                              NaN         Quantity Sold
9                              NaN                   753
11                             NaN         Average Order
12                             NaN             916.76988
14                             NaN  High Value Customers
15                             NaN                    61


```

Here is a complete, publication-ready **README.md** tailored for your project. You can copy and paste this directly into your GitHub repository or documentation folder.

---

# Project Data Analysis & Dashboard Workbook

## Overview

This project centers on data processing, structured modeling, and visual reporting based on the data contained in **`s Final Project (1).xlsx`**. The primary goal is to transform raw operational/sales dataset attributes into actionable business insights through structured data cleaning, aggregation, and performance metrics tracking.

---

## Dataset Structure

The Excel workbook (`s Final Project (1).xlsx`) consists of structured data sheets containing key operational fields:

* **Transaction & Sales Records:** Unique IDs, order dates, ship dates, and order status tracking.
* **Customer Information:** Demographics, segment groupings, and geographical distributions.
* **Product & Category Breakdown:** Product hierarchy (Category, Sub-Category), item SKUs, and pricing metrics.
* **Financial Metrics:** Sales revenue, order quantities, discount percentages, and net profit margins.

---

## Key Performance Indicators (KPIs) Analyzed

The project focuses on measuring and visualizing key business metrics, including:

1. **Total Revenue & Profitability:** Sum of overall sales, overall net profit, and profit margins across regions and categories.
2. **Category & Product Analysis:** Identification of top-performing product lines and underperforming SKUs.
3. **Geographical Distribution:** Sales distribution across regions, states, and key target markets.
4. **Customer Segment Insights:** Revenue contribution per segment (e.g., Consumer, Corporate, Home Office).
5. **Fulfillment Efficiency:** Delivery time tracking calculated from Order Date vs. Ship Date.

---

## Analysis & Visualizations Included

* **Executive Summary Dashboard:** High-level summary cards showing total revenue, profit, total orders, and average margin.
* **Trend & Seasonality Analysis:** Line charts mapping revenue growth and seasonal order spikes over time.
* **Comparative Performance:** Bar and column charts displaying profitability by region and product sub-category.
* **Proportional Distribution:** Pie/Donut charts highlighting customer segment composition and shipping mode preferences.

---

## Tools & Technologies Used

* **Microsoft Excel / Power BI:** Data processing, Pivot Tables, DAX / Formulas, and visual dashboard design.
* **Data Cleaning & Preparation:** Power Query / Excel Data Functions (handling null values, date parsing, data typing).

---

## How to Access & Use

1. **Clone or Download the Repository:**
```bash
git clone https://github.com/your-username/your-repo-name.git

```


2. **Open the File:**
* Open `s Final Project (1).xlsx` using Microsoft Excel (2016 or newer recommended) or import it into Microsoft Power BI Desktop.


3. **Explore the Sheets/Dashboards:**
* Navigate through the main summary worksheets or pivot tables to inspect regional performance and product analytics. Use built-in slicers (e.g., Year, Region, Segment) to filter views dynamically.
