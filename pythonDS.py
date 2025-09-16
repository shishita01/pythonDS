import pandas as pd
import zipfile
import docx


zip_path = r"DSassignment3/Python for Data Science - Module 2 (1).zip"

with zipfile.ZipFile(zip_path, 'r') as z:
    print("Files in the zip archive:", z.namelist())
    
    docx_file = 'Python for Data Science - Assignments/Problem Statement/DS Internship - PDS - Problem Statement.docx'
    
    with z.open(docx_file) as f:
        doc = docx.Document(f)
        
        for para in doc.paragraphs:
            print(para.text)



zip_path = r"DSassignment3/Python for Data Science - Module 2 (1).zip"

with zipfile.ZipFile(zip_path, 'r') as z:
    z.extract('Python for Data Science - Assignments/Data/SalesData.xlsx', path='.')

file_path = 'Python for Data Science - Assignments/Data/SalesData.xlsx'


sales = pd.read_excel(file_path)

print(sales.head(10))



# Group by 'Item' and find the minimum 'Sale_amt' for each item
result = sales.groupby('Item')['Sale_amt'].min().reset_index()

print(result)




sales['OrderDate'] = pd.to_datetime(sales['OrderDate'])

sales['Year'] = sales['OrderDate'].dt.year

result = sales.groupby(['Year', 'Region'])['Sale_amt'].sum().reset_index()

print(result)




sales['OrderDate'] = pd.to_datetime(sales['OrderDate'])

reference_date = pd.to_datetime('2023-12-31')

sales['days_diff'] = (reference_date - sales['OrderDate']).dt.days

print(sales.head(10))





result_q4 = sales.groupby('Manager')['SalesMan'].unique().reset_index()
result_q4.rename(columns={'SalesMan': 'list_of_salesmen'}, inplace=True)

print("Question 4 Result:")
print(result_q4)


result_q5 = sales.groupby('Region').agg(
    salesmen_count = ('SalesMan', 'nunique'),
    total_sales = ('Sale_amt', 'sum')
).reset_index()


print(result_q5)




total_sales = sales['Sale_amt'].sum()

manager_sales = sales.groupby('Manager')['Sale_amt'].sum().reset_index()

manager_sales['percent_sales'] = (manager_sales['Sale_amt'] / total_sales) * 100

result_q6 = manager_sales[['Manager', 'percent_sales']]

print("Question 6 Result:")
print(result_q6)








zip_path = r"DSassignment3/Python for Data Science - Module 2 (1).zip"

with zipfile.ZipFile(zip_path, 'r') as z:
    z.extract('Python for Data Science - Assignments/Data/imdb.csv', path='.')

file_path = 'Python for Data Science - Assignments/Data/imdb.csv'
imdb = pd.read_csv(file_path, on_bad_lines='skip')

print(imdb.head())



imdb.columns = imdb.columns.str.strip()

print("Columns in imdb dataset:")
print(imdb.columns)


result_q7 = imdb.loc[4, 'imdbRating']

print("\nQuestion 7 Result:")
print("IMDb rating of the fifth movie:", result_q7)






# Find shortest and longest durations
shortest_duration = imdb['duration'].min()
longest_duration = imdb['duration'].max()


shortest_title = imdb[imdb['duration'] == shortest_duration]['title'].values[0]
longest_title = imdb[imdb['duration'] == longest_duration]['title'].values[0]

print("\nQuestion 8 Result:")
print("Shortest movie title:", shortest_title)
print("Longest movie title:", longest_title)






# Sort by 'year' (earliest first) and 'imdbRating' (highest first)
result_q9 = imdb.sort_values(by=['year', 'imdbRating'], ascending=[True, False])

print("\nQuestion 9 Result:")
print(result_q9[['title', 'year', 'imdbRating']].head(10))






# Filter movies with duration between 30 and 180 minutes
result_q10 = imdb[(imdb['duration'] >= 30) & (imdb['duration'] <= 180)]

print("\nQuestion 10 Result:")
print(result_q10[['title', 'duration']].head(10))









zip_path = r"DSassignment3/Python for Data Science - Module 2 (1).zip"

# Extract the diamonds.csv file
with zipfile.ZipFile(zip_path, 'r') as z:
    z.extract('Python for Data Science - Assignments/Data/diamonds.csv', path='.')

file_path = 'Python for Data Science - Assignments/Data/diamonds.csv'
diamonds = pd.read_csv(file_path, on_bad_lines='skip')

diamonds.columns = diamonds.columns.str.strip()


print("First 5 rows of diamonds dataset:")
print(diamonds.head())

# Print column names
print("\nColumns in diamonds dataset:")
print(diamonds.columns)






# Count the duplicate rows of diamonds DataFrame
duplicate_count = diamonds.duplicated().sum()
print("Question 11 Result:")
print("Number of duplicate rows:", duplicate_count)




# Drop rows in case of missing values in 'carat' and 'cut' columns
diamonds_cleaned = diamonds.dropna(subset=['carat', 'cut'])
print("\nQuestion 12 Result:")
print("Rows after dropping missing 'carat' and 'cut':", len(diamonds_cleaned))



# Subset the dataframe with only numeric columns
diamonds_numeric = diamonds.select_dtypes(include='number')
print("\nQuestion 13 Result:")
print("Numeric columns in the dataset:")
print(diamonds_numeric.head())



# Create a new column 'volume' based on the condition
diamonds['volume'] = diamonds.apply(
    lambda row: row['x'] * row['y'] * row['z'] if row['depth'] > 60 else 8,
    axis=1
)
print("\nQuestion 14 Result:")
print("Volume column added (showing first 5 rows):")
print(diamonds[['x', 'y', 'z', 'depth', 'volume']].head())




# Impute missing price values with the mean price
mean_price = diamonds['price'].mean()
diamonds['price'].fillna(mean_price, inplace=True)
print("\nQuestion 15 Result:")
print("Missing price values imputed with mean price:")
print(diamonds['price'].isna().sum())


