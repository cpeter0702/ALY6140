import pandas as pd
import time
import matplotlib.pyplot as plt


def read_excel(file_path):
    df = pd.read_excel(file_path, 0)
    return df


def clean_df (df):
    manhattan_df.drop(index=[0, 1, 2, 3], axis=1, inplace=True)


def rename_columns(df):
    col_names = ['BOROUGH', 'NEIGHBORHOOD', 'BUILDING_CLASS_CATEGORY', 'TAX_CLASS_AT_PRESENT',
                 'BLOCK', 'LOT', 'EASE_MENT', 'BUILDING_CLASS_AT_PRESENT', 'ADDRESS',
                 'APARTMENT_NUMBER', 'ZIPCODE', 'RESIDENTIAL_UNITS', 'COMMERCIAL_UNITS',
                 'TOTAL_UNITS', 'LAND_SQUARE_FEET', 'GROSS_SQUARE_FEET', 'YEAR_BUILT',
                 'TAX_CLASS_AT_TIME_OF_SALE', 'BUILDING_CLASS_AT_TIME_OF_SALE',
                 'SALE_PRICE', 'SALE_DATE']
    df.columns = col_names


def sortout_df(df):
    df.drop(df[df['SALE_PRICE'] == 0].index, inplace=True)
    df.drop(df[df['YEAR_BUILT'] == 0].index, inplace=True)
    year_month = df['SALE_DATE'].dt.strftime('%Y%m')
    df.insert(21, 'yearMonth', year_month)
    df.insert(22, 'yearOld', (this_year - df['YEAR_BUILT']))
    df.reset_index()


def create_estate_trend(df):
    month_sale_data = df.groupby(["yearMonth"], as_index=False)["SALE_PRICE"] \
        .agg({"sum": "sum", "count": "count"})

    month_sale_data.insert(3, 'mean', round(month_sale_data['sum'] / month_sale_data['count'], 2))
    plt.bar(month_sale_data['yearMonth'], month_sale_data['mean'])
    plt.title('The Trend of Average Sales in 2017 to 2018 at Manhattan')
    plt.xlabel('Average Sales Price')
    plt.ylabel('Month')
    plt.show()


def compare_building_age(df):
    year_old_data = df[['yearMonth', 'YEAR_BUILT', 'SALE_PRICE', 'yearOld']]
    year_old_data.insert(4, 'oldType', year_old_data['yearOld'].map(lambda x: 'old' if x <= 30 else 'young'))
    group_data = year_old_data.groupby(["yearMonth", "oldType"], as_index=False)["SALE_PRICE"] \
        .agg({"sum": "sum", "count": "count"})
    group_data.insert(4, 'mean', round(group_data['sum'] / group_data['count'], 2))

    old_data = group_data[group_data['oldType'] == 'old']['mean']
    young_data = group_data[group_data['oldType'] == 'young']['mean']
    old_month = group_data[group_data['oldType'] == 'old']['yearMonth']
    young_month = group_data[group_data['oldType'] == 'young']['yearMonth']

    plt.clf()
    plt.plot(old_month, old_data)
    plt.plot(young_month, young_data)
    plt.title('AAA')
    plt.xlabel('Average Sales Price')
    plt.ylabel('Month')
    plt.legend(['old apt', 'young apt'], loc='upper left')
    plt.show()


# setting
local_path = "/Users/peterliao/Desktop/NYC Real Estate/"
file_name = "rollingsales_manhattan.xls"
this_year = int(time.strftime("%Y", time.localtime()))

# main logic
manhattan_df = read_excel(local_path + file_name)       # read data set
clean_df(manhattan_df)                                  # cleaning the dataset
rename_columns(manhattan_df)                            # columns naming
sortout_df(manhattan_df)                                # removing/filling missing data
create_estate_trend(manhattan_df)                       # 1. Sale Price vs Month
compare_building_age(manhattan_df)                      # 2. old building vs new building

# Analysis:
# The bar plot shows that the whole apartment prices are rising from November 2017 to October 2018
# I divided the data set into two group.
# Younger houses less than 30 years old, and more than 30 are old building
# We can find out that they had a similar trend
# but The price response of the old apt will be earlier than the young apt.

df = pd.DataFrame({'link': ['=HYPERLINK("https://www1.nyc.gov/assets/finance/downloads/pdf/rolling_sales/rollingsales_manhattan.xls", "some website")']})
df.to_excel('test.xlsx')

