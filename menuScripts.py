import numpy as np
import xlwings as xw
from datetime import datetime
import requests
import json
import pandas as pd

def findProduct(product):

    #===STEP 1: clean variables and initiate workbook===
    wb = xw.Book.caller()
    x_api_id = wb.sheets['Menu'].range("C2").value
    x_api_key = wb.sheets['Menu'].range("C3").value
    searchType = wb.sheets['Menu'].range("C4").value
    rtnCount = int(wb.sheets['Menu'].range("C5").value)

    #===STEP 2: initiate API search for related products===
    url = 'https://trackapi.nutritionix.com/v2/search/instant'
    headers = {
            'x-app-id':x_api_id,
            'x-app-key':x_api_key,
            'x-remote-user-id':'0'}
    data = {
        'query':product}

    response = requests.post(url, headers=headers, data=data)
    dct_search = json.loads(response.text)

    #===STEP 3: create requested products list of length return count===

    #----STEP 3.1: loop through product instances and request their nutritional information depending on the search type----
    productList = []
    df_nutrients = wb.sheets['Nutrients'].range('A1:C8').options(pd.DataFrame).value
    for product in dct_search[searchType]:
        if searchType == 'common':
            food_name = product['food_name']
            base_url = "https://trackapi.nutritionix.com/v2/natural/nutrients"
            query = {
                "query":food_name
            }
            response = requests.request("POST", base_url, headers=headers, data=query)

        else:
            item_id = product['nix_item_id']
            base_url = 'https://trackapi.nutritionix.com/v2/search/item?nix_item_id=' + item_id
            response = requests.get(base_url, headers=headers)

        #----STEP 3.2: process the request and extract info from json/dict object----
        dct_item = json.loads(response.text)
        food_name = dct_item['foods'][0]['food_name']
        brand_name = dct_item['foods'][0]['brand_name'] if dct_item['foods'][0]['brand_name'] is not None else "NA"
        serving_weight_grams = dct_item['foods'][0]['serving_weight_grams']
        df_item = pd.DataFrame(dct_item['foods'][0]['full_nutrients']).set_index('attr_id')

        #----STEP 3.3: merge the requested item nutrients list back to template nutrients list----
        df_results = df_nutrients.merge(df_item, how='left', left_index=True, right_index=True)
        df_results['value'] = df_results['value'].fillna(0)
        
        #----STEP 3.4: create and append requested item results to product list and determine whether to break or continue----
        list_result = [food_name] + [brand_name] + [serving_weight_grams] + df_results['value'].tolist()
        productList.append(list_result)
        if len(productList) == rtnCount:
            break

    #===STEP 4: transform product list and export back to excel sheet===
    cols = ['Return Name', 'Brand Name', 'Serving Size', 'Calories', 'Protein', 'Carbs', 'Fiber', 'Added Sugars', 'Total Fat', 'Saturated Fat']
    df_to_sheet = pd.DataFrame(productList, columns=cols)
    wb.sheets['Menu'].range('F2').options(index=False, header=False).value = df_to_sheet

def addProduct(items):

    wb = xw.Book.caller()

    #===STEP 1: process input parameter===
    item_list = items.split()
    items = list(map(int, items)) # list mapping
    items = [x - 1 for x in items] # list comprehension

    #===STEP 2: record dataframes from Foods and Menu sheet
    df_results = wb.sheets['Menu'].range('F1:O11').options(pd.DataFrame).value
    df_results.reset_index(inplace=True)

    lastRow = wb.sheets['Foods'].range('A' + str(wb.sheets['Foods'].cells.last_cell.row)).end('up').row
    strRng = 'A1:K' + str(lastRow)
    df_foods = wb.sheets['Foods'].range(strRng).options(pd.DataFrame).value
    df_foods.index = list(map(int, list(df_foods.index)))

    #===STEP 3: for each checked food, add it to the foods dataframe===
    for i in items:
        df_foods.loc[len(df_foods.index)] = list(df_results.iloc[i])

    #===STEP 4: output the newly updated foods dataframe back to Foods sheet===
    wb.sheets['Foods'].range('A2').options(index=True, header=False).value = df_foods



# === CODE BELOW FOR DEBUGGING PURPOSES ===

# if __name__ == '__main__':
#     items = "1"
#     # Expects the Excel file next to this source file, adjust accordingly.
#     xw.Book('MACROS.xlsm').set_mock_caller()
#     addProduct(items)
