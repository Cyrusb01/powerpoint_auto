"""
This script will connect to the google sheets and return all the data needed 
"""

import json
import os 
import pandas as pd
import gspread

def connect_to_sheet():
    print("FETCHING ALL THE DATA")
    PRIVATE_KEY = json.loads(os.getenv("INVESTOR_KEY"))
    PRIVATE_KEY_ID = os.getenv("INVESTOR_KEY_ID")
    
    credentials = {
        "type": "service_account",
        "project_id": "investor-relations-336504",
        "private_key_id": PRIVATE_KEY_ID,
        "private_key": PRIVATE_KEY,
        "client_email": "gspread-investor-dashboard@investor-relations-336504.iam.gserviceaccount.com",
        "client_id": "104166031420505880990",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/gspread-investor-dashboard%40investor-relations-336504.iam.gserviceaccount.com"
    }
    
    gc = gspread.service_account_from_dict(credentials)
    
    social_media_sheet = gc.open_by_url("https://docs.google.com/spreadsheets/d/1vG7-hic7poIB7HiwiqoixnOTwQXCBa3qq_r0UV-xxyo/edit#gid=1364942338")
    since_inception = social_media_sheet.get_worksheet(1)
  
    #Line Chart Data
    df = pd.DataFrame(since_inception.get_all_records())

    df = df.set_index("Date")
    df.index = pd.to_datetime(df.index)

    line_chart  = df.iloc[: , :2]
    
    drawdown_chart  = df.iloc[: , 4:6]
    drawdown_chart['Max_Drawdown_Multi-Strat Fund'] = drawdown_chart['Max_Drawdown_Multi-Strat Fund'].str.replace('%', "").astype('float') / 100
    drawdown_chart['Max_Drawdown_Bitcoin'] = drawdown_chart['Max_Drawdown_Bitcoin'].str.replace('%', "").astype('float') / 100
    # print(df['Max_Drawdown_Multi-Strat Fund'])
    # print(drawdown_chart)

    tables = df.iloc[: , 7:]
    table_list = tables.values.tolist()
    performance_table = table_list[1:4]
    performance_table[0][0] = "Performance(%)"
    performance_table[1][0] = "BFC Multi-Strat. Fund"

    holdings = tables["Table1"].to_list()
    holdings = holdings[holdings.index('Holdings'):]
    holdings = [x for x in holdings if x != '']

    weights = tables["Table2"].to_list()
    weights = weights[weights.index('Weights'):]
    weights = [x for x in weights if x != '']

    raw_strategy = tables["Table4"].to_list()
    strategy = raw_strategy[raw_strategy.index('Strategy'): raw_strategy.index('Assets ')]
    strategy = [x for x in strategy if x != '']
    
    strategy_weights = tables["Table5"].to_list()
    strategy_weights = strategy_weights[raw_strategy.index('Strategy'): raw_strategy.index('Assets ')]
    strategy_weights = [x for x in strategy_weights if x != '']
    strategy_weights = [float(x.replace("%", ""))/100 for x in strategy_weights if x != "Weights"]

    # print(strategy)
    # print(strategy_weights)


    raw_hedge = tables["Table4"].to_list()
    hedge = raw_hedge[raw_hedge.index('Assets '):]
    hedge = [x for x in hedge if x != '']
    
    hedge_weights = tables["Table5"].to_list()
    hedge_weights = hedge_weights[raw_hedge.index('Assets '):]
    hedge_weights = [x for x in hedge_weights if x != '']
    hedge_weights = [float(x.replace("%", ""))/100 for x in hedge_weights if x != "Weights"]
    
    # print(hedge)
    # print(hedge_weights)

    print("DONE WITH DATA")
    return [line_chart, drawdown_chart, performance_table, holdings, weights, strategy, strategy_weights, hedge, hedge_weights]
