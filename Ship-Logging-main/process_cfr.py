import pandas as pd

def process_cfr(cfrpath, mrldf):
    print("Opening cfr file")
    cfrdf = pd.read_excel(cfrpath, parse_dates=False)
    print("Opening mrl file")
    # Only reliable row counter is the length of the spreadsheet
    # Take range of length so I can index row later
    mrllength = len(mrldf.index)

    mrltiplocations = {}

    # mrltiplocations is a dictionary with keys as indexes and values as corresponding QA/WAF # values
    print("Finding IMS matches")
    for i in range(mrllength):
        if mrldf.loc[i, 'MATCH'] == "IMS":
            mrltiplocations[i] = mrldf.loc[i, 'CFR #']

    for n, val in enumerate(cfrdf['Serial']):
        match = False
        for j in mrltiplocations:
            if mrltiplocations[j] == val:
                match = True
                # Replace each value with given 'tip11' value
                mrldf.loc[j, 'CFR #'] = cfrdf.loc[n, 'Serial'], # MRL COL V
                mrldf.loc[j, 'Sub \nCFR #'] = cfrdf.loc[n, 'TradeSerial'], # MRL COL W
                mrldf.loc[j, 'Work Item'] = str(cfrdf.loc[n, 'ItemNo']) + ".4", # MRL COL C
                mrldf.loc[j, 'TITLE'] = cfrdf.loc[n, 'RptType'], # MRL COL D
                mrldf.loc[j, 'Component'] = cfrdf.loc[n, 'RptDescription'], # MRL COL AV
                mrldf.loc[j, 'Spec'] = cfrdf.loc[n, 'Para'], # MRL COL AX
                mrldf.loc[j, 'Submitted Date'] = cfrdf.loc[n, 'DateSubmitted'], # MRL COL X
                # mrldf.loc[j, 'Answered Date'] = cfrdf.loc[n, 'AnswerDate'], # MRL COL X
                mrldf.loc[j, 'RCC#'] = cfrdf.loc[n, 'RCCNo'], # MRL COL Z
                mrldf.loc[j, 'Status'] = cfrdf.loc[n, 'Status'], # MRL COL AP
                mrldf.loc[j, 'Criteria'] = cfrdf.loc[n, 'Condition'], # MRL COL BB
                mrldf.loc[j, 'RA/QA Action'] = cfrdf.loc[n, 'Recommendation'], # MRL COL BE
                mrldf.loc[j, 'Integrator Comments'] = cfrdf.loc[n, 'RptAction'], # MRL COL BF
                mrldf.loc[j, 'T&I Comments'] = cfrdf.loc[n, 'Comment'], # MRL COL AW
                mrldf.loc[j, 'Duration'] = cfrdf.loc[n, 'DaysOpen'], # MRL COL AG

        if match == False:
            # If there is no match then add the values to the end of the dataframe
            if str(cfrdf.loc[n, 'ItemNo']) == 'nan' or cfrdf.loc[n, 'ItemNo'] == None or cfrdf.loc[n, 'ItemNo'] == 0 or str(cfrdf.loc[n, 'ItemNo']) == 'NaN':
                itemval = None
            else:
                itemval = str(cfrdf.loc[n, 'ItemNo']) + '.4'
            mrldf = mrldf.append({
                'MATCH': 'IMS',
                'CFR #': cfrdf.loc[n, 'Serial'], # MRL COL V
                'Sub \nCFR #': cfrdf.loc[n, 'TradeSerial'], # MRL COL W
                'Work Item': itemval, # MRL COL C
                'TITLE': cfrdf.loc[n, 'RptType'], # MRL COL D
                'Component': cfrdf.loc[n, 'RptDescription'], # MRL COL AV
                'Spec': cfrdf.loc[n, 'Para'], # MRL COL AX
                'Submitted Date': cfrdf.loc[n, 'DateSubmitted'], # MRL COL X
                'Answered Date': cfrdf.loc[n, 'AnswerDate'], # MRL COL Y
                'RCC#': cfrdf.loc[n, 'RCCNo'], # MRL COL Z
                'Status': cfrdf.loc[n, 'Status'], # MRL COL AP
                'Criteria': cfrdf.loc[n, 'Condition'], # MRL COL BB
                'RA/QA Action': cfrdf.loc[n, 'Recommendation'], # MRL COL BE
                'Integrator Comments': cfrdf.loc[n, 'RptAction'], # MRL COL BF
                'T&I Comments': cfrdf.loc[n, 'Comment'], # MRL COL AW
                'Duration': cfrdf.loc[n, 'DaysOpen'], # MRL COL AG
                }, ignore_index=True)
        print(n, "out of", len(cfrdf.index), "complete.")

    return mrldf
