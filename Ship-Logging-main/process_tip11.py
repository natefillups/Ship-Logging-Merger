import pandas as pd
def process_tip(tippath, mrldf):
    print("Opening tip file")
    tipdf = pd.read_excel(tippath)
    print("Opening mrl file")
    # mrldf = pd.read_excel(mrlpath)
    tipdf = tipdf.drop(index=0)
    # Only reliable row counter is the length of the spreadsheet
    # Take range of length so I can index row later
    mrllength = len(mrldf.index)

    mrltiplocations = {}

    # mrltiplocations is a dictionary with keys as indexes and values as corresponding QA/WAF # values
    print("Finding TIP matches")
    for i in range(mrllength):
        if mrldf.loc[i, 'MATCH'] == "TIP":
            mrltiplocations[i] = mrldf.loc[i, 'QA/WAF #']
    
    for n, val in enumerate(tipdf['* BAE Insp ID']):
        if str(val) == "Nan" or val == None or str(val) == "nan":
            continue
        match = False
        for j in mrltiplocations:
            if mrltiplocations[j] == val:
                # Operation that prevents nans for being printed
                if str(tipdf.loc[n + 1, '* Work Spec No']) == 'nan' or str(tipdf.loc[n + 1, '* Work Spec No']) == 'NaN' or tipdf.loc[n + 1, '* Work Spec No'] == None:
                    wsval = None
                else:
                    wsval = str(tipdf.loc[n + 1, '* Work Spec No']) + '.2'
                match = True
                # Replace each value with given 'tip11' value
                mrldf.loc[j, 'Work Item'] = wsval
                mrldf.loc[j, 'Dept/Shop #'] = tipdf.loc[n + 1, 'BAE Shop']
                mrldf.loc[j, 'Key Trade (BAE)'] = tipdf.loc[n + 1, 'BAE Shop Name']
                mrldf.loc[j, 'Dept/Shop #'] = tipdf.loc[n + 1, 'Subcontractor ID']
                mrldf.loc[j, 'Key Trade (SUB)'] = tipdf.loc[n + 1, 'Subcontractor Name']
                mrldf.loc[j, 'Location'] = tipdf.loc[n + 1, '* Inspection Location']
                mrldf.loc[j, 'QA/WAF #'] = tipdf.loc[n + 1, '* BAE Insp ID']
                mrldf.loc[j, 'CP'] = tipdf.loc[n + 1, 'Check Point Type']
                mrldf.loc[j, 'Accepted'] = tipdf.loc[n + 1, 'Accept Date']
                mrldf.loc[j, 'Rejected'] = tipdf.loc[n + 1, 'Reject Date']
                mrldf.loc[j, 'Status'] = tipdf.loc[n + 1, 'TIP Status']
                mrldf.loc[j, 'MS'] = tipdf.loc[n + 1, '* Key Event ID']
                mrldf.loc[j, 'Component'] = tipdf.loc[n + 1, '* Inspected Component']
                mrldf.loc[j, 'T&I Comments'] = tipdf.loc[n + 1, '* Pass/Fail Criteria']
                mrldf.loc[j, 'Spec'] = tipdf.loc[n + 1, 'Spec Para  No']
                mrldf.loc[j, 'Test Description'] = tipdf.loc[n + 1, '* Para Description']
                mrldf.loc[j, 'NSI'] = tipdf.loc[n + 1, 'Std Item No']
                mrldf.loc[j, 'Para'] = tipdf.loc[n + 1, 'Std Item Para No']
                mrldf.loc[j, 'Criteria'] = tipdf.loc[n + 1, 'Inspection to  be Performed']
        if match == False:
            # If there is no match then add the values to the end of the dataframe
            if str(tipdf.loc[n, '* Work Spec No']) == 'nan':
                wsval = None
            else:
                wsval = str(tipdf.loc[n, '* Work Spec No']) + '.2'
            mrldf = mrldf.append({
                'Work Item': wsval, # MRL COL C
                'Dept/Shop #': tipdf.loc[n, 'BAE Shop'], # MRL COL H
                'TITLE': 'TIP',
                'MATCH': 'TIP',
                'Key Trade (BAE)': tipdf.loc[n, 'BAE Shop Name'], # MRL COL I
                'Dept/Shop #': tipdf.loc[n, 'Subcontractor ID'], # MRL COL J
                'Key Trade (SUB)': tipdf.loc[n, 'Subcontractor Name'], # MRL COL K
                'Location': tipdf.loc[n, '* Inspection Location'], # MRL COL L
                'QA/WAF #': tipdf.loc[n, '* BAE Insp ID'], # MRL COL M
                'CP': tipdf.loc[n, 'Check Point Type'], # MRL COL N
                'Accepted': tipdf.loc[n, 'Accept Date'], # MRL COL O
                'Rejected': tipdf.loc[n, 'Reject Date'], # MRL COL P
                'Status': tipdf.loc[n, 'TIP Status'], # MRL COL AP
                'MS': tipdf.loc[n, '* Key Event ID'], # MRL COL AU
                'Component': tipdf.loc[n, '* Inspected Component'], # MRL COL AV
                'T&I Comments': tipdf.loc[n, '* Pass/Fail Criteria'], # MRL COL AW
                'Spec': tipdf.loc[n, 'Spec Para  No'], # MRL COL AX
                'Test Description': tipdf.loc[n, '* Para Description'], # MRL COL AY
                'NSI': tipdf.loc[n, 'Std Item No'], # MRL COL AZ
                'Para': tipdf.loc[n, 'Std Item Para No'], # MRL COL BA
                'Criteria': tipdf.loc[n, 'Inspection to  be Performed'], # MRL COL BB
                }, ignore_index=True)
        print(n, "out of", len(tipdf.index), "complete.")


    print("Writing to new file")

    return mrldf
