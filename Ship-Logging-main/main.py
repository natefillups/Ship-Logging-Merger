import os
import pandas as pd
from process_cfr import process_cfr
from process_mrlview import process_mrlview
from process_tip11 import process_tip

files = os.listdir()

def finishdoc(tempmrl):
    print("Styling excel sheet.")
    tempmrl = trimdate(tempmrl, 'Answered Date')
    tempmrl = trimdate(tempmrl, 'Early/\nActual\nStart')
    tempmrl = trimdate(tempmrl, 'Early/\nActual\nStop')
    tempmrl = trimdate(tempmrl, 'Late Start')
    tempmrl = trimdate(tempmrl, 'Late Stop')
    tempmrl = trimdate(tempmrl, 'Submitted Date')
    tempmrl = trimdate(tempmrl, 'Issued/Date Entered')
    tempmrl = trimdate(tempmrl, 'Rejected')
    tempmrl = trimdate(tempmrl, 'Answered Date')
    tempmrl = trimdate(tempmrl, 'Accepted')
    

    tempmrl = tempmrl.style.apply(highlight_tip, axis=1)
    print("Writing to new excel file.")
    tempmrl.to_excel('Newdf.xlsx', index=False)
    return tempmrl

def highlight_tip(s):
    if s.loc['MATCH'] == 'TIP': return ['color: #C00000'] * len(s)
    if s.loc['MATCH'] == 'IMS': return ['color: #ee7600'] * len(s)
    if s.loc['MATCH'] == 'KEMS': return ['color: #000000; font-weight: bold;'] * len(s)
    if s.loc['MATCH'] == 'APPROVED': return ['font-weight: bold; color: #000000;'] * len(s)
    if s.loc['MATCH'] == 'PS': return ['color: #191919;'] * len(s)
    if s.loc['MATCH'] == 'P': return ['font-weight: bold; color: #191919;'] * len(s)
    if s.loc['MATCH'] == 'WAF': return ['color: #964B00;'] * len(s)
    if s.loc['MATCH'] == 'TIP-NIS': return ['color: #8B0000;'] * len(s)
    if s.loc['MATCH'] == 'RR': return ['color: #800080;'] * len(s)
    if s.loc['MATCH'] == 'NIS': return ['color: #000000; font-weight: bold;'] * len(s)
    else: return ['color: black;'] * len(s)

def ask1(operation):
    for i, k in enumerate(files):
        print("[{}] {}".format(i, k))
    mrl = int(input('Which integer represents the mrl file?'))
    mrldf = pd.read_excel(files[mrl])
    if operation == 1:
        cfr = int(input('Which integer represents the cfr file?'))
        mrldf = process_cfr(cfrpath=files[cfr], mrldf=mrldf)
    if operation == 2:
        mrlview = int(input('Which integer represents the mrlview file?'))
        mrldf = process_mrlview(mrlviewpath=files[mrlview], mrldf=mrldf)
    if operation == 3:
        tip = int(input('Which integer represents the tip file?'))
        mrldf = process_tip(tippath=files[tip], mrldf=mrldf)
    return mrldf


def ask2(operation):
    for i, k in enumerate(files):
        print("[{}] {}".format(i, k))
    mrl = int(input('Which integer represents the mrl file?'))
    mrldf = pd.read_excel(files[mrl])
    if operation == 4:
        cfr = int(input('Which integer represents the cfr file?'))
        mrlview = int(input('Which integer represents the mrlview file?'))
        mrldf = process_mrlview(mrlviewpath=files[mrlview], mrldf=mrldf)
        mrldf = process_cfr(cfrpath=files[cfr], mrldf=mrldf)
    elif operation == 5:
        cfr = int(input('Which integer represents the cfr file?'))
        tip = int(input('Which integer represents the tip file?'))
        mrldf = process_tip(tippath=files[tip], mrldf=mrldf)
        mrldf = process_cfr(cfrpath=files[cfr], mrldf=mrldf)
    elif operation == 6:
        mrlview = int(input('Which integer represents the mrlview file?'))
        tip = int(input('Which integer represents the tip file?'))
        mrldf = process_tip(tippath=files[tip], mrldf=mrldf)
        mrldf = process_mrlview(mrlviewpath=files[mrlview], mrldf=mrldf)

    return mrldf



def ask3():
    for i, k in enumerate(files):
        print("[{}] {}".format(i, k))
    mrl = int(input('Which integer represents the mrl file? '))
    cfr = int(input('Which integer represents the cfr file? '))
    mrlview = int(input('Which integer represents the mrlview file? '))
    tip = int(input('Which integer represents the tip file? '))
    mrldf = pd.read_excel(files[mrl])
    mrldf = process_tip(tippath=files[tip], mrldf=mrldf)
    mrldf = process_mrlview(mrlviewpath=files[mrlview], mrldf=mrldf)
    mrldf = process_cfr(cfrpath=files[cfr], mrldf=mrldf)
    return mrldf

def trimdate(tempmrl, colname):
    tempmrl[colname] = tempmrl[colname].map(str)
    look_up = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May',
        '06': 'Jun', '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'}
    for num, val in enumerate(tempmrl[colname]):
        try:
            # If the first character isnt a digit the value won't be written
            cell = str(val)
            int(cell[0])
            date = cell.split(' ')[0]
            temparray = date.split('-') #YEAR, MONTH, DAY
            newval = str(temparray[2]) + "-" + look_up[str(temparray[1])] + "-" + temparray[0][-2:]
            tempmrl.loc[num, colname] = newval
        except:
            tempmrl.loc[num, colname] = None
            continue
    return tempmrl

if __name__ == '__main__':
    print("Which files would you like to be processed (into mrl)? ")
    print("[1] cfr file only")
    print("[2] mrlview file only")
    print("[3] tip file only")
    print("[4] 1 and 2")
    print("[5] 1 and 3")
    print("[6] 2 and 3")
    print("[7] All")
    operation = int(input("Which operation would you like to pick (type an integer): "))
    if operation == 1 or operation == 2 or operation == 3:
        tempmrl = ask1(operation)
    elif operation == 4 or operation == 5 or operation == 6:
        tempmrl = ask2(operation)
    elif (operation == 7):
        tempmrl = ask3()
    else:
        print("Invalid input")
        quit()
    finishdoc(tempmrl)

    print("----------Success!----------")
