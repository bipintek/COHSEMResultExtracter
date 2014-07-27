import requests
from bs4 import BeautifulSoup
import xlsxwriter
import timeit

def isInt(value):
    try:
        a = int(value)
    except ValueError:
        return False
    else:
        return True


def posting(regno):
    url = "http://manresults.nic.in/hse/index.htm"
    link = "http://manresults.nic.in/hse/results.asp"
    payload = {'regno':regno,'B1':'Submit'}
    headers = {
                'User-Agent':'Mozilla/5.0',
                'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Encoding':'gzip, deflate',
                'Accept-Language':'en-US,en;q=0.5',
                'Connection':'keep-alive',
                'Cookie':'manresultcookie=HSE; ASPSESSIONIDQSSRCSSA=HFBJFHGBLACNKAJBGIAICDGA',
                'Host':'manresults.nic.in',
                'Referer':'http://manresults.nic.in/hse/index.htm'
            }
    resp = requests.post(link,headers=headers,data=payload)
    print "Requesting data from server for roll no "+str(regno)
    soup = BeautifulSoup(resp.text)
    data = soup.body.table.table.table.findAll(text=True)
    
    for i in data :
        if len(i)<2 :
            data.remove(i)
    data = data[0:54]
    for i in data:
        result.append(i.string)  
    return 
def writeToFile(result,row,col):
    worksheet.write('A1','Roll_no',)
    worksheet.write('B1','Name',)
    worksheet.write('C1','English',)
    worksheet.write('D1','MIL/ALT',)
    worksheet.write('E1','Phy',)
    worksheet.write('F1','Chm',)
    worksheet.write('G1','Mth',)
    worksheet.write('H1','Bio',)
    worksheet.write('I1','Csc',)
    worksheet.write('J1','Hsc',)
    worksheet.write('K1','Name of Institute',)
    worksheet.write('L1','Total',)
    name = result[3]
    institute = result[9]

    eng = 0
    mil = 0
    phy = 0
    chm = 0
    mth = 0
    bio = 0
    csc = 0
    hsc = 0

    first = result[32]
    second = result[38]
    third = result[44]
    fourth = result[49]
    first = first[-3:]
    second = second[-3:]
    third = third[-3:]
    fourth = fourth[-3:]

    eng = result[21]
    mil = result[27]
    
    if first=="Phy":
        phy = result[33]
    elif first=="Chm":
        chm=result[33]
    elif first=="Mth":
        mth = result[33]
    elif first == "Bio":
        bio = result[33]
    elif first == "Csc":
        csc = result[33]
    elif first == "Hsc":
        hsc = result[33]
    else :
       mark = 0
    
    if second=="Phy":
        phy = result[39]
    elif second=="Chm":
        chm=result[39]
    elif second=="Mth":
        mth = result[39]
    elif second == "Bio":
        bio = result[39]
    elif second == "Csc":
        csc = result[39]
    elif second == "Hsc":
        hsc = result[39]
    else :
        mark = 0

    if third=="Phy":
        phy = result[45]
    elif third=="Chm":
        chm=result[45]
    elif third=="Mth":
        mth = result[45]
    elif third == "Bio":
        bio = result[45]
    elif third == "Csc":
        csc = result[45]
    elif third == "Hsc":
        hsc = result[45]
    else :
        mark = 0
    
    if fourth=="Phy":
        phy = result[50]
    elif fourth=="Chm":
        chm=result[50]
    elif fourth=="Mth":
        mth = result[50]
    elif fourth == "Bio":
        bio = result[50]
    elif fourth == "Csc":
        csc = result[50]
    elif fourth == "Hsc":
        hsc = result[50]
    else :
        mark = 0


    if isInt(eng):
        eng = int(eng)
    else:
        eng = 0
    if isInt(mil):
        mil = int(mil)
    else:
        mil = 0

    if isInt(phy):
        phy = int(phy)
    else:
        phy = 0
    if isInt(chm):
        chm = int(chm)
    else:
        chm = 0

    if isInt(mth):
        mth = int(mth)
    else:
        mth = 0
    if isInt(bio):
        bio = int(bio)
    else:
        bio = 0

    if isInt(csc):
        csc = int(csc)
    else:
        csc = 0
    if isInt(hsc):
        hsc = int(hsc)
    else:
        hsc = 0
    #A list of marks of elective subjects
    electives = [phy,chm,mth,bio,csc,hsc]
    electives.sort(reverse=True)
    total = eng+mil+electives[0]+electives[1]+electives[2]
       
    
    print "Printing marks to file for roll no "+str(regno1)
    worksheet.write(row,col,regno1,)
    worksheet.write(row,col+1,name,)
    worksheet.write(row,col+2,eng,)
    worksheet.write(row,col+3,mil,)
    worksheet.write(row,col+4,phy,)
    worksheet.write(row,col+5,chm,)
    worksheet.write(row,col+6,mth,)
    worksheet.write(row,col+7,bio,)
    worksheet.write(row,col+8,csc,)
    worksheet.write(row,col+9,hsc,)
    worksheet.write(row,col+10,institute,)
    worksheet.write(row,col+11,total,)
    print "Done printing...."

    return


start = timeit.default_timer()
#Creating a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('resultlist.xlsx')
worksheet = workbook.add_worksheet()
count = 0;
result = list()
regno1 = 1
row = 1
col = 0
print "Preparing......"
while regno1<16556:
    
    try:
        posting(regno1)
        if len(result)>5:
            if result[20]!='English':
                print result[20]
                print 'Invalid candidate'
                count += 1
            else:
                writeToFile(result,row,col)
        else:
            print 'Invalid candidate'
            count += 1
    except UnicodeEncodeError:
        print 'Invalid candidate'
        count += 1
        
    regno1 += 1
    del result[:]
    row += 1
    col = 0
print "Finished extraction..... "
print "Invalid roll nos :"+str(count)
stop = timeit.default_timer()

print stop - start
workbook.close()


    
