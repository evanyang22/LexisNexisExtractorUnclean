#!/usr/bin/env python
# coding: utf-8

# In[1]:


#loops through each text file and creates array with all of them
import os
masterLineArray=[] #master array of all the stuff, use lineArray[0] for each one
masterTxtName=[] #masterArray for all the text names

for filename in os.scandir():
    if filename.is_file() and filename.path.endswith('.txt') and filename.path.endswith("_Table.txt")== False: #makes sure it is text file
        with open(filename.path) as f: 
            lineArray=f.readlines()
            masterLineArray.append(lineArray)
        masterTxtName.append(filename.path[2:])
        
        
#creates two corresponding lists, one with text names and one with the text arrays


# In[ ]:





# In[2]:


results=[]
addThis={}
completeARList=[]
completeDRList=[]

masterCounter=0
#for lineArray in masterLineArray:
while masterCounter<len(masterLineArray):
    lineArray=masterLineArray[masterCounter]
    addThis={}
    
    #adds pdfname
    addThis["PDFName"]=masterTxtName[masterCounter][:-3]+"pdf" #strips txt and replaces with pdf
    # EXTRACTING INFORMATION SECTION BELOW
    #extracts gender
    #should work in theory, not executed yet since test is blank
    counter=0
    while counter<len(lineArray):
        if("Gender" in lineArray[counter]):
            if("DOB" in lineArray[counter]):
                gender=lineArray[counter+1][lineArray[counter].find("Gender"):lineArray[counter].find("LexID")]
                addThis["Gender"]= gender
        counter=counter+1

    #extracts SSN by alternate method through voter registration
    SSNCounter=0
    RegistrantLine=0 #need to reset variables before each iteration
    VoterLine=0
    
    
    while SSNCounter<len(lineArray):
        if("Registrant Information" in lineArray[SSNCounter]):
            RegistrantLine=SSNCounter
        if("Voter Information" in lineArray[SSNCounter]):
            VoterLine=SSNCounter
        SSNCounter+=1
    newLoop=RegistrantLine
    
    while(newLoop<=VoterLine):
        if("SSN" in lineArray[newLoop]):
            SSN=lineArray[newLoop][lineArray[newLoop].find(":")+1:]
            addThis["SSN"]=SSN.strip()
        newLoop+=1
    

    #extracts emails
    emailMax=6
    emailCounter=0
    while emailCounter<len(lineArray):
        if("Email" in lineArray[emailCounter]):
            if("SSN" in lineArray[emailCounter]): #nested if statements helps extract email line
                i=1
                while i<=emailMax:
                    emailNum= "Email"+str(i)
                    emailHolder=lineArray[emailCounter+i][lineArray[emailCounter].find("Email"):]
                    #processing to remove \n at the end of each email
                    #apparently \n is treated as one charcter?????
                    if(emailHolder[-1:] == "\n"):
                        emailHolder=emailHolder[:-1]
                    addThis[emailNum]=emailHolder
                   # print(emailNum, emailHolder)
                    i=i+1
        emailCounter=emailCounter+1

    #extracts date of birth month and year 
    DOBcounter=0
    while DOBcounter<len(lineArray):
        if("DOB" in lineArray[DOBcounter]):
            if("SSN" in lineArray[DOBcounter] and "Gender" in lineArray[DOBcounter]):
                #bad formatting in text file forces me to go down 2 lines instead of 1 line
                DOB=lineArray[DOBcounter+2][lineArray[DOBcounter].find("DOB"):lineArray[DOBcounter].find("Gender")]
        DOBcounter=DOBcounter+1
    
    #splicing DOB
    #Question- any specification on formatting to clean data??? Ex: 4 vs april
    DOBMonth=DOB[:DOB.find("/")]
    DOBYear=DOB[DOB.find("/")+1:]
    addThis["DOBMonth"]=DOBMonth
    addThis["DOBYear"]=DOBYear.strip()
   # print("DOBMonth=" +DOBMonth)
   # print("DOBYear="+ DOBYear)

        
############## new lexID section
    oldPDFName= addThis["PDFName"]
    newPDFName=oldPDFName.replace(".pdf","_Table.txt")
    f = open(newPDFName,'r')
    tableFile=f.readlines() #array of strings
    
    IDCounter=0
    while IDCounter<len(tableFile):
        if("LexID" in tableFile[IDCounter]):
            if("Email" in tableFile[IDCounter]):
                lexID=tableFile[IDCounter+2][tableFile[IDCounter].find("LexID"):tableFile[IDCounter].find("Email")]
                addThis["LexID"]=lexID.strip()
            
        IDCounter=IDCounter+1
    
    


#######################end new lexID
        

    #Extracting current!!! Address and county
    #finding endpoint using additional personal information as endpoint marker
    endCounter=0
    endpoint=0
    while endCounter<len(lineArray):
        if("ADDITIONAL PERSONAL INFORMATION" in lineArray[endCounter]):
            endpoint=endCounter
        endCounter+=1
    
    addyCounter=0
    while addyCounter<len(lineArray):
        if("Address" in lineArray[addyCounter]):
            if("County" in lineArray[addyCounter] and "Phone" in lineArray[addyCounter]):
                a=1
                currentAddress=''
                b=endpoint-addyCounter-1
                lineCounter=0
                while a <= b: 
                    currentAddress+=lineArray[addyCounter+a][lineArray[addyCounter].find("Address"):lineArray[addyCounter].find("County")].strip()
                    if lineArray[addyCounter+a][lineArray[addyCounter].find("Address"):lineArray[addyCounter].find("County")].strip() != '':
                        currentAddress+=" " #adds single space
                        lineCounter+=1
                        if "COUNTY" in lineArray[addyCounter+a]:
                            addThis["County"]=lineArray[addyCounter+a][lineArray[addyCounter].find("Address"):lineArray[addyCounter].find("County")].strip()
                    a+=1
        addyCounter+=1
   # print("CurrentAddress="+ currentAddress)
    addThis["CurrentAddress"]= currentAddress.replace(addThis["County"],"").strip()

    #adds to county from this section
    
   # countyCounter=0
    #while countyCounter<len(tableFile):
     #   if("Full Name" in tableFile[countyCounter] and "County" in tableFile[countyCounter] and "Address" in tableFile[counterCounter]):
      #      print(tableFile[countyCounter+2])
    
    #extracts phone
    #currently extracting only the one from the beginning records page
    #note here- some minor formatting differences between pdf and text file
    phoneCounter=0
    while phoneCounter<len(lineArray):
        if("Phone" in lineArray[phoneCounter]):
            if("County" in lineArray[phoneCounter] and "Address" in lineArray[phoneCounter]): #makes sure it is the beginning record page
                phoneNum=lineArray[phoneCounter+1][lineArray[phoneCounter].find("Phone"):]
        phoneCounter=phoneCounter+1

   # print("phone number ="+ phoneNum)
    addThis["PhoneNumber"]=phoneNum.strip()


    #extracting name
    #bad formatting issue, seems Full Name on additional information is always in line with "COUNTY" though

   

    addThis["FirstName"]=addThis["PDFName"][:addThis["PDFName"].find("_")].strip()
    addThis["LastName"]=addThis["PDFName"][addThis["PDFName"].find("_")+1:-4].strip()
    addThis["FullName"]=addThis["FirstName"]+' ' + addThis["LastName"]

    #reading excel file with moodys linked in data
    #using panda to analyze and work with the excel data
    import pandas as pd
    workbook = pd.read_excel('Moodys_LinkedIn_Data.xlsx')
    workbook.head()

    analystData=workbook.loc[:,["id","analyst"]]
    analystDataList=analystData.values.tolist()

    #installed fuzzywuzzy package
    from fuzzywuzzy import fuzz

    #fuzzy matching to analyst name
    #adding analystname and ID to list
    
    
    fullName= addThis["FullName"]
    holdRatio=0
    
    for item in analystDataList:
        uncleanName=item[1]
        if uncleanName[-5:] == ", CFA": #strips CFA
            uncleanName=uncleanName[:-5]
        uncleanName=uncleanName.replace('.', '') #strips periods
        ratio=fuzz.ratio(fullName.lower(),uncleanName.lower())
        
        if ratio>holdRatio:
            holdRatio=ratio
            addThis["analystName"]=uncleanName
            addThis["analystID"]=item[0]
            addThis["fuzzRatio"]=holdRatio
        
        
   ############################################# all address date range section
    #address date range- detailed address information
    # needs work with multiple addresses
    # might be a challenge, bad formatting causes this to be difficult

    #calculates number of addresses
 #   addressCounter=0
  #  while addressCounter<len(lineArray):
   #     if("Address Summary" in lineArray[addressCounter]):
#            #print(lineArray[addressCounter][lineArray[addressCounter].find("-")+1:lineArray[addressCounter].find("records")])
 #           numAddresses=lineArray[addressCounter][lineArray[addressCounter].find("-")+1:lineArray[addressCounter].find("records")]
  #          numAddresses=int(numAddresses)
   #     addressCounter+=1

    #finds beginning point and endpoint
    #beginning point = "Address Details"
 #   beginCounter=0
  #  startPoint=0
   # while beginCounter<len(lineArray):
    #    if("Address Details" in lineArray[beginCounter]):
     #       startPoint=beginCounter
      #  beginCounter+=1


    #create two arrays, one for the addresses and one for the date ranges
#    k=1
#    searchThis=str(k)+":"
#    addyArray=[]
 #   dateArray=[]
  #  while startPoint<len(lineArray):
   #     if searchThis in lineArray[startPoint]:
    #        #print(lineArray[startPoint])
     #       addressHolder=lineArray[startPoint][2:]
      #      if("Dates" in addressHolder):
       #         addressHolder=addressHolder[:addressHolder.find("Dates")] # just some cleaning
        #    if("       " in addressHolder):
#                addressHolder=addressHolder[:addressHolder.find("      ")] #cleaning- removes a bunch of spaces and then phone numer
 #           addyArray.append(addressHolder.strip()) # more cleaning
  #          k+=1
   #         if k>numAddresses: #if k is larger than number of address, k gets changed to null and thus rendered useless
    #            k='null'
     #       searchThis=str(k)+":"
    
    #    if("Dates" in lineArray[startPoint] and "Phone" in lineArray[startPoint]):
     #       #dateRange= "None"
      #      dateRange=lineArray[startPoint+1][lineArray[startPoint].find("Dates"):lineArray[startPoint].find("Phone")] #goes down one line and takes date range, might be blank
       #     dateArray.append(dateRange.strip())
        
       # startPoint+=1

        
    #fixing error if date array length is less than address array length
    #adds blanks until len(addyArray) matches len of dateArray
#    if len(addyArray)>len(dateArray):
 #       while len(addyArray)>len(dateArray):
  #          dateArray.append("")
    
    #adding these dates to address Key value pair dictionairy
#    e=0
 #   addressRangeList=[]
  #  while e<len(addyArray):
   #     holdingDict={}
#        holdingDict["PropertyAddress"]=addyArray[e].replace(":","").strip()
 #       holdingDict["PropertyAddress"]=holdingDict["PropertyAddress"].strip()
  #      holdingDict["DateRange"]=dateArray[e]
   #     addressRangeList.append(holdingDict)
    #    e+=1
    
    #print("Address Range List")
    #print(addressRangeList)
    
    #take each dictionairy in address range list and merge it with analyst dictionairy (addThis)
#    fullAnalystAddressList=[]
 #   for tempDict in addressRangeList:
  #      tempHolder={**addThis,**tempDict} #merges into tempHolder dict
   #     fullAnalystAddressList.append(tempHolder)
    

    ####################### end address date range section
    
    #######new property address section
        #address date range- detailed address information
    # needs work with multiple addresses
    # might be a challenge, bad formatting causes this to be difficult

    #calculates number of addresses
    addressCounter=0
    while addressCounter<len(tableFile):
        if("Address Summary" in tableFile[addressCounter]):
            numAddresses=tableFile[addressCounter][tableFile[addressCounter].find("-")+1:tableFile[addressCounter].find("records")]
            numAddresses=int(numAddresses)
        addressCounter+=1

    #finds beginning point and endpoint
    #beginning point = "Address Details"
    beginCounter=0
    startPoint=0
    while beginCounter<len(tableFile):
        if("Address Details" in tableFile[beginCounter]):
            startPoint=beginCounter
        beginCounter+=1


    #create two arrays, one for the addresses and one for the date ranges
    k=1
    searchThis=str(k)+":"
    addyArray=[]
    dateArray=[]
    while startPoint<len(tableFile):
        if searchThis in tableFile[startPoint]:
            #print(lineArray[startPoint])
            addressHolder=tableFile[startPoint][2:]
            if("Dates" in addressHolder):
                addressHolder=addressHolder[:addressHolder.find("Dates")] # just some cleaning

            ##replaces double spaces with single spaces
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')  
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            addressHolder=addressHolder.replace("  ",' ')
            
            addyArray.append(addressHolder.strip()) # more cleaning
            k+=1
            if k>numAddresses: #if k is larger than number of address, k gets changed to null and thus rendered useless
                k='null'
            searchThis=str(k)+":"
    
        if("Dates" in tableFile[startPoint] and "Phone" in tableFile[startPoint]):
            #dateRange= "None"
            #print(tableFile[startPoint+2])
            dateRange=tableFile[startPoint+2][tableFile[startPoint].find("Dates"):tableFile[startPoint].find("Phone")] #goes down one line and takes date range, might be blank
            dateArray.append(dateRange.strip())
        
        startPoint+=1

        
    #fixing error if date array length is less than address array length
    #adds blanks until len(addyArray) matches len of dateArray
  #  if len(addyArray)>len(dateArray):
   #     while len(addyArray)>len(dateArray):
    #        dateArray.append("")
    
    #adding these dates to address Key value pair dictionairy
    e=0
    addressRangeList=[]
    while e<len(addyArray):
        holdingDict={}
        holdingDict["PropertyAddress"]=addyArray[e].replace(":","").strip()
        holdingDict["PropertyAddress"]=holdingDict["PropertyAddress"].strip()
        holdingDict["DateRange"]=dateArray[e]
        addressRangeList.append(holdingDict)
        e+=1
    
    #print("Address Range List")
    #print(addressRangeList)
    
    #take each dictionairy in address range list and merge it with analyst dictionairy (addThis)
    fullAnalystAddressList=[]
    for tempDict in addressRangeList:
        tempHolder={**addThis,**tempDict} #merges into tempHolder dict
        fullAnalystAddressList.append(tempHolder)
    
    
    
    
    #############end new property address section
    
    #results.append(addThis)
    #analystsOnly.appaddend(addThis)
    #print(analystsOnly)
    results=results + fullAnalystAddressList
    #print(addThis)
    

    
    ### new assessmentStart and assessmentEnd section begin
    tempCounter=0
    assessmentStart=0
    assessmentEnd=0
    
    firstInstance=True
    firstEndInstance=True
    
    while tempCounter < len(lineArray):
        if("Assessment Record" in lineArray[tempCounter] or "Deed Record" in lineArray[tempCounter]):
            if firstInstance: #finds the first instance of assessment record or deed record and sets it as assessment start
                assessmentStart=tempCounter
                firstInstance= False
        tempCounter+=1
    
    
    tempCounter=0
    while tempCounter < len(lineArray):
        if("Boats - " in lineArray[tempCounter] or "Potential Relatives -" in lineArray[tempCounter]):
            if firstEndInstance and tempCounter>assessmentStart: #finds first ending instance of boats or potential relatives
                #note- boat should come first before potential relatives 
                assessmentEnd=tempCounter
                firstEndInstance= False
        tempCounter+=1
    #print(assessmentStart, assessmentEnd)
    #### end assessmentstart and assessment end section
    
    #looping through the real property section the numRecord times and assessing each '#:'
    recordDict={}
    i=1

    #creating assessment and deed key value pair dictionairy list
    assessmentRecords=[]
    deedRecords=[]
    
    
    ####
    count=assessmentStart
    search=str(i)+ ":"
    while count<assessmentEnd:
        if(search in lineArray[count]):
            #print("Found "+ search)
            if(lineArray[count][lineArray[count].find(search)+2:lineArray[count].find("Record")].strip()=="Assessment"):
                #assessment records
                addThis2={} #key value pair dictionairy to be added
                endpoint=count+1
                while lineArray[endpoint].find("Record for")==-1 and lineArray[endpoint].find("Boat")==-1 and lineArray[endpoint].find("Potential Relatives")==-1:# returns -1 if not found, loops until it finds "Record for"
                    endpoint+=1
                recordDict["Assessment"+str(search)]=[]
                #loops to extract each interesting variable
                temp=count
                while temp<endpoint:
                    if("Address" in lineArray[temp]):
                        addThis2["ARAddress"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("County" in lineArray[temp]):
                        addThis2["ARCounty"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Recording Date" in lineArray[temp]):
                        addThis2["ARRecordingDate"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Sale Date" in lineArray[temp]):
                        addThis2["ARSaleDate"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Sale Price" in lineArray[temp]):
                        addThis2["ARSalePrice"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Assessed Value" in lineArray[temp]):
                        addThis2["ARAssessedValue"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Market Land Value" in lineArray[temp]):
                        addThis2["ARMarketLandValue"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Market Improvement Value" in lineArray[temp]):
                        addThis2["ARMarketImprovementValue"]= lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Total Market Value" in lineArray[temp]):
                        addThis2["ARTotalMarketValue"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    temp+=1      
                assessmentRecords.append(addThis2)
               # print("Assessment")
            if(lineArray[count][lineArray[count].find(search)+2:lineArray[count].find("Record")].strip()=="Deed"):
                #deed records
                addThis3={}
                endpoint=count+1
                while(lineArray[endpoint].find("Record for")==-1 and lineArray[endpoint].find("Boat")==-1 and lineArray[endpoint].find("Potential Relatives")==-1):# returns -1 if not found, loops until it finds "Record for"
                    endpoint+=1
                recordDict["Deed"+str(search)]=[]
                #loops to extract each interesting variable
                temp=count
                while temp<endpoint:
                    if("Address" in lineArray[temp]):
                        addThis3["DRAddress"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Contract Date" in lineArray[temp]):
                    #print(lineArray[temp][lineArray[temp].find(":")+1:])
                        addThis3["DRContractDate"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Recording Date" in lineArray[temp]):
                    #print(lineArray[temp][lineArray[temp].find(":")+1:])
                        addThis3["DRRecordingDate"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Loan Amount" in lineArray[temp]):
                    #print(lineArray[temp][lineArray[temp].find(":")+1:])
                        addThis3["DRLoanAmount"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Loan Type" in lineArray[temp]):
                    #print(lineArray[temp][lineArray[temp].find(":")+1:])
                        addThis3["DRLoanType"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Title Company" in lineArray[temp]):
                    #print(lineArray[temp][lineArray[temp].find(":")+1:])
                        addThis3["DRTitleCompany"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Transaction Type" in lineArray[temp]):
                    #print(lineArray[temp][lineArray[temp].find(":")+1:])
                        addThis3["DRTransactionType"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Description" in lineArray[temp]):
                    #print(lineArray[temp][lineArray[temp].find(":")+1:])
                        addThis3["DRDescription"]=lineArray[temp][lineArray[temp].find(":")+1:].strip()
                    if("Lender Information" in lineArray[temp]):# lender information is formatted differently so go down a line to extract
                    #print(lineArray[temp+1].find(":")+1)
                        addThis3["DRLenderInformation"]=lineArray[temp+1].find(":")+1
                    temp+=1
            
                deedRecords.append(addThis3)
                #print("Deed")
            i+=1
            search=str(i)+":"
        count+=1


#creating a single key value pair dictionairy for each address

#compares assessment records and merges ones with similiar address

    cleanAR=[]
    cleanDR=[]

    a=0
    while a<len(assessmentRecords):
        base=assessmentRecords[a]
        b=a+1
        while b<len(assessmentRecords)-a:
            comparison=assessmentRecords[b]
            ratio=fuzz.ratio(base["ARAddress"].lower(),comparison["ARAddress"].lower())
            if(ratio>80): #levenhstein ration >80
                agregaEsto={**base,**comparison}
                cleanAR.append(agregaEsto)
            b+=1
        a+=1
            

    #compares deedrecords and merges one with similiar address
    a=0
    while a<len(deedRecords):
        base=deedRecords[a]
        b=a+1
        while b<len(deedRecords)-a:
            comparison=deedRecords[b]
            ratio=fuzz.ratio(base["DRAddress"].lower(),comparison["DRAddress"].lower())
            if(ratio>80): #levenhstein ration >80
                agregaEsto={**base,**comparison}
                cleanDR.append(agregaEsto)
            b+=1
        a+=1
            
    #if clean records are empty, then it sets them equal to the original record lists because it means no match
    if(len(cleanAR)==0):
        cleanAR=assessmentRecords
    if(len(cleanDR)==0):
        cleanDR=deedRecords
    
    #print(cleanAR)
    #print(cleanDR)
    
    ### creation of master list by merging addThis (dictionairy of identifying information) with cleanAR/DR
    for thing in cleanAR:
        addAR={**addThis,**thing}
        completeARList.append(addAR)
    
    for thing in cleanDR:
        addDR={**addThis,**thing}
        completeDRList.append(addDR)
    
    masterCounter+=1


# In[23]:


for item in analystsOnly:
    print(item)


# In[27]:


for item in results:
    print(item)


# In[19]:


#writing master AR and DRList with the correct logic
import csv
csv_columns=["analystName","analystID","fuzzRatio","PDFName", "FullName", "FirstName", "LastName","County","PhoneNumber","SSN","DOBMonth","DOBYear","Gender","LexID","Email1","Email2","Email3","Email4","Email5","Email6","CurrentAddress","PropertyAddress","DateRange"]
AR_columns=["ARAddress","ARCounty","ARRecordingDate","ARSaleDate","ARSalePrice","ARAssessedValue","ARMarketLandValue","ARMarketImprovementValue","ARTotalMarketValue"]
DR_columns=["DRAddress","DRContractDate","DRRecordingDate","DRLoanAmount","DRLoanType","DRTitleCompany","DRTransactionType","DRDescription","DRLenderInformation"]
all_columns=csv_columns+AR_columns+DR_columns

masterList=[]

for addyDict in results: #cycles through ALL addresses
    mergeThis={}
    for ARDict in completeARList:
        tempRatio= fuzz.ratio(ARDict["ARAddress"].lower().strip(),addyDict["PropertyAddress"].lower().strip()) #levenhstein ratio
        if tempRatio> 90: #arbitrary set point, set very high due to mismatches
            addyDict={**addyDict,**ARDict} 
            
    for DRDict in completeDRList:
        holdingRatio= fuzz.ratio(DRDict["DRAddress"].lower().strip(),addyDict["PropertyAddress"].lower().strip())
        if holdingRatio>90:
            addyDict={**addyDict,**DRDict}
    
    masterList.append(addyDict)
    #print(addyDict)

#creating master excel
csv_file="alphabetizedMasterARDR.csv"
masterList=sorted(masterList,key=lambda x: x['FirstName'])
    
try:
    with open(csv_file,'w') as csvfile:
        writer=csv.DictWriter(csvfile,fieldnames=all_columns)
        writer.writeheader()
        for data in masterList:
            writer.writerow(data)
except IOError:
    print("IOError")


# In[ ]:





# In[5]:


#writing analysts only file
#for item in analystsOnly:
    #print(item)

import csv

csv_columns=["analystName","analystID","fuzzRatio","PDFName", "FullName", "FirstName", "LastName","County","PhoneNumber","SSN","DOBMonth","DOBYear","Gender","LexID","Email1","Email2","Email3","Email4","Email5","Email6","CurrentAddress","PropertyAddress","DateRange"]
dict_data=analystsOnly
#print(dict_data)
csv_file ="AnalystsOnly.csv"
try:
    with open(csv_file,'w') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
        writer.writeheader()
        for data in dict_data:
            writer.writerow(data)
except IOError:
    print("I/O error")


# In[12]:


cleanAnalystsOnly=[]
#for item in analystsOnly:
 #   item.pop("PropertyAddress",None)
  #  if item not in cleanAnalystsOnly:
   #     cleanAnalystsOnly.append(item)
        

for i in range(len(analystsOnly)):
    if analystsOnly[i] not in analystsOnly[i + 1:]:
        cleanAnalystsOnly.append(analystsOnly[i]) 
        
    
print(cleanAnalystsOnly)


# In[ ]:





# In[ ]:





# In[ ]:





# In[12]:


#writing masterAR and DR list with reversed logic 
import csv
csv_columns=["analystName","analystID","fuzzRatio","PDFName", "FullName", "FirstName", "LastName","County","PhoneNumber","SSN","DOBMonth","DOBYear","Gender","LexID","Email1","Email2","Email3","Email4","Email5","Email6","CurrentAddress","PropertyAddress","DateRange"]
AR_columns=["ARAddress","ARCounty","ARRecordingDate","ARSaleDate","ARSalePrice","ARAssessedValue","ARMarketLandValue","ARMarketImprovementValue","ARTotalMarketValue"]
DR_columns=["DRAddress","DRContractDate","DRRecordingDate","DRLoanAmount","DRLoanType","DRTitleCompany","DRTransactionType","DRDescription","DRLenderInformation"]
all_columns=csv_columns+AR_columns+DR_columns


#merge every dict in results and merge every dict in completeARList and completeDRList
#we reversed the logic####
masterList=[]
trueMasterList=[]
for ARDict in completeARList:
    largestRatio=0 #finds highest levenhstein ratio
    mergeThis={}
    for addyDict in results:
        tempRatio= fuzz.ratio(ARDict["ARAddress"].lower().strip(),addyDict["PropertyAddress"].lower().strip())
        if tempRatio>largestRatio:
            mergeThis=addyDict
            largestRatio=tempRatio

    mergedDict={}
    mergedDict={**ARDict,**mergeThis}
    print(mergedDict)
    masterList.append(mergedDict)
    
for DRDict in completeDRList:
    print("DRDict")
    
    largeRatio=0
    mergeEse={}
    for largeDict in masterList:
        holdingRatio=fuzz.ratio(DRDict["DRAddress"].lower().strip(),largeDict["PropertyAddress"].lower().strip())
        if holdingRatio>largeRatio:
            mergeEse=largeDict
            largeRatio=holdingRatio
    
    mergeDict={}
    mergeDict={**DRDict,**mergeEse}
    trueMasterList.append(mergeDict)
    

#creating master excel
csv_file="masterAR_DR.csv"

    
try:
    with open(csv_file,'w') as csvfile:
        writer=csv.DictWriter(csvfile,fieldnames=all_columns)
        writer.writeheader()
        for data in trueMasterList:
            writer.writerow(data)
except IOError:
    print("IOError")


# In[2]:


#dictionairy tests
empty={}
testDict={"key1":"hey","key2":"hello"}
merger={**empty,**testDict}
print(merger)
merge2={**testDict,**empty}
print(merge2)


# In[11]:


#writing completeARList and completeDRList to excel 

import csv
csv_columns=["analystName","analystID","fuzzRatio","PDFName", "FullName", "FirstName", "LastName","County","PhoneNumber","SSN","DOBMonth","DOBYear","Gender","LexID","Email1","Email2","Email3","Email4","Email5","Email6","CurrentAddress","PropertyAddress","DateRange"]
AR_columns=["ARAddress","ARCounty","ARRecordingDate","ARSaleDate","ARSalePrice","ARAssessedValue","ARMarketLandValue","ARMarketImprovementValue","ARTotalMarketValue"]
DR_columns=["DRAddress","DRContractDate","DRRecordingDate","DRLoanAmount","DRLoanType","DRTitleCompany","DRTransactionType","DRDescription","DRLenderInformation"]

#creating AR excel
csv_file="AR.csv"
allARColumns=csv_columns+AR_columns
try:
    with open(csv_file,'w') as csvfile:
        writer=csv.DictWriter(csvfile,fieldnames=allARColumns)
        writer.writeheader()
        for data in completeARList:
            writer.writerow(data)
except IOError:
    print("I/O error")
        

#creating DR excel
csv_file="DR.csv"
allDRColumns=csv_columns+DR_columns
    
try:
    with open(csv_file,'w') as csvfile:
        writer=csv.DictWriter(csvfile,fieldnames=allDRColumns)
        writer.writeheader()
        for data in completeDRList:
            writer.writerow(data)
except IOError:
    print("IOError")


# In[3]:


#converting to CSV file, creating analysts.csv
import csv

csv_columns=["analystName","analystID","fuzzRatio","PDFName", "FullName", "FirstName", "LastName","County","PhoneNumber","SSN","DOBMonth","DOBYear","Gender","LexID","Email1","Email2","Email3","Email4","Email5","Email6","CurrentAddress","PropertyAddress","DateRange"]
dict_data=results
#print(dict_data)
csv_file ="AnalystsAndAllAddresses.csv"
try:
    with open(csv_file,'w') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
        writer.writeheader()
        for data in dict_data:
            writer.writerow(data)
except IOError:
    print("I/O error")


# In[ ]:





# In[ ]:





# In[ ]:


#don't run below bc it generates the text files with -layout formula


# In[1]:


# creates table version with -table 
import os, sys, subprocess

for filename in os.scandir():
    if filename.is_file():
        filePath=filename.path
        pdfName=filePath[filePath.find('.')+2:filePath.find(".pdf")]
        txtFileName= pdfName+"_Table.txt"
        command= "C:\\Users\\evany_cdhq038\\OneDrive\\Desktop\\Economics_Research\\xpdf-tools-win-4.03\\bin64\\pdftotext.exe -table C:\\Users\\evany_cdhq038\\OneDrive\\Desktop\\Economics_Research\\All_Reports\\"
        command= command + pdfName+".pdf " + txtFileName
        print(command)
        
        p= subprocess.Popen(command,shell=True, stdout=subprocess.PIPE)
    


# In[ ]:





# In[2]:


import os    
import sys
import subprocess
for filename in os.scandir():
    if filename.is_file():
        filePath=filename.path
        pdfName=filePath[filePath.find('.')+2:filePath.find(".pdf")]
        txtFileName= pdfName+".txt"
        command= "C:\\Users\\evany_cdhq038\\OneDrive\\Desktop\\Economics_Research\\xpdf-tools-win-4.03\\bin64\\pdftotext.exe -layout C:\\Users\\evany_cdhq038\\OneDrive\\Desktop\\Economics_Research\\All_Reports\\"
        command= command + pdfName+".pdf " + txtFileName
        print(command)
        
        p= subprocess.Popen(command,shell=True, stdout=subprocess.PIPE)


# In[ ]:


try:
			q = subprocess.Popen([PDFTOTEXT_PATH,'-layout',pdfPath,'-'], stdout=subprocess.PIPE, stderr=subprocess.PIPE) 
			pdfText, err = q.communicate()
		except:
			print 'pdftotext failed'


# In[ ]:


PDFTOTEXT_PATH = '/usr/local/bin/pdftotext'

