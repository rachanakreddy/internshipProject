import os, zipfile, gzip, shutil, csv, datetime
import xml.etree.ElementTree as Xet

#Temp folder path
dirPath = "C:\Data\Highstreet IT Solutions LLC\IT Documentation - General\DMARC\DMARC Temp" 

#extensions
zip = ".zip"
gz = ".gz"

#master csv data table
csvPath = "C:\Data\Highstreet IT Solutions LLC\IT Documentation - General\DMARC\DMARC CSV Data\masterDMARCData.csv"  

##only run to set up initial headers for a master csv file () - NOT NECESSARY if master table exists
#master = open(csvPath, "w") 
#writer = csv.writer(master)
#header = ["orgName", "reportID", "date", "sourceIP", "disposition", "dkim", "spf", "dmarc"]
#writer.writerow(header)

#opens in "appending" mode
master = open(csvPath, "a", newline = "") 
writer = csv.writer(master)

# for loop that iterates through gz and zip files in Temp folder to unzip, deletes compressed
for path in os.listdir(dirPath):
    if os.path.isfile(os.path.join(dirPath, path)): #check if it is a file
        if path.endswith(zip): #check if it is a zip
            fileName = os.path.join(dirPath, path)
            zipRef = zipfile.ZipFile(fileName, "r") 
            zipRef.extractall(dirPath) #extract zip file to temp folder
            zipRef.close()
            os.remove(fileName) #remove zip

        if path.endswith(gz): #checks if it is a gz
            fileName = os.path.splitext(path)[0] #new file name without ".gz" to write to, ends in .xml
            fileInPath = os.path.join(dirPath, path) #gz file path
            fileOutPath = os.path.join(dirPath,fileName) #new xml file path (output)
            f_in = gzip.open(fileInPath, "rb") 
            f_out = open(fileOutPath, "wb")
            shutil.copyfileobj(f_in, f_out) #copies unzipped gz file content to xml
            f_in.close()
            f_out.close()
            os.remove(fileInPath) #removes gz


#for loop to parse each xml file into csv
for path in os.listdir(dirPath):
    xmlFile = os.path.join(dirPath, path) 
    if xmlFile.endswith(".xml"):
        #try to parse and ignore parse error
        try:
            xmlParse  = Xet.parse(xmlFile)
        except Xet.ParseError:
            pass 
        
        #csv row information set to empty string
        orgName = reportID = date = sourceIP = disposition = dkim = spf = dmarc = ""

        #traverse report metadata tag
        for report in xmlParse.findall("report_metadata"):
            if(report is not None):
                orgName = getattr(report.find("org_name"), 'text', None)
                reportID = getattr(report.find("report_id"), 'text', None) 
                #go into date range tag
                for d in report.findall("date_range"):
                    date = getattr(d.find("end"), 'text', None)
                    date = int(date) 
                    date = datetime.datetime.fromtimestamp(date) #convert unix to regular date
                    date = date.date() #retrieve date with no time
       
        #traverse record tag
        for record in xmlParse.findall("record"):
            if(record is not None):
                #traverse row tag in each record
                for row in record.findall("row"):
                    sourceIP = getattr(row.find("source_ip"), 'text', None)

                    for policy in row.findall("policy_evaluated"):
                        disposition = getattr(policy.find("disposition"), 'text', None)

                #traverse auth result tag in each record
                for auth in record.findall("auth_results"):
                    for dkimAuth in auth.findall("dkim"):
                        dkim = getattr(dkimAuth.find("result"), 'text', None)

                    for spfAuth in auth.findall("spf"):
                        spf = getattr(spfAuth.find("result"), 'text', None)
            #check if dmarc passes
            if(dkim == 'pass' or spf == 'pass'):
                dmarc = 'pass'
            else:
                dmarc = 'fail'
            
            #write row to csv with information from each record (orgName, reportID, and date repeat)
            writer.writerow([orgName, reportID, date, sourceIP, disposition, dkim, spf, dmarc])
   
    #remove traversed  xml Fil   
    os.remove(xmlFile)
