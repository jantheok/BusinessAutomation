import win32com.client #Reading Outlook
import os #Folder operations
import datetime #Timestamp
import pandas as pd #We will use it for reading XLS / XLSX and writing to CSV
import numpy as np #To check empty values


#Connect to Outlook
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

#Print first email account address in connected Outlook
print('Monitored Email :' + mapi.Accounts[0].DeliveryStore.DisplayName)

#Select mails from the main folder
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items

#Here we will store the Excel forms for upload into VMD
outputDir = r"C:\BusinessAutomation\2021\01.VMD\VMDFormDropOff"

#Directory for Queue File (input for next script)
queueFile = r"C:\BusinessAutomation\2021\01.VMD\VMDProcessingFile\Queue.csv"

#Reading Excel Attachments
def readExcel(excelFile):
    df = pd.read_excel(r'%s' % excelFile, header=None)
    try:
        #VendorID, Change Reason and Requestor are mandatory fields
        vendorID = df.iloc[2,2]
        changeReason = df.iloc[10,2]
        requestor = df.iloc[11,2]
        #We are using numpy as Pandas assigns 'nan' in case no value available
        if np.isnan(vendorID) or np.isnan(changeReason) or np.isnan(requestor):
            print('Mandatory fields missing')
        else:
            #Extract all fields from an attachment
            vendorName = df.iloc[3,2]
            vendorAdress = df.iloc[4,2]
            vendorPerson = df.iloc[5,2]
            vendorEmail = df.iloc[6,2]
            vendorPhone = df.iloc[7,2]
            vendorBankAccount = df.iloc[8,2]

            #Prepare data for DataFrame
            data = [{'Sender' : sender,'Subject' : subject, 'Attachment': excelPath, 'VendorID' : vendorID, 'ChangeReason' : changeReason, 'Requestor' : requestor, 
            'VendorName': vendorName,'VendorAddress' : vendorAdress, 'VendorPerson' : vendorPerson, 'VendorEmail' : vendorEmail, 'VendorPhone' : vendorPhone, 
            'VendorBankAccount' : vendorBankAccount}]

            #Create Pandas DataFrame
            queueData = pd.DataFrame(data)

            # Save the queue information. If file exists, append the row. If not, create a file with headers
            if os.path.isfile(queueFile):
                queueData.to_csv(queueFile, mode='a', index=False, header=False)
            else:
                queueData.to_csv(queueFile, mode='a', index=False, header=True)
            print('Succesful')
    except:
        print('Error during Attachment processing')


#Processing Email items with basic Exception handling
try:
    for message in list(messages):
        #Get Sender and Subject
        sender = message.sender
        subject = message.subject
        print(subject)
        try:
            #Filter emails that contain 'VMD'in subject
            if 'VMD' in message.Subject:
                if message.Attachments.count > 0 :
                    for attachment in message.Attachments:
                        file_name = attachment.FileName
                        if 'VMD' in file_name:
                            #Create time and date stamp to avoid duplication in files - the upload to VMD will be handled by different script
                            timestamp = datetime.datetime.now() 
                            date = str(timestamp.date()) 
                            time = str(timestamp.strftime("%Y%m%d_%H%M%S.%f")) #More information on https://www.w3schools.com/python/python_datetime.asp

                            #Rename attachment and save it to output folder
                            file_name = date + "_" + time + "_" + file_name
                            excelPath = os.path.join(outputDir, file_name)
                            attachment.SaveASFile(excelPath)

                            print(f"Attachment {file_name} from {sender} saved")
                            #Run the reading excel funtion
                            readExcel(excelPath)
                        else:
                            print('Attachment does not contain VMD')
                else:
                    print("No Attachment")
            else:
                print("Subject does not contain VMD")
        except Exception as e:
            print("Error when identify attachment:" + str(e))
except Exception as e:
    print("Error when processing emails messages:" + str(e))
