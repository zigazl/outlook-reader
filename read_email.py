import win32com.client
import sys
import os
import datetime


def fill_table(messages, yesterday, this_file):
    data = []
    for message in reversed(messages):
        date = message.SentOn.strftime("%d-%m-%y")
        if date == yesterday:
            data.append(message.Subject)
    data.sort()
    return data
    

def write_messages(table, this_file):
    for i in table:
        this_file.writelines(i+"\n")
        print(i)
    print("\n" + str(len(table)) + " job anomalies on " + str(yesterday))
    this_file.write("\n"+str(len(table)) + " job anomalies on " + str(yesterday))

    

if __name__ == "__main__":
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    job_failing_inbox = outlook.GetDefaultFolder(6).Folders.Item('BiSandbox').Folders.Item('Job failing')
    job_duration_unusual_inbox = outlook.GetDefaultFolder(6).Folders.Item('BiSandbox').Folders.Item('Job duration unusual')

    job_failing_messages = job_failing_inbox.Items
    job_duration_unusual_messages = job_duration_unusual_inbox.Items

    yesterday = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d-%m-%y")

    print(" ")
    for account in outlook.Accounts:
        print(account.DeliveryStore.DisplayName)
    print(" ")

    job_failing_file_name = "job_failing_report_"+yesterday+".txt"
    job_failing_file = open(job_failing_file_name, 'w+')

    job_duration_unusual_file_name = "job_duration_unusual_report_"+yesterday+".txt"
    job_duration_unusual_file = open(job_duration_unusual_file_name, 'w+')

    job_duration_table = fill_table(job_duration_unusual_messages,yesterday,job_duration_unusual_file)
    job_failing_table = fill_table(job_failing_messages, yesterday, job_failing_file)
    
    write_messages(job_duration_table, job_duration_unusual_file)
    write_messages(job_failing_table, job_failing_file)


    mail = outlook.CreateItem(1)
    mail.To = 'zlogar.ziga@gmail.com'
    mail.Subject = 'Job error report for '+ yesterday
    mail.Body = "V priponki pošiljam job error report za omenjeni datum. Sporočilo je avtomatsko BEEP BOOP BOP. LP ZZ"
    mail.Attachments.Add('D:/GitHub/python-for-outlook/'+job_failing_file)
    mail.Attachments.Add('D:/GitHub/python-for-outlook/'+job_duration_unusual_file)
    #mail.CC = 'somebody@company.com'
    mail.Send()
    #send_mail(mail, yesterday, job_duration_unusual_file, job_failing_file)