import win32com.client
import openpyxl
import sys
import datetime

def writeToExcel(new_failed_list):
    
    workbook = openpyxl.Workbook()

    # 获取默认的工作表（Sheet1）
    sheet1 = workbook.active
    sheet1.title = "new_failed"
    sheet1.append(["testplan", "case_name", "failed_item", "owner", "failed_version"])
    for row in new_failed_list:
        sheet1.append(row)

    # 保存 Excel 文件
    workbook.save("D:\\auto-trace\\output_{}.xlsx".format(start_of_yesterday.strftime('%m-%d-%Y')))
    print("Excel 文件写入完成！")


# 获取当前日期和时间
now = datetime.datetime.now()

# 获取昨天和今天的日期
yesterday = now - datetime.timedelta(days=1)
today = now

# 构造昨天的日期
#start_of_yesterday = datetime.datetime(yesterday.year, yesterday.month, yesterday.day, 21, 0, 0)
start_of_yesterday = datetime.datetime(2024, 1, 29, 20, 0, 0)

# 连接 Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
Accounts = outlook.Folders

for account in Accounts:
    if account.Name == "yuhq@primarius-tech.com":
        user_account = account
        break

# 选择指定文件夹
Folders = user_account.Folders
for folder in Folders:
    if str(folder) == "Auto-trace":
        auto_trace_folder = folder
        break

# 构造筛选条件
filter_str = "[ReceivedTime] > '{}'".format(start_of_yesterday.strftime('%m/%d/%Y %H %p'))
emails = auto_trace_folder.Items.Restrict(filter_str)
emails_count = emails.Count
print("共有 {} 封auto-trace邮件".format(emails_count))
print("------")

if emails_count == 0:
    #print("No new mail.")
    sys.exit()

# 获取邮件内容并组织数据
end_lines = []
for each_email in emails:
    if not str(each_email.Subject).strip().startswith("Trace error "):
        continue
    print("主题：", each_email.Subject)
    print("------")
    lines = str(each_email.Body).split("\n")
    for each in lines:
        each = each.strip()
        if each.startswith("Hi"):
            each = each.strip(":")
            owner = (each.split(',')[1]).strip()
            continue
        if each == "":
            continue
        if each.startswith("testplan"):
            continue
        if "\t" in each:
            each_list = each.split("\t")
            each_list.insert(-1, owner)
            end_lines.append(each_list)

writeToExcel(end_lines)