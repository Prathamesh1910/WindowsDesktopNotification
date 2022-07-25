# WindowsDesktopNotification
This program reads German words and other data from an excel file and displays a toast message as well as notepad file containing the information from excel file. This program can be automated by putting it on Task Scheduler.

Save both the files viz. the Python file and Excel file in the same folder.

In order to run the program, you need to have following Python modules:
1) winotify (Notification, audio)
2) os
3) time
4) pandas
5) datetime
6) random
7) xlrd
8) openpyxl (load_workbook)

No need to provide any path in "os.chdir()". It is needed only while running the program using Task Scheduler.
