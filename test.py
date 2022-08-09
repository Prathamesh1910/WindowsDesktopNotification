from winotify import Notification, audio
import os
import time
import pandas as pd
import datetime
import random

os.chdir(r"D:\Python programming vscode\Desktop_Notification")

if __name__ == '__main__':

    ## * Reading the current date and time *
    today = datetime.datetime.now().strftime("%d-%m-%Y")
    now_time = datetime.datetime.now().strftime("%H:%M:%S")


    ## * Reading the excel file *
    df = pd.read_excel("german_words_data.xlsx")

    ## * Itering through rows of excel sheet and collecting the data *
    iteration = list(df.iterrows())

    ## * After logging on next day date will be changed to current one and all the repetitions will be changed to 0
    for j, item in iteration:
        if today > item['Today']:

            ## * Updating the date to today's date
            change_date = df.loc[j, 'Today']
            df.loc[j, 'Today'] = today

            ## * updating the repetitions to 0 as it is a new day, new beginning *
            rpt = df.loc[j, 'Repetition']
            df.loc[j, 'Repetition'] = 0
    
    ran = False # To check if the program ran or not

    ## * To note down the index of the word just displayed *
    writeInd = []

    ## * List of indexes of data in sheet so that according to the randomly generated index we can fetch the data *
    indexes = []
    for i in iteration:
        indexes.append(i[0])

    ## * Randomly generaing an integer within the range of 0 to total data/words in sheet
    random_index = random.randrange(indexes[0], len(indexes))

    ## * to increase the repetition by 1 once the word is displayed *
    def incrementRpt():
        for i in writeInd:
            rpt = df.loc[i, 'Repetition']
            df.loc[i, 'Repetition'] = rpt + 1
        df.to_excel('german_words_data.xlsx', index = False)

    ## * Open notepad function *
    def openNotepad(text_file):
        f= open("explanation.txt","w+")
        f.write(f"** This is the explanatory file **\n\n{msg}")
        f.close()

        osCommandString = "notepad.exe explanation.txt"
        toast.set_audio(audio.Reminder, loop=False)

        #toast.add_actions(label="Click here for explanation", launch=webbrowser.open('explanation.html'))

        toast.show()
        os.system(osCommandString)

        ## * After closing the notepad, the explanation.txt file will be automatically deleted to avoid storing unnecessary files *
        time.sleep(5)
        os.remove("D:\Python programming vscode\Desktop_Notification\explanation.txt")
        print("File removed")

    ## * Main program of generating notification * 
    for index, item in iteration:
        if item['Repetition'] < 2:                          # If the word has repeated twice then the code will not be executed for that particular word
            if item['Serial No.'] == random_index+1:        # Indexing starts from 0, so incremented by 1 so as to match to the serial number to fetch the data
                de = item['German']
                eng = item['English']
                ex = item['Example']

                ## * The actual notification *
                toast = Notification(
                    app_id="My app",
                    title = "***Look here!***",
                    msg = f"German: {de}\nEnglish: {eng}\nExample: {ex}",
                    duration="short"
                )
                ran = True
                msg = f"German: {de}\nEnglish: {eng}\nExample: {ex}"
                writeInd.append(random_index)               # Appending the index of the word that just displayed
                incrementRpt()                              # Calling the incrementRpt function to increment the repition by 1 in excel file
                openNotepad(msg)
                break

    ## * If all the words are repeated twice and the program is not executed, change all the repetitions back to 0 to start the execution of program *
    ## * However no word will be displayed this time. Only the repetitions will be changed back to 0 and we will start getting notifications afterwards *
    if ran == False:
        for k, item in iteration:
            rpt = df.loc[k, 'Repetition']
            df.loc[k, 'Repetition'] = 0
        df.to_excel('german_words_data.xlsx', index = False)




