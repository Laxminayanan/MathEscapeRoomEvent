from greetParticipants import *
from rules import rulesForProgram
from responseSheetInExcel import *
from mailValidation import is_valid_email
from thankYou import thankYouForYourParticipation
from calculatingTime import get_time_difference
from excelCodeForJudgement import AppendParticipantResponseInMainExcelForOurJudgement
import os
import pandas as pd
import smtplib
from email.message import EmailMessage
import getpass
import numberOfparticipantsTillNow
import listofAvailableTressureNumbers
from listOfQuestionSets import listOfDictionariesContainingTheQuestions

# Option fetching
def asciiNumber(character):
    return ord(character)


def checkForValidOption(option):
    checklist = ['a','b','c','d']
    if option not in checklist:
        return -1
    else:
        return option

def validateOption(option):
    if len(option) > 1:
        return -1
    elif (len(option) == 0):
        return -1
    elif (asciiNumber(option) >= 65 and asciiNumber(option) <= 90 ) or (asciiNumber(option) >= 97 and asciiNumber(option) <= 122):
        optionLowerCase = option.lower()
        validOptionOrNot = checkForValidOption(optionLowerCase)
        if (validOptionOrNot != -1):
            return 0
        else:
            return -1
    else:
        return -1




# Class For Collecting The Credentials of the Partciants


class Credentials:
    def __init__(self):
        self.name = None
        self.rollNumber = None
        self.mailId = None
    def takeUserCredentials(self):
        self.takeName("Enter Your Full Name: ")
        self.takeRollNumber("Enter The Full Roll Number: ")
        self.takeMailId("Please enter your complete email ID (e.g., example@gmail.com). Ensure correct spelling, case sensitivity, and include the full domain (e.g., @gmail.com): ")
    def takeName(self,prompt):
        self.name = input(prompt).strip()


    def takeRollNumber(self,prompt):
        self.rollNumber = input(prompt).strip()
        while(len(self.rollNumber) != 10):
            print("Entered Roll Number is Not Valid, may be It is Not Full Like (24RA1A...), So Pls Enter The Full Roll Number!")
            self.rollNumber = input(prompt)

    def takeMailId(self,prompt):
        self.mailId = input(prompt).strip()
        while(is_valid_email(self.mailId) == False):
            print("Entered Mail ID Is Not Valid!, Pls Enter It Again!")
            self.mailId = input(prompt)
        print("Entered Mail ID is Valid!")


def takeOptionAsInput(prompt):
    option = input(prompt).strip()
    while validateOption(option) != 0:
        print("Entered Input For This Question is Not Valid, So Pls Enter The Option From Only (A/B/C/D)!")
        option = input(prompt)
    return option.lower()





# Function To Calculate The Length Of The Given Number.
def lengthOfNumber(number):
    count = 0
    while (number > 0):
        count+=1
        number//=10
    return count


# Function To State Whether The Inouted String Contains Only Digits or Not.
def isDigit(collectedStringInput):
    for i in collectedStringInput:
        if ord(i) >= 48 and ord(i) <= 57:
            continue
        else:
            return False
    return True






# Tressure Number Fetching

def inputHuntNumber(prompt):
    takenInput = input(prompt)
    while True:
        if (takenInput == ''):
            print("Entred Tressure Number Is Not Valid!, Enter Again.")
            takenInput = input(prompt)
        elif isDigit(takenInput) == True:
            if(int(takenInput) >= 1 and int(takenInput) <= 32): # 32 Beacuse There Are Only 32 Sets of Questions.Increase The Set Of Questions As Per The Requirement In The Future.
                nowCheckListOfValidTressures  = listofAvailableTressureNumbers.listOfTressures
                for i in nowCheckListOfValidTressures:
                    if int(takenInput) == i:
                        # OverWrittingTheListOfAuthours
                        dupOfListOfTressures = listofAvailableTressureNumbers.listOfTressures
                        dupOfListOfTressures.remove(int(takenInput))
                        leftOverValidTressures = dupOfListOfTressures
                        with open('listofAvailableTressureNumbers.py', 'w') as f:
                            f.write(f'listOfTressures = {leftOverValidTressures}\n')
                        return int(takenInput)
                else:
                    print("Entred Tressure Number Is Not Valid!, Enter Again.")
                    takenInput = input(prompt)
            else:
                print("Entred Tressure Number Is Not Valid!, Enter Again.")
                takenInput = input(prompt)
        else:
            print("Entred Tressure Number Is Not Valid!, Enter Again.")
            takenInput = input(prompt)
    return int(takenInput)







# Only 32 Tressures Will be Hidden Inside The Campus, Since There Are Only 32 Question Sets. If Needed Add More Sets of Questions and Increase The Tressures Count As Per The Requirement.


def yesOrNoSaveEndTime(prompt):
    flag = False
    userResponse = input(prompt).strip().lower()
    while (userResponse != "yes" and userResponse != "no"):
        print("Entered Input is Not Valid!")
        userResponse = input(prompt).strip().lower()
    if userResponse == "yes":
        flag = True
        return flag
    else:
        return flag


def endQuiz():
    print()
    print()
    print()
    print()
    result = yesOrNoSaveEndTime("You Had Attempted All The 5 Questions, Do You Want To Record The End Time (\"Type : Yes/No \"): ")
    return result

def yesOrNo():
    flag = False
    userResponse = input("Type \'Yes\' if You Need The Response Sheet of Your Attempt, else Type \'No\': ").strip().lower()
    while (userResponse != "yes" and userResponse != "no"):
        print("Entered Input is Not Valid!")
        userResponse = input("Type \'Yes\' if You Need The Response Sheet of Your Attempt, else Type \'No\': ").strip().lower()
    if userResponse == "yes":
        flag = True
        return flag
    else:
        return flag

def printStartResonse():
    print()
    print()
    print()
    print("                                                                                       Lets Start The Game")

entireScoreForAllFiveQuestions = 0


def sendMail(recipientMail,excelFileName):
    recipient_email = recipientMail
    sender_email = "laxminayanan546@gmail.com"
    app_password = ""  # Using the 16-digit app password from Selected Mail - Service.

    subject = "Response Sheet Of The Math Escape Room"
    body = "Please find the attached Excel file."

    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.set_content(body)


    excel_file_path = f"{excelFileName}"+".xlsx"
    with open(excel_file_path, 'rb') as f:
        file_data = f.read()
        file_name = f.name

    msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(msg)
        return 0
    except Exception as e:
        return -1
    

def main():
    global entireScoreForAllFiveQuestions
    listOfResponsesOfEachQuestion = []
    for i in range(6):
        listOfResponsesOfEachQuestion.append(None)
    listsOflistOfResponseOfEachQuestion = []
    print()
    print()
    print()
    treasureHuntNumber = inputHuntNumber("Enter The Tressure Hunt Number You Had Found: ")
    print()
    print()
    print()
    counter = 0
    for i in listOfDictionariesContainingTheQuestions[treasureHuntNumber - 1]:
        counter += 1

        # Creating a new list for each question's response
        listOfResponsesOfEachQuestion = [None] * 6

        listOfResponsesOfEachQuestion[0] = counter
        listOfResponsesOfEachQuestion[1] = treasureHuntNumber
        listOfResponsesOfEachQuestion[2] = counter
        print(i["Question"])
        print()
        for j in i["Options"]:
            print(j)
        listOfResponsesOfEachQuestion[3] = takeOptionAsInput("Enter The Option (A/B/C/D): ")
        listOfResponsesOfEachQuestion[4] = i["Answer"][0].lower()
        if (listOfResponsesOfEachQuestion[3] == listOfResponsesOfEachQuestion[4]):
            listOfResponsesOfEachQuestion[5] = "5/5"
            entireScoreForAllFiveQuestions+=5
        else:
            listOfResponsesOfEachQuestion[5] = "0/5"
            entireScoreForAllFiveQuestions+=0

        listsOflistOfResponseOfEachQuestion.append(listOfResponsesOfEachQuestion.copy())
        print()
        print()
        print()
    return listsOflistOfResponseOfEachQuestion



def firstIterationJudgementExcelCreation(name,firstRow):

    # Converting To The DataFrame
    df = pd.DataFrame([firstRow], columns=[
        'Participant1Name', 'Participant2Name', 
        'Participant - 1 RNo.', 'Participant - 2 RNo.','ScoreForAll5Questions', 'TimeTakenToCompleteThetest'
    ])
    # Saving to Excel
    excel_filename = f"{name}.xlsx"
    try:
        df.to_excel(excel_filename, index=False, engine='openpyxl')
        return 0
    except Exception as e:
        return -1

entireTimetakenToCompleteTheTest = None
greet_participants() # Function Call for Greeting  The Participants
statusOfSUserResonseToStartTheGame = rulesForProgram()
if (statusOfSUserResonseToStartTheGame == True):
    printStartResonse()
    participant1 = Credentials()  # Creating The Object/Instance "participant1" of Credentials Class.
    print()
    print()
    print()
    print("Enter The Details of Participant 1: ")
    participant1.takeUserCredentials()
    print("\nEntered Participant 1 Details: ")
    print("    Entered Name Of Participant - 1: ",participant1.name)
    print("    Entered Roll Number Of Participant - 1: ",participant1.rollNumber)
    print("    Entered Mail id of Participant - 1: ",participant1.mailId)

    participant2 = Credentials()  # Creating The Object/Instance "participant2" of Credentials Class.
    print()
    print()
    print()
    print("Enter The Details of Participant 2: ")
    participant2.takeUserCredentials()
    print("\nEntered Participant 2 Details: ")
    print("    Entered Name Of Participant - 2: ",participant2.name)
    print("    Entered Roll Number Of Participant - 2: ",participant2.rollNumber)
    print("    Entered Mail id of Participant - 2: ",participant2.mailId)
    excelFileName = (participant1.rollNumber + participant2.rollNumber) # Execl File Name Is The Concatenation of The Two partcipants Roll Numbers.
    # LORFEQ -> List Of Responses For Each Question.


    # Recording The start time
    start_time = datetime.now().strftime("%H:%M:%S")
    print()
    print()
    print()
    print("Recorded Quiz Started Time:", start_time)
    recievedLORFEQToCreateTheDataFrame = main()
    # Recording end time Once the Answering of All 5 Questions is Completed
    end_time = datetime.now().strftime("%H:%M:%S")
    quizEndResult = endQuiz()
    if (quizEndResult == True):
        print()
        print()
        print()
        print("Recorded Quiz Ended Time:", end_time)
    else:
        print("Ok, That's Not A problem")
    timeTakenToCompleteThetest = get_time_difference(start_time,end_time)
    flagResultOfCreationOfUserResponse = createRespnseExcelForParticipants(f"{excelFileName}",recievedLORFEQToCreateTheDataFrame)
    if (flagResultOfCreationOfUserResponse == 0):
        print()
        print("Your All Answers Were Recorded Successfully!")

        # Appending The Participants Response To Our Main Excel For Our Final JudgeMent.
        listOfValuesOfRowForFinalJudgement = [None for i in range(1, 7)]
        listOfValuesOfRowForFinalJudgement[0] = participant1.name
        listOfValuesOfRowForFinalJudgement[1] = participant2.name
        listOfValuesOfRowForFinalJudgement[2] = participant1.rollNumber
        listOfValuesOfRowForFinalJudgement[3] = participant2.rollNumber
        listOfValuesOfRowForFinalJudgement[4] = entireScoreForAllFiveQuestions
        listOfValuesOfRowForFinalJudgement[5] = timeTakenToCompleteThetest
        appendResultForOurJudgement = None 
        if(numberOfparticipantsTillNow.numberOfParticipantsTillNow == 0): # In The First Iteration We Will be First Creating The mainExcelForFinalJudgement.xlsx with The listOfValuesOfRowForFinalJudgement.
            appendResultForOurJudgement = firstIterationJudgementExcelCreation("mainExcelForFinalJudgement.xlsx",listOfValuesOfRowForFinalJudgement)
            # OverWrittingTheSiNoInTheEveryRubOfTheProgram
            dupOfnumberOfparticipantsTillNow = numberOfparticipantsTillNow.numberOfParticipantsTillNow
            dupOfnumberOfparticipantsTillNow += 1
            with open('numberOfparticipantsTillNow.py', 'w') as f:
                f.write(f'numberOfParticipantsTillNow = {dupOfnumberOfparticipantsTillNow}\n')      
        
        else:
            appendResultForOurJudgement = AppendParticipantResponseInMainExcelForOurJudgement("mainExcelForFinalJudgement.xlsx",listOfValuesOfRowForFinalJudgement)
            # OverWrittingTheSiNoInTheEveryRubOfTheProgram
            dupOfnumberOfparticipantsTillNow = numberOfparticipantsTillNow.numberOfParticipantsTillNow
            dupOfnumberOfparticipantsTillNow += 1
            with open('numberOfparticipantsTillNow.py', 'w') as f:
                f.write(f'numberOfParticipantsTillNow = {dupOfnumberOfparticipantsTillNow}\n')      
                
        if (appendResultForOurJudgement == 0):
            print("Your Entire Activity With The System Has been Stored In Our DataBase.")
        else:
            print("Your Details and Responses For The Questions Are Not Stored In Our Database, So Please Contact The Faculty Coordinator or Student Coordinator As Soon As Possible.")
        resultOfresponseNeed = yesOrNo()
        if (resultOfresponseNeed == True):

            # Sending The ResponseExcel To Mail of the Participant1.
            resonseSheetToMailOfparticipant1Result = sendMail(participant1.mailId,f"{excelFileName}")
            if (resonseSheetToMailOfparticipant1Result == 0):
                print("The Response Sheet Has Been Successfully🎉🎊 Mailed To The ",participant1.name,".",sep = '')
            else:
                print("Sorry Error⚠️ occured While Mailing The Response Sheet To The ",participant1.name,".",sep = '')
            # Sending The ResponseExcel To Mail of the Participant2.
            resonseSheetToMailOfparticipant2Result = sendMail(participant2.mailId,f"{excelFileName}")
            if (resonseSheetToMailOfparticipant2Result == 0):
                print("The Response Sheet Has Been Successfully🎉🎊 Mailed To The ",participant2.name,".",sep = '')
            else:
                print("Sorry Error⚠️ occured While Mailing The Response Sheet To The ",participant2.name,".",sep = '')
            
            # If mail sended To Participant1 and Partcipant2 Then print("ResponseExecl is Sended To the Mails of Two Participants") and print The remaining Other cases Accordingly.
            if(resonseSheetToMailOfparticipant1Result == 0 and resonseSheetToMailOfparticipant2Result == 0):
                print("ResponseExecl is Successfully Sended To The Both Mail id'S of", participant1.name,"and",participant2.name)
            elif (resonseSheetToMailOfparticipant1Result == 0 and resonseSheetToMailOfparticipant2Result != 0):
                print("ResponseExecl Successfully Sended for Only", participant1.name)
            elif(resonseSheetToMailOfparticipant2Result == 0 and resonseSheetToMailOfparticipant1Result != 0):
                print("ResponseExecl Successfully Sended for Only", participant2.name)
            else:
                print("Sorry 🙇‍♂️ Unable to Send The Response Excel For The Both Participants!", participant1.name)
            thankYouForYourParticipation(participant1.name,participant2.name)
        else:
            print("\n\nOk! Not A problem")
            # Thanking The Participants
            thankYouForYourParticipation(participant1.name,participant2.name)
            numberOfparticipantsTillNow.numberOfParticipantsTillNow += 1
             # OverWrittingTheSiNoInTheEveryRubOfTheProgram
            dupOfnumberOfparticipantsTillNow = numberOfparticipantsTillNow.numberOfParticipantsTillNow
            dupOfnumberOfparticipantsTillNow += 1
            with open('numberOfparticipantsTillNow.py', 'w') as f:
                f.write(f'numberOfParticipantsTillNow = {dupOfnumberOfparticipantsTillNow}\n')

    else:
        print("Your Answers Were Not Recorded Due To Technical Issue, Pls Inform This Situation To The Student Coordinator or Faculty Coordinator Of This Event.")
else:
    print("Ok Take Your Time,But Start The Game As Quickly As Possible")