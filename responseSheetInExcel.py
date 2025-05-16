import pandas as pd
def createRespnseExcelForParticipants(name,listOfResponsesForEachQuestion):

    # Converting To The DataFrame
    df = pd.DataFrame(listOfResponsesForEachQuestion, columns=['S.No.', 'Question Set Number', 'Question Number', 'Choosen Option', 'Correct Option', 'Marks Scored In This Question'])
    # Saving to Excel
    excel_filename = f"{name}.xlsx"
    try:
        df.to_excel(excel_filename, index=False, engine='openpyxl')
        return 0
    except Exception as e:
        return -1