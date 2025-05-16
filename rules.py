

def rulesForProgram():
    flag = False
    print("📜 Read All Instructions Carefully Before Starting the Quiz:")
    print()
    print()
    print("            1️⃣ You will be asked a total of 5 questions.")

    print("            2️⃣ Each question will have multiple-choice options, but only one correct answer.")
    print("            3️⃣ You must enter the correct option letter (A/B/C/D) to submit your answer.")
    print("            4️⃣ Each correct answer awards you 5 points.")
    print("            5️⃣ There is no negative marking for incorrect answers.")
    print("            6️⃣ Questions cannot be skipped once displayed.")
    print("            7️⃣ ⚠️ Important: Do not switch to other applications while the quiz is running. If the system detects this behavior, the quiz will automatically clos  and you may be disqualified.")
    print("            8️⃣ You must select an option for the current question before moving on to the next one.")
    print("            9️⃣ 🧠 Think carefully before submitting an answer. Once an option is selected, you cannot revisit or change it.")
    print("            🔟 🚫 Malpractice is strictly prohibited. The use of external electronic devices like mobile phones, smartwatches, calculators, or similar tools is not allowe  during the quiz.")
    print()
    print()
    userResponse = input("Type \'Yes\' if You Are Ready And Want To Start Answering The Questions else enter \'No\': ").strip().lower()
    while (userResponse != "yes" and userResponse != "no"):
        print("Entered Input is Not Valid!")
        userResponse = input("Type \'Yes\' if You Are Ready And Want To Start Answering The Questions else enter \'No\': ").strip().lower()
    if userResponse == "yes":
        flag = True
        return flag
    else:
        return flag
