def asciiNumber(character):
    return ord(character)

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

def checkForValidOption(option):
    checklist = ['a','b','c','d']
    if option not in checklist:
        return -1
    else:
        return option


def takeOptionAsInput(prompt):
    option = input(prompt).strip()
    while validateOption(option) != 0:
        print("Entered Input For This Question is Not Valid, So Pls Enter The Option From Only (A/B/C/D)!")
        option = input(prompt)
    return option.lower()

def main():
    inputOption = takeOptionAsInput("Enter The Option (A/B/C/D): ")
    print(inputOption)


if __name__ == "__main__":
    main()