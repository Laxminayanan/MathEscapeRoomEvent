from mailValidation import is_valid_email

class Credentials:
    def __init__(self):
        self.name = None
        self.rollNumber = None
        self.mailId = None
    def takeUserCredentials(self):
        self.takeName("      Enter Your Full Name: ")
        self.takeRollNumber("      Enter The Full Roll Number: ")
        self.takeMailId("      Please enter your complete email ID (e.g., example@gmail.com). Ensure correct spelling, case sensitivity, and include the full domain (e.g., @gmail.com): ")
    def takeName(self,prompt):
        self.name = input(prompt).strip()


    def takeRollNumber(self,prompt):
        self.rollNumber = input(prompt).strip()
        while(len(self.rollNumber) != 10):
            print("      Entered Roll Number is Not Valid, may be It is Not Full Like (24RA1A...), So Pls Enter The Full Roll Number!")
            self.rollNumber = input(prompt)

    def takeMailId(self,prompt):
        self.mailId = input(prompt).strip()
        while(is_valid_email(self.mailId) == False):
            print("      Entered Mail ID Is Not Valid!, Pls Enter It Again!")
            self.mailId = input(prompt)
        print("      Entered Mail ID is Valid!")


if __name__ == "__main__":
    participant1 = Credentials()
    participant1.takeUserCredentials()
    # Checking The Entered Details
    print("Entered Details Are: ")
    print(participant1.name)
    print(participant1.rollNumber)
    print(participant1.mailId)
