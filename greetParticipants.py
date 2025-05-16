from datetime import datetime

def greet_participants():
    current_hour = datetime.now().hour

    if 5 <= current_hour < 12:
        greeting = "Good Morning"
    elif 12 <= current_hour < 17:
        greeting = "Good Afternoon"
    elif 17 <= current_hour < 21:
        greeting = "Good Evening"
    else:
        greeting = "Hello"  # for night time or late hours
    print()
    print()
    print()
    print(f"                                                                           {greeting}, Welcome To The Math Escape Room!")
    print()
    print()
    print()

    