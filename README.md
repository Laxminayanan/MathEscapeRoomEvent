# MathEscapeRoom
- Math Escape Room is a hybrid digital-physical event where teams compete to solve Aptitude Questions within a time constraint. The Console Based Program manages the complete workflow from team registration to question delivery, response tracking, and result compilation. This system was successfully deployed during our college's Tech Fest, providing participants with an engaging and challenging experience.


# 🔄 Event Flow
- Registration: Teams of 3-5 participants arrive at the venue and register their team details.
- Device Submission: Participants submit their electronic gadgets to prevent unfair advantages.
- Clue Distribution: One team member randomly selects a clue/riddle from a transparent jar.
- Treasure Hunt: The team decodes the clue to find a specific numbered plastic ball hidden somewhere Within the campus.
- Time Constraint: Teams must return with the correct treasure number within 15 minutes or face elimination.
- Digital Challenge: Qualifying teams proceed to a computer station where they:

Enter their credentials
Input the found treasure number
Access a unique set of math problems tied to their treasure number
Solve the problems while the system tracks time and responses


Results: Upon completion, the system:

Generates an Excel file with detailed performance metrics
Automatically emails the results to the team using their registered email address



# ✨ Features

- Time-based Greeting: Dynamic greeting based on the time of day for a personalized experience
- Email Validation: Ensures valid email addresses for reliable communication
- Authentication System: Secure login for teams with credential verification
- Treasure-to-Question Mapping: Dynamic question set delivery based on the treasure number found
- Timer System: Accurate tracking of start and end times for fair evaluation
- Response Tracking: Records all participant answers for each question
- Score Calculation: Automatic evaluation of responses against correct answers
- Excel Report Generation: Comprehensive performance report creation in Excel format
- Automated Email Dispatch: Instant delivery of results to participant email addresses

# 🛠 Technologies Used

- Programming Language: Python
- Data Processing: Pandas for data manipulation and Excel file operations
- Email Functionality: smtplib and email libraries for automated communications
- Time Management: datetime module for precise timing operations
- Email Validation: dns.resolver for domain verification
- File Operations: os module for filesystem interactions
- Excel Integration: openpyxl engine for Excel file handling

# 🚀 Setup and Installation

#### Clone the repository:

```bash
git clone https://github.com/Laxminayanan/mathEscapeRoomEvent.git
```


# Install dependencies:
```bash
pip install pandas openpyxl dnspython
```

# Configure email settings:

- Update the email credentials in the sendMail function with your SMTP server details

📋 Usage
#### Run the main application:
- python main.py

# Follow the prompts to:

- Register team details
- Enter the treasure number found
- Complete the presented mathematical challenges
- View results and receive the emailed performance report
# ⚙ Configuration
### Email Configuration
The system uses SMTP to send emails. To configure your email settings:

- If using Gmail, enable "Less secure app access" or create an App Password
- Update the sender credentials in the sendMail function

### Question Sets
Questions are mapped to treasure numbers. To modify or add question sets:

Edit the listOfQuestionSets.py file. Ensure each treasure number has a corresponding set of questions

# 🤝 Contributing
Contributions are welcome! Please feel free to submit a Pull Request.
