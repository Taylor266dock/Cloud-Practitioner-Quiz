# Import Modules
# Lets us create the GUI
from PySimpleGUI import *
# Imports a module to reade from the Excel file
import xlrd
# Imports the random function letting us generate random questions
from random import randint
import os

# Initialises the Question and Answer Bank from the Excel sheet
Q = []
A = {}

# Excel grid explained - how to reference
# a1 = cell_value(0,0)
# a2 = cell_value(1,0)
# b3 = cell_value(2,1)

#############################
# Read from Excel and populate the above		#
#############################
book_read = xlrd.open_workbook(os.path.join("AWSCATrackerTemplate.xls"))  # Choose Excel file
sheet = book_read.sheet_by_name("Q&A")  # Choose Excel sheet
i = 1  # Choose start column/row (to be used for looping)
while True:
    try:
        f = sheet.cell_value(i, 2)  # Choose start column/row (row,column)
        Q.append(f)  # Add selected cell to list
        i += 1  # Increment i, change cell
    except IndexError:  # When cell is empty
        break  # End loop
# Read answers
for i in range(len(Q)):
    A[Q[i]] = sheet.cell_value(i + 1, 3)


# Initial Values for variables used in app
selection = ''
result = "NC"
noQuestions = 0
correctQ = 0
n = 0
pass_fail = ""
question_bank = {}
options_list = []
correct_answer = ""
answerA = ""
answerB = ""
answerC = ""
answerD = ""
qb = 0

# This generates a random set of 65 questions and answers to simulate an AWS Exam
while qb <= 64:
    if qb != len(question_bank):
        qb -= 1
    q = randint(1, len(Q) - 1)
    question_bank[Q[q]] = f"{A[Q[q]]}"
    qb += 1


# This generates 3 additional random answers and then shuffles them around with the correct answer
def random_answers():
    global options_list, correct_answer, answerD, answerC, answerB, answerA
    opt1 = A[Q[random.randint(1, len(Q) - 1)]]
    opt2 = A[Q[random.randint(1, len(Q) - 1)]]
    opt3 = A[Q[random.randint(1, len(Q) - 1)]]
    options_list = [opt1, opt2, opt3, list(question_bank.values())[n]]
    correct_answer = options_list[3]
    # using random shuffle function to mix up the answers otherwise the correct answer would be A every time
    random.shuffle(options_list)
    answerA = options_list[0]
    answerB = options_list[1]
    answerC = options_list[2]
    answerD = options_list[3]


# Initialise the first set of random answers
random_answers()


# GUI Layout
# This is the Main Quiz Layout
layout = [[Text(f"{list(question_bank)[n]}", key="-NXTQ-", size=(115, 5))],  # size=(width, lines)
          [Button('A'), Text(f"{answerA}", key="-A-")],
          [Button('B'), Text(f"{answerB}", key="-B-")],
          [Button('C'), Text(f"{answerC}", key="-C-")],
          [Button('D'), Text(f"{answerD}", key="-D-")],
          [Text('', size=(15, 1), font=('Helvetica', 18),
                text_color='black', key='input')],
          [Text(f"Number of Questions:{noQuestions}", key="-NOOFQ-")],
          [Text(f"Correct Answers:{correctQ}", key="-NOCOR-")],
          [Button('Quit')]
          ]

# This is the second screen, revealing the results of the test - Not yet working
final_screen = [[Text("Well Done")],
                [Text(f"You got {correctQ} answers!")],
                [Text(f"That is a {pass_fail}", key="-PF-")],
                [Button('Quit')]
                ]

# Set PySimpleGUI
window = FlexForm('AWS Cloud Architect', default_button_element_size=(5, 2),
                  auto_size_buttons=False, grab_anywhere=False, margins=(25, 25), size=(1000, 600),
                  element_justification="left")
window.Layout(layout)

# Make Infinite Loop
while True:
    # Button Values
    button, value = window.read()
    # Check Press Button Values, update result to the text in the answer
    if button == 'A':
        result = f"{answerA}"
        window.find_element('input').update(result)
    elif button == 'B':
        result = f"{answerB}"
        window.find_element('input').update(result)
    elif button == "C":
        result = f"{answerC}"
        window.find_element('input').Update(result)
    elif button == "D":
        result = f"{answerD}"
        window.find_element('input').Update(result)
        pass
    # close the window
    elif button == 'Quit':
        break
    # Results - checks result against the correct answer and lets the user know if they were right or wrong
    if result == list(question_bank.values())[n]:
        window.find_element('input').Update("Correct")
        correctQ += 1
        noQuestions += 1
    elif result != list(question_bank.values())[n]:
        window.find_element('input').Update("Incorrect")
        noQuestions += 1

    n += 1  # moves on to the next question
    random_answers()  # creates a new set of random answers
    # Updates the screen with a new question, answers and running totals
    window.find_element('-NXTQ-').Update(f"{list(question_bank)[n]}")
    window.find_element('-A-').Update(f"{answerA}")
    window.find_element('-B-').Update(f"{answerB}")
    window.find_element('-C-').Update(f"{answerC}")
    window.find_element('-D-').Update(f"{answerD}")
    window.find_element('-NOOFQ-').Update(f"Number of Questions:{noQuestions}")
    window.find_element('-NOCOR-').Update(f"Correct Answers:{correctQ}")

    # checks if number of questions correct was enough to pass the test
    if correctQ >= 46:
        pass_fail = "PASS!"
    else:
        pass_fail = "Fail"

    # checks if test is complete to move the user onto the next screen
    if noQuestions == 65:
        window.find_element('-PF-').Update(f"That is a {pass_fail}")