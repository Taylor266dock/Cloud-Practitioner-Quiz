import random

import xlrd
from random import randint, shuffle

import os


Q = []
A = {}

#############################
#	Excel grid explained	#
#############################
#	a1 = cell_value(0,0)	#
#	a2 = cell_value(1,0)	#
#	b3 = cell_value(2,1)	#
#############################


name = 'Questions'		# Excel sheet name

#############################
#	Read from Excel 		#
#############################


book_read = xlrd.open_workbook(os.path.join("AWSCATrackerTemplate.xls"))		# Choose Excel file
sheet = book_read.sheet_by_name("Q&A")								# Choose Excel sheet
i = 1																# Choose start column/row (to be used for looping)
while True:
    try:
        f = sheet.cell_value(i,2)									# Choose start column/row (row,column)
        Q.append(f)													# Add selected cell to list
        i += 1														# Increment i, change cell
    except IndexError:												# When cell is empty
        break														# End loop


# Read answers
for i in range(len(Q)):
    A[Q[i]]=sheet.cell_value(i+1,3)



# Initial values
tries = 0
streak = 0
streakrecord = 0
q = randint(0,len(Q)-1)
stillPlaying = True
answer = ""
noQuestions = 0
correctQ = 0


print('+-------->')
print('| Quiz time: %s' %(name))
print('+---------------------------------+')
print("| Commands: 'quit' 'exit'         |")
print('+---------------------------------+')
print('| 1st question;')

def question():
    global answer
    global stillPlaying
    print('\n| %s' % (Q[q]))
    opt1 = A[Q[random.randint(1, 510)]]
    opt2 = A[Q[random.randrange(1, 511)]]
    opt3 = A[Q[random.randrange(1, 511)]]
    options_list = [opt1, opt2, opt3, A[Q[q]]]
    random.shuffle(options_list)
    print(f"A: {options_list[0]}")
    print(f"B: {options_list[1]}")
    print(f"C: {options_list[2]}")
    print(f"D: {options_list[3]}")
    answer = input("| Answer: ")
    if answer == "A":
        answer = options_list[0]
    elif answer == "B":
        answer = options_list[1]
    elif answer == "C":
        answer = options_list[2]
    elif answer == "D":
        answer = options_list[3]
    elif answer == 'quit' or answer == 'exit':
        print('| Correct answer was: %s' % (A[Q[q]]))
        stillPlaying = False
    else:
        print("Not a valid answer please try again.")
        question()
    print(answer)

# Starting questions loop
while stillPlaying == True:

    question()
    print(answer)
    tries += 1
    noQuestions += 1


    if answer != A[Q[q]]:
        print('| Wrong answer!')
    elif answer == A[Q[q]]:
        correctQ += 1
        percentQ = correctQ / noQuestions * 100
        print('| Correct! try: %i' %(tries))
        print(noQuestions)
        print(correctQ)
        print(f"{percentQ}%")
        if tries <= 1:
            streak += 1
        if streak >= 5:
            print("| Nice streak! %i in a row!" %(streak))
            if streak > streakrecord:
                print('+-------------------------->')
                print("| New streak record! %i" %(streak))
                print('+--------------------------------->')
        print('| Next Question;')
        q = randint(0,len(Q)-1)
        tries = 0




    elif tries > 5:
        print('| You failed 5 times, correct answer: %s' %(A[Q[q]]))
        tries = 0
        if streak > streakrecord:
            streakrecord = streak
        #		print '+-------------------------->'
        #		print "| New streak record! %i" %(streak)
        #		print '+--------------------------------->'
        streak = 0
        q = randint(0,len(Q)-1)
        print('| Next Question;')

