import sys
import os
import xlrd
import xlsxwriter
from time import sleep
from random import randint

status = "on"
money = 0
help_score = 2
jokers = ["A) The 50/50", "B) The Audience", "C) The Telephone"]
result_columns = ["Question", "Correct Answer", "'s Answer", "Amount"]
row = 0
col = 0
total_amount = 0
file_location = os.getcwd() + '/questions.xlsx'
questions = []

def ask_question(question, answers, correct, amount, audience, phone):
  global col, row
  print(question)
  sleep(3)
  for answer in answers:
    print(answer)
    sleep(1)
  user_answer = input("What is your answer?(A-D or J for joker) ")
  worksheet.write(row, col, question)
  worksheet.write(row, col+1, correct)
  worksheet.write(row, col+2, user_answer)
  col = 3
  if user_answer.upper() == "J":
    use_joker(correct, amount, audience, phone)
  elif user_answer.upper() == correct:
    print(" ")
    correct_answer(amount)
    sleep(2)        
  else:
    global help_score
    if help_score > 0:
      help()
      for answer in answers:
        print(answer)
        sleep(1)
      user_answer2 = input("What is your answer?(A-D or J for joker) ")
      if user_answer2.upper() == "J":
        use_joker(correct, amount, audience, phone)
      elif user_answer2.upper() == correct:
        print(" ")
        correct_answer(amount)
        sleep(2)
      else:
        print(" ")
        game_over()
    else:
      print(" ")
      game_over()

def correct_answer(amount):
  print("I THINK.....")
  print("THAT'S CORRECT ANSWER !!!!")
  print(" ")
  global money, col, row
  money = amount
  worksheet.write(row, col, money)
  row += 1
  col = 0
  print(" ")
  sleep(1)
  print(f"WELL DONE {name}, YOU WON £{money}!")
  print(" ")

def use_joker(correct, amount, audience, phone):
  print(" ")  
  global jokers
  if len(jokers) == 0:
    print("Sorry, you have no jokers left!")
    sleep(2)
    user_answer = input("What is your answer? ")
    if user_answer.upper() == correct:
      print(" ")
      correct_answer(amount)
      sleep(2)
    else:
      print(" ")
      game_over()    
  else:    
    print("You have the following jokers:")
    sleep(2)
    for joker in jokers:
      print(f"{joker}-Joker")
      sleep(1)
    joker_selection = input("Which joker would you like to use?")
    if joker_selection.upper() == "A":
      jokers.remove("A) The 50/50")
      jokerA(correct, amount)
    elif joker_selection.upper() == "B":
      jokers.remove("B) The Audience")
      jokerB(correct, amount, audience)
    elif joker_selection.upper() == "C":
      jokers.remove("C) The Telephone")
      jokerC(correct, amount, phone)

def jokerA(correct, amount):
  answers = ["A", "B", "C", "D"]
  joker_answer = [correct]  
  answers.remove(correct)
  number = randint(0, 2)
  joker_answer.append(answers[number])
  joker_answer.sort()
  sleep(1)
  print(".")
  sleep(1)
  print("..")
  sleep(1)
  print("...")
  sleep(1)  
  print(f"The remaining answers are {joker_answer[0]} and {joker_answer[1]}")
  sleep(2)
  user_answer = input("What is your answer? ")
  if user_answer.upper() == correct:
    print(" ")
    correct_answer(amount)
    sleep(2)
  else:
    print(" ")
    game_over()

def jokerB(correct, amount, audience):
  sleep(1)
  print(".")
  sleep(1)
  print("..")
  sleep(1)
  print("...")
  sleep(1)
  print(f"The audience vote is: {audience}")
  sleep(2)
  user_answer = input("What is your answer? ")
  if user_answer.upper() == correct:
    print(" ")
    correct_answer(amount)
    sleep(2)
  else:
    print(" ")
    game_over()

def jokerC(correct, amount, phone):
  sleep(1)
  print(".")
  sleep(1)
  print("..")
  sleep(1)
  print("...")
  sleep(1)
  print(f"Here is what your Telephone Joker said:")
  sleep(1.5)
  print(phone)
  sleep(2)
  user_answer = input("What is your answer? ")
  if user_answer.upper() == correct:
    print(" ")
    correct_answer(amount)
    sleep(1)
  else:
    print(" ")
    game_over()

def help():
  global help_score
  help_score -= 1
  sleep(1.5)
  print(" ")
  print("...are you SURE that is correct?")
  sleep(2)
  print("again the possibilities are:")
  sleep(2)    

def game_over():
  global status, col, row
  worksheet.write(row, col, 0)
  row += 2
  worksheet.merge_range('A'+str(row)+':'+'C'+str(row), 'Winning Amount')
  worksheet.write('D'+str(row), money)
  workbook.close()
  status = "off"
  print("That is wrong answer!!!!")
  sleep(1)
  print(f"sorry {name}, you lost!")
  print(" ")
  print(" ")
  sleep(1)
  print("Winning Amount: £",money)
  print("##################################################################################")
  print("                                    GAME OVER "                                    )
  print("##################################################################################")

if os.path.isfile(file_location):
  wb = xlrd.open_workbook(file_location)
  sheet = wb.sheet_by_index(0)
  headers = [row.value for row in sheet.row(0)]
  for row_no in range(sheet.nrows):
    if row_no <= 0:
      continue
    line = list(map(lambda row:isinstance(row.value, bytes) and row.value.encode('utf-8') or str(row.value), sheet.row(row_no)))
    questions.append(
      {
        'question': line[0], 
        'answers': line[1].replace(' ', '').split(','),
        'correct': line[2],
        'amount': int(float(line[3])),
        'audience': line[4].replace(' ', '').split(','),
        'phone': line[5],
      })
else:
  sys.exit("File not Found!")

print(" ")
print(" ")
print(" ")  
print("Ladies and Gentlemen!")
print(" ")
sleep(1.3)
print("Welcome to the game")
print(" ")
sleep(0.7)
print("WHO WILL BE A MILLIONAIRE!!!!")
sleep(0.7)
sleep(1.3)

name = input("MAY I KNOW YOUR NAME PLEASE? (enter your name) ")
workbook = xlsxwriter.Workbook(str(name)+'_result.xlsx')
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format({'bold': True, 'font_size': 12})
worksheet.set_column('A1:D1', 20)
for column in result_columns:
  if col == 2:
    column = name+column
  worksheet.write(row, col, column, cell_format)
  col += 1
row += 1
col = 0
print("OK THEN LET'S START THE GAME !!")
print(" ")
sleep(1)
print("FIRST A REMONDER, YOU HAVE 3 JOKERS:")
sleep(1)
for joker in jokers:
  print(f"{joker}-Joker")
print("You can only use ONE joker for each question.")
print("OK, let's go!")
print(" ")
sleep(1.5)

for question in questions:
  if status == "on":  
    ask_question(question.get('question'), question.get('answers'), question.get('correct'), question.get('amount'), question.get('audience'), question.get('phone'))

if status == "on":
  worksheet.merge_range('A'+str(row)+':'+'C'+str(row), 'Winning Amount')
  worksheet.write('D'+str(row), money)
  workbook.close()
  print("CONGRATULAAAAAAAAAAAAATIONS!!!!!!")
  sleep(1)
  print("YOU ARE THE WINNER OF ONE MILLION POUNDS")
  print("THE END")
  sleep(1)
  print("THANK YOU FOR YOUR PARTICIPATION.")



