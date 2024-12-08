import array
import Switches
import csv
import xlwings as xw

workbook = xw.Book('BotC-Stats.xlsx')
sheet = workbook.sheets['Sheet1']

def separate_file() -> array:
    with open('results.csv', newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        Data = []
        for row in reader:
            Data.append(row)
        return Data

def replace_player_array(Player: array) ->array:
    FixedPlayer = []
    for entry in Player:
        player = entry[0]
        player = player.lower()
        FixedPlayer.append(Switches.find_player(player))
    return FixedPlayer

def replace_role_array(Role: array) ->array:
    FixedRole = []
    for entry in Role:
        role = entry[0]
        role = role.lower()
        FixedRole.append(Switches.find_role(role))
    return FixedRole

def update_good_stat(column: int, row: int, good_win: int) -> None:
    if(good_win[0] == '0'): #since good has not won increment lost instead of win
        column = increment_col(column)
    cell = str(column) + str(row)
    print(cell)
    value = sheet[cell].value
    value += 1
    sheet[cell].value = value
    workbook.save()
    workbook.save(r'C:\Users\rainb\Documents\code\BotC-GroupTracker\BotC-Stats.xlsx')
    

def update_evil_stat(column: int, row: int, good_win: int) -> None:
    if(good_win[0] == '1'): #since good has won evil has not and thus must increment lost instead
        column = increment_col(column)
    cell = str(column) + str(row)
    value = sheet[cell].value
    value += 1
    sheet[cell].value = value
    workbook.save()
    workbook.save(r'C:\Users\rainb\Documents\code\BotC-GroupTracker\BotC-Stats.xlsx')

def update_stats(Data: array) -> None:
    i = 0
    Player = []
    Role = []
    for entry in Data:
        if(i == 0):
            good_win = entry
        elif(i % 2 == 0):
            Role.append(entry)
        else:
            Player.append(entry)
        i += 1
    i = 0
    column = replace_player_array(Player)
    row = replace_role_array(Role)
    for entry in row:
        if(row[i] >= 106):
            update_evil_stat(column[i],row[i],good_win)
        else:
            update_good_stat(column[i],row[i],good_win)
        i += 1

def increment_col(string: str):
    lst = list(string) 
    result = []
    while lst:
        carry, next_ = increment_char(lst.pop())
        result.append(next_)
        if not carry:
            break
        if not lst:
            result.append('A')   
    result += lst[::-1]
    return ''.join(result[::-1])

def increment_char(char: chr):
        if char in ('Z'):
            return 1, 'A'
        else:
            return 0, chr(ord(char) + 1)

def decrement_col(string: str):
    lst = list(string)
    result = []
    index = 0
    for char in lst:
        if(char == lst[len(lst)-1]):
            char = decrement_char(char)
        if(char != 'ERROR'):
            result.append(char)
        index += 1
    return ''.join(result)

def decrement_char(char: chr):
        if char in ('A'):
            return 'ERROR'
        else:
            return chr(ord(char) - 1)
