import xlwings as xw
import csv
import array
import discord
import Switches
from Token import token
from discord import app_commands
from discord.ext import commands
from pathlib import Path

#------bot stuff--------
bot = commands.Bot(command_prefix='!', intents = discord.Intents.all())
workbook = xw.Book('BotC-Stats.xlsx')
sheet = workbook.sheets['Sheet1']


@bot.event
async def on_ready():
    print(f'Running as {bot.user}')
    print(bot.user.id)

guild = discord.Object(id=token)
@bot.command()
@commands.guild_only()
@commands.is_owner()  # Prevent other people from using the command
async def sync(ctx: commands.Context) -> None:
    """Sync app commands to Discord."""
    try:
        synced = await bot.tree.sync()
        print(f"Synced {len(synced)} command(s)")
    except Exception as e:
        print(e)
    

@bot.tree.command(name="personalaverage", description="Check your average winrates for sides")
async def personalaverage(interaction: discord.Interaction):
    column = Switches.GetPlayerFromDiscord(interaction.user.name)
    column = incrementcol(column)
    column = incrementcol(column)
    cell = str(column)+str(102)
    TotalGood = sheet[cell].value
    cell = str(column)+str(156)
    TotalEvil = sheet[cell].value
    cell = str(column)+str(157)
    Total = sheet[cell].value
    await interaction.response.send_message(f"Your average winrate is {str(Total)} consisting of {str(TotalGood)} whilst good and {str(TotalEvil)} whilst evil!", ephemeral=True)



@bot.tree.command(name="personalrolestats", description="Check any your stats for a particular role")
@app_commands.describe(role = 'Role you would like to check (Townsfolk for Townsfolk total and Total Good for all good)')
async def personalrolestats(interaction: discord.Interaction, role: str):
    column = Switches.GetPlayerFromDiscord(interaction.user.name)
    column = decrementcol(column)
    row = Switches.FindRole(role.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True)
    else:
        data = []
        for i in range(4):     
            cell = str(column)+str(row)
            data.append(sheet[cell].value)
            column = incrementcol(column)
        await interaction.response.send_message(f'you have played the {role} {data[0]} times, of those you won {data[1]} and lost {data[2]} which makes your winrate {data[3]}.')

@bot.tree.command(name="playeraverage", description="Check any players average winrates")
@app_commands.describe(player = 'Player you would like to see the averages for')
async def playeraverage(interaction: discord.Interaction, player: str):
    column = Switches.FindPlayer(player.lower())
    if(column == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column = incrementcol(column)
        column = incrementcol(column)
        data = []
        rows = [Switches.FindRole('total good'),Switches.FindRole('total evil'),Switches.FindRole('total')]
        for i in range(3):       
            cell = str(column)+str(rows[i])
            data.append(sheet[cell].value)
        await interaction.response.send_message(f'{player}\'s average winrates are as follows: Good-{data[0]} Evil-{data[1]} Total-{data[2]}')


@bot.tree.command(name="roletotalstats", description="Check the stats for a particular role")
@app_commands.describe(role = 'Role you would like to check, (Townsfolk for Townsfolk total and Total Good for all good)')
async def roletotalstats(interaction: discord.Interaction, role: str):
    row = Switches.FindRole(role.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True) 
    else: 
        column = 'C'
        data = []
        for i in range(4):       
                cell = str(column)+str(row)
                data.append(sheet[cell].value)
                column = incrementcol(column)
        await interaction.response.send_message(f'The {role} has been played {data[0]} times of those they won {data[1]} and lost {data[2]} which makes their winrate {data[3]}.')

@bot.tree.command(name='playerrolestats', description="Check any players stats for a particular role")
@app_commands.describe(role = 'Role you would like to check, (Townsfolk for Townsfolk total and Total Good for all good)')
@app_commands.describe(player = 'Player you would like to check (same name as on spreadsheet)')
async def playerrolestats(interaction: discord.Interaction, role: str, player: str):
    row = Switches.FindRole(role.lower())
    column = Switches.FindPlayer(player.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True)
    elif(column == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        data = []
        for i in range(4):       
                cell = str(column)+str(row)
                data.append(sheet[cell].value)
                column = incrementcol(column)
        await interaction.response.send_message(f'{player} has played {role} {data[0]} times of those they won {data[1]} and lost {data[2]} which makes their winrate {data[3]}.', ephemeral=True)   
    
        

@bot.tree.command(name="uploaddatabase", description="Uploads the current entire database to be checked")
async def uploaddatabase(interaction: discord.Interaction):
    await interaction.response.send_message(file=discord.File(r'BotC-Stats.xlsx'))

@commands.is_owner()  # Prevent other people from using the command
@bot.tree.command(name="updatespreadsheet", description="if program has a csv file it uses it to update the spreadsheet")
async def updatespreadsheet(interaction: discord.Interaction)-> None:
    my_file = Path("Results.csv")
    if my_file.is_file():
        data = separateFile()
        print(data)
        updateStats(data)
        await interaction.response.send_message(f'spreadsheet updated', ephemeral=True)   

@bot.tree.command(name="updaterole", description="auto update roles from database")
async def updateroles(interaction : discord.Interaction):
    server = bot.get_guild(1192193425739612281)
    role0 = server.get_role(1294069983151919134)
    role20 = server.get_role(1294281073043177535)
    role40 = server.get_role(1294068836764614788)
    role60 = server.get_role(1294069566409801760)
    role80 = server.get_role(1294068801448574996)
    role100 = server.get_role(1294068667801276426)
    member = interaction.user
    column = Switches.GetPlayerFromDiscord(member.name)
    column = incrementcol(column)
    column = incrementcol(column)
    cell = column + '158'
    cellval = sheet[cell].value
    if(cellval < 20):
        await member.add_roles(role0)
        await interaction.response.send_message(f'You have been assigned the 0+ games role as you have not yet attended enough to increase', ephemeral=True)
    elif(cellval < 40):
        await member.add_roles(role20)
        await member.remove_roles(role0)
        await interaction.response.send_message(f'You have been assigned the 20+ games role as you have attended enough to increase', ephemeral=True)
    elif(cellval < 60):
        await member.add_roles(role40)
        await member.remove_roles(role20)
        await interaction.response.send_message(f'You have been assigned the 40+ games role as you have attended enough to increase', ephemeral=True)
    elif(cellval < 80):
        await member.add_roles(role60)
        await member.remove_roles(role40)
        await interaction.response.send_message(f'You have been assigned the 60+ games role as you have attended enough to increase', ephemeral=True)
    elif(cellval < 100):
        await member.add_roles(role80)
        await member.remove_roles(role60)
        await interaction.response.send_message(f'You have been assigned the 80+ games role as you have attended enough to increase', ephemeral=True)
    else:
        await member.add_roles(role100)
        await member.remove_roles(role80)
        await interaction.response.send_message(f'You have been assigned the 100+ games role as you have attended enough to increase', ephemeral=True)
        
        
def separateFile() -> array:
    with open('results.csv', newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        Data = []
        for row in reader:
            Data.append(row)
        return Data

def replacePlayerArray(Player: array) ->array:
    FixedPlayer = []
    for entry in Player:
        player = entry[0]
        player = player.lower()
        FixedPlayer.append(Switches.FindPlayer(player))
    return FixedPlayer

def replaceRoleArray(Role: array) ->array:
    FixedRole = []
    for entry in Role:
        role = entry[0]
        role = role.lower()
        FixedRole.append(Switches.FindRole(role))
    return FixedRole

def updateGoodStat(column: int, row: int, good_win: int) -> None:
    if(good_win == ['0']): #since good has not won increment lost instead of win
        column = incrementcol(column)
    cell = str(column) + str(row)
    value = sheet[cell].value
    value += 1
    sheet[cell].value = value
    workbook.save()
    workbook.save(r'C:\Users\rainb\Documents\code\BotC-GroupTracker\BotC-Stats.xlsx')
    

def updateEvilStat(column: int, row: int, good_win: int) -> None:
    if(good_win == ['1']): #since good has won evil has not and thus must increment lost instead
        column = incrementcol(column)
    cell = str(column) + str(row)
    value = sheet[cell].value
    value += 1
    sheet[cell].value = value
    workbook.save()
    workbook.save(r'C:\Users\rainb\Documents\code\BotC-GroupTracker\BotC-Stats.xlsx')

def updateStats(Data: array) -> None:
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
    print(Role)
    print(Player)
    column = replacePlayerArray(Player)
    row = replaceRoleArray(Role)
    print(column)
    print(row)
    print(good_win)
    for entry in row:
        if(row[i] >= 106):
            updateEvilStat(column[i],row[i],good_win)
        else:
            updateGoodStat(column[i],row[i],good_win)
        i += 1

def incrementcol(string: str):
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

def decrementcol(string: str):
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


bot.run('MTMwMzcxOTAwOTc5NTA0NzUyNg.GWHdIn.4qN6deiWx2QhX6rsC-YBQlPTeAjOKcMK90dbqM')