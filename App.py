import xlwings as xw
import csv
import array
import discord
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

guild = discord.Object(id='1303745588302708807')
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
    column = GetPlayerFromDiscord(interaction.user.name)
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
async def roleaverage(interaction: discord.Interaction, role: str):
    column = GetPlayerFromDiscord(interaction.user.name)
    column = decrementcol(column)
    row = FindRole(role.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True)
    elif(row == 102 or row == 156 or row == 157):
        cell1 = 'C'+str(row)
        cell2 = 'F'+str(row)
        await interaction.response.send_message(f'your {sheet[cell1].value} is {sheet[cell2].value}')
    elif(row == 153 or row == 130 or row == 99 or row == 73):
        data = []
        for i in range(4):       
            cell = str(column)+str(row)
            data.append(sheet[cell].value)
            column = incrementcol(column)
        await interaction.response.send_message(f'you have been the {role} {data[0]} times, of those you won {data[1]} and lost {data[2]} which makes your winrate {data[3]}.', ephemeral=True)
    else:
        data = []
        for i in range(4):       
            cell = str(column)+str(row)
            data.append(sheet[cell].value)
            column = incrementcol(column)
        await interaction.response.send_message(f'you have played the {role} {data[0]} times, of those you won {data[1]} and lost {data[2]} which makes your winrate {data[3]}.', ephemeral=True)

@bot.tree.command(name="playeraverage", description="Check any players average winrates")
@app_commands.describe(player = 'Player you would like to see the averages for')
async def playeraverage(interaction: discord.Interaction, player: str):
    column = FindPlayer(player.lower())
    if(column == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column = incrementcol(column)
        column = incrementcol(column)
        data = []
        rows = [102,156,157]
        for i in range(3):       
            cell = str(column)+str(rows[i])
            data.append(sheet[cell].value)
        await interaction.response.send_message(f'{player}s average winrates are as follows: Good-{data[0]} Evil-{data[1]} Total-{data[2]}')


@bot.tree.command(name="roletotalstats", description="Check the stats for a particular role")
@app_commands.describe(role = 'Role you would like to check, (Townsfolk for Townsfolk total and Total Good for all good)')
async def roletotalstats(interaction: discord.Interaction, role: str):
    row = FindRole(role.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True) 
    elif(row == 102 or row == 156 or row == 157):
        cell1 = 'C'+str(row)
        cell2 = 'F'+str(row)
        await interaction.response.send_message(f'the {sheet[cell1].value} is {sheet[cell2].value}')
    elif(row == 153 or row == 130 or row == 99 or row == 73):
        column = 'C'
        data = []
        for i in range(4):       
            cell = str(column)+str(row)
            data.append(sheet[cell].value)
            column = incrementcol(column)
        await interaction.response.send_message(f'{role}(s) have been played {data[0]} times, of those they won {data[1]} and lost {data[2]} which makes their winrate {data[3]}.', ephemeral=True)
    else: 
        column = 'C'
        data = []
        for i in range(4):       
                cell = str(column)+str(row)
                data.append(sheet[cell].value)
                column = incrementcol(column)
        await interaction.response.send_message(f'The {role} has been played {data[0]} times of those they won {data[1]} and lost {data[2]} which makes their winrate {data[3]}.', ephemeral=True)

@bot.tree.command(name='playerrolestats', description="Check any players stats for a particular role")
@app_commands.describe(role = 'Role you would like to check, (Townsfolk for Townsfolk total and Total Good for all good)')
@app_commands.describe(player = 'Player you would like to check (same name as on spreadsheet)')
async def playerrolestats(interaction: discord.Interaction, role: str, player: str):
    row = FindRole(role.lower())
    column = FindPlayer(player.lower())
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
        updateStats(data)

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
    column = GetPlayerFromDiscord(member.name)
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
    
        




def GetPlayerFromDiscord(name: str) -> str:
    match name:
        case "rainbowhead":
            return "I"
        case "slane3470":
            return "I"
        case ".celari":
            return "N"
        case "dadude":
            return "S"
        case "draconic_lord":
            return "X"
        case "thereligionofpeanut":
            return "AC"
        case "orourkustortoise":
            return "AH"
        case "toomai1970":
            return "AM"
        case "lazyvult":
            return "AR"
        case "antinium1312.":
            return "AW"
        case "brotatornator666":
            return "BB"
        case "ianoid":
            return "BG"
        case "roobinski":
            return "BL"
        case "bossors":
            return "BQ"
        case "tsarplatinum":
            return "BV"
        case "ordainedtick266":
            return "CA"
        case "hamsternaut":
            return "CF"
        case ".defize":
            return "CK"
        case "._._neon_._.":
            return "CP"
        case "vandyss":
            return "CU"
        case "brob5046":
            return "CZ"
        case "benign_skies":
            return "DE"
        case "cgotnr":
            return "DJ"
        case "seventyseven_77":
            return "DO"
        case "trees17":
            return "DT"
        case "jetotavio":
            return "DY"
        case "livburrowss":
            return "ED"
        case "sincerenumber82":
            return "EI"
        case "withasideofsalt":
            return "EN"
        case "thorijus":
            return "ES"
        case _: #Error if not found player
            return "ERROR"

        
def FindPlayer(player: str) -> chr:
    match player:
        case "oliver":
            return "I"
        case "sophie":
            return "N"
        case "ethan":
            return "S"
        case "richard":
            return "X"
        case "dan":
            return "AC"
        case "callum":
            return "AH"
        case "ryan":
            return "AM"
        case "stuart":
            return "AR" 
        case "ant":
            return "AW"
        case "ronnie":
            return "BB"
        case "ian":
            return "BG"
        case "reuben":
            return "BL"
        case "findlay":
            return "BQ"
        case "daniel":
            return "BV"
        case "emily":
            return "CA"
        case "laurence":
            return "CF"
        case "benn":
            return "CK"
        case "bens":
            return "CP"
        case "andy":
            return "CU"
        case "ben3":
            return "CZ"
        case "carrick":
            return "DE"
        case "connor":
            return "DJ"
        case "drystan":
            return "DO"
        case "heather":
            return "DT"
        case "etienne":
            return "DY"
        case "liv":
            return "ED"
        case "rory":
            return "EI"
        case "scott":
            return "EN"
        case "tj":
            return "ES"
        case _: #Error if not found player
            return "ERROR"
        
def FindRole(Role: str) -> int:
    match Role:
        #Townsfolk
        case "acrobat":
            return 5
        case "alchemist":
            return 6
        case "alsaahir":
            return 7
        case "amnesiac":
            return 8
        case "artist":
            return 9
        case "atheist":
            return 10
        case "ballonist":
            return 11
        case "banshee":
            return 12
        case "bounty hunter":
            return 13
        case "cannibal":
            return 14
        case "chambermaid":
            return 15
        case "chef":
            return 16
        case "choirboy":
            return 17
        case "clockmaker":
            return 18
        case "courtier":
            return 19
        case "cult Leader":
            return 20
        case "dreamer":
            return 21
        case "empath":
            return 22
        case "engineer":
            return 23
        case "exorcist":
            return 24
        case "farmer":
            return 25
        case "fisherman":
            return 26
        case "flowergirl":
            return 27
        case "fool":
            return 28
        case "fortune teller":
            return 29
        case "gambler":
            return 30
        case "general":
            return 31
        case "gossip":
            return 32
        case "grandmother":
            return 33
        case "high priestess":
            return 34
        case "huntsman":
            return 35
        case "innkeeper":
            return 36
        case "investigator":
            return 37
        case "juggler":
            return 38
        case "king":
            return 39
        case "knight":
            return 40
        case "librarian":
            return 41
        case "lycanthrope":
            return 42
        case "magician":
            return 43
        case "mathmetician":
            return 44
        case "mayor":
            return 45
        case "minstrel":
            return 46
        case "monk":
            return 47
        case "nightwatchman":
            return 48
        case "noble":
            return 49
        case "oracle":
            return 50
        case "pacifist":
            return 51
        case "philosopher":
            return 52
        case "pixie":
            return 53
        case "poppy grower":
            return 54
        case "preacher":
            return 55
        case "professor":
            return 56
        case "ravenkeeper":
            return 57
        case "sage":
            return 58
        case "sailor":
            return 59
        case "savant":
            return 60
        case "seamstress":
            return 61
        case "shugenja":
            return 62
        case "slayer":
            return 63
        case "snake Charmer":
            return 64
        case "soldier":
            return 65
        case "steward":
            return 66
        case "tea lady":
            return 67
        case "town crier":
            return 68
        case "undertaker":
            return 69
        case "village idiot":
            return 70
        case "virgin":
            return 71
        case "washerwoman":
            return 72
        case "townsfolk":
            return 73
        #Outsiders
        case "barber":
            return 77
        case "butler":
            return 78
        case "damsel":
            return 79
        case "drunk":
            return 80
        case "golem":
            return 81
        case "goon":
            return 82
        case "hatter":
            return 83
        case "heretic":
            return 84
        case "klutz":
            return 85
        case "lunatic":
            return 86
        case "moonchild":
            return 87
        case "mutant":
            return 88
        case "ogre":
            return 89
        case "plague doctor":
            return 90
        case "politician":
            return 91
        case "puzzlemaster":
            return 92
        case "recluse":
            return 93
        case "saint":
            return 94
        case "snitch":
            return 95
        case "sweetheart":
            return 96
        case "tinker":
            return 97
        case "zealot":
            return 98
        case "outsider":
            return 99
        #Minions
        case "assassin":
            return 106
        case "baron":
            return 107
        case "boffin":
            return 108
        case "boomdandy":
            return 109
        case "cerenovus":
            return 110
        case "devil's advocate":
            return 111
        case "evil twin":
            return 112
        case "fearmonger":
            return 113
        case "goblin":
            return 114
        case "godfather":
            return 115
        case "harpy":
            return 116
        case "marionette":
            return 117
        case "mastermind":
            return 118
        case "mezepheles":
            return 119
        case "organ grinder":
            return 120
        case "pit-hag":
            return 121
        case "poisoner":
            return 122
        case "psychopath":
            return 123
        case "scarlet woman":
            return 124
        case "spy":
            return 125
        case "summoner":
            return 126
        case "vizier":
            return 127
        case "widow":
            return 128
        case "witch":
            return 129
        case "minion":
            return 130
        #Demons
        case "al-hadikhia":
            return 134
        case "fang gu":
            return 135
        case "imp":
            return 136
        case "kazali":
            return 137
        case "legion":
            return 138
        case "leviathan":
            return 139
        case "lil' monsta":
            return 140
        case "lleech":
            return 141
        case "lord of typhon":
            return 142
        case "no dashii":
            return 143
        case "ojo":
            return 144
        case "po":
            return 145
        case "pukka":
            return 146
        case "riot":
            return 147
        case "shabaloth":
            return 148
        case "vigormortis":
            return 149
        case "vortox":
            return 150
        case "yaggababble":
            return 151
        case "zombuul":
            return 152
        case "demon":
            return 153
        #Totals
        case "total good":
            return 102
        case "total evil":
            return 156
        case "total":
            return 157
        #Error case if not found
        case _:
            return 0
        
        
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
        FixedPlayer.append(FindPlayer(player))
    return FixedPlayer

def replaceRoleArray(Role: array) ->array:
    FixedRole = []
    for entry in Role:
        role = entry[0]
        FixedRole.append(FindRole(role))
    return FixedRole

def updateGoodStat(column: int, row: int, good_win: bool) -> None:
    if(not(good_win)): #since good has not won increment lost instead of win
        column = incrementcol(column)
    cell = str(column) + str(row)
    value = sheet[cell].value
    value += 1
    sheet[cell].value = value
    workbook.save()
    workbook.save(r'C:\Users\rainb\Documents\code\BotC-GroupTracker\BotC-Stats.xlsx')
    

def updateEvilStat(column: int, row: int, good_win: bool) -> None:
    if(good_win): #since good has won evil has not and thus must increment lost instead
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
    column = replacePlayerArray(Player)
    row = replaceRoleArray(Role)
    for entry in row:
        if(row[i] >= 106):
            updateEvilStat(column[i],row[i],good_win)
        else:
            updateGoodStat(column[i],row[i],good_win)
        i += 1
    Total = sheet['H5'].value

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
    while lst:
        carry, next_ = decrement_char(lst.pop())
        result.append(next_)
        if not carry:
            break
        if not lst:
            result.pop()  #since when AA is decremented it goes to Z so should pop i think
    result += lst[::-1]
    return ''.join(result[::-1])

def decrement_char(char: chr):
        if char in ('A'):
            return 0, 'Z'
        else:
            return 1, chr(ord(char) - 1)


bot.run('MTMwMzcxOTAwOTc5NTA0NzUyNg.GWHdIn.4qN6deiWx2QhX6rsC-YBQlPTeAjOKcMK90dbqM')