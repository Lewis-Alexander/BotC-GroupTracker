import xlwings as xw
import csv
import array
import discord
from discord import app_commands
from discord.ext import commands
from pathlib import Path
import enum

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
    searchcolumn = GetPlayerFromDiscord(interaction.user.name)
    searchcolumn=chr(ord(searchcolumn)+2)
    cell = str(searchcolumn)+str(102)
    TotalGood = sheet[cell].value
    cell = str(searchcolumn)+str(156)
    TotalEvil = sheet[cell].value
    cell = str(searchcolumn)+str(157)
    Total = sheet[cell].value
    await interaction.response.send_message(f"Your average winrate is {str(Total)} consisting of {str(TotalGood)} whilst good and {str(TotalEvil)} whilst evil!", ephemeral=True)



@bot.tree.command(name="personalrolestats", description="Check any players stats for a particular role")
@app_commands.describe(role = 'Role you would like to check (Townsfolk for Townsfolk total and Total Good for all good)')
async def roleaverage(interaction: discord.Interaction, role: str):
    searchcolumn = GetPlayerFromDiscord(interaction.user.name)
    searchcolumn=chr(ord(searchcolumn)-1)
    row = FindRole(role.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True)
    else:
        data = []
        for i in range(4):       
            cell = str(searchcolumn)+str(row)
            data.append(sheet[cell].value)
            searchcolumn=chr(ord(searchcolumn)+1)
        await interaction.response.send_message(f'you have played the {role} {data[0]} times of those you won {data[1]} and lost {data[2]} which makes your winrate {data[3]}.', ephemeral=True)

@bot.tree.command(name="playeraverage", description="Check any players average winrates")
@app_commands.describe(player = 'Player you would like to see the averages for')
async def playeraverage(interaction: discord.Interaction, player: str):
    column = FindPlayer(player.lower())
    if(column == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column=chr(ord(column)+2)
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
    else: 
        searchcolumn = 'C'
        data = []
        for i in range(4):       
                cell = str(searchcolumn)+str(row)
                data.append(sheet[cell].value)
                searchcolumn=chr(ord(searchcolumn)+1)
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
                column=chr(ord(column)+1)
        await interaction.response.send_message(f'{player} has played {role} {data[0]} times of those they won {data[1]} and lost {data[2]} which makes their winrate {data[3]}.', ephemeral=True)   
    
        

@bot.tree.command(name="uploaddatabase", description="Uploads the current entire database to be checked")
async def uploaddatabase(interaction: discord.Interaction):
    await interaction.response.send_message(file=discord.File(r'BotC-Stats.xlsx'))

@commands.is_owner()  # Prevent other people from using the command
@bot.tree.command(name="updatespreadsheet", description="if program has a csv file it uses it to update the spreadsheet")
async def updatespreadsheet(ctx)-> None:
    my_file = Path("Results.csv")
    if my_file.is_file():
        data = separateFile()
        updateStats(data)


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
            return 134
        case "imp":
            return 134
        case "kazali":
            return 134
        case "legion":
            return 134
        case "leviathan":
            return 134
        case "lil' monsta":
            return 134
        case "lleech":
            return 134
        case "lord of typhon":
            return 134
        case "no dashii":
            return 134
        case "ojo":
            return 134
        case "po":
            return 134
        case "pukka":
            return 134
        case "riot":
            return 134
        case "shabaloth":
            return 134
        case "vigormortis":
            return 134
        case "vortox":
            return 134
        case "yaggababble":
            return 134
        case "zombuul":
            return 134
        case "demon":
            return 135
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

def updateGoodStat(searchcolumn: int, searchrow: int, good_win: bool) -> None:
    if(not(good_win)): #since good has not won increment lost instead of win
        searchcolumn=chr(ord(searchcolumn)+1)
    cell = str(searchcolumn) + str(searchrow)
    value = sheet[cell].value
    value += 1
    sheet[cell].value = value
    workbook.save()
    workbook.save(r'C:\Users\rainb\Documents\code\BotC-GroupTracker\BotC-Stats.xlsx')
    

def updateEvilStat(searchcolumn: int, searchrow: int, good_win: bool) -> None:
    if(good_win): #since good has won evil has not and thus must increment lost instead
        searchcolumn=chr(ord(searchcolumn)+1)
    cell = str(searchcolumn) + str(searchrow)
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
    searchcolumn = replacePlayerArray(Player)
    searchrow = replaceRoleArray(Role)
    for entry in searchrow:
        if(searchrow[i] >= 106):
            updateEvilStat(searchcolumn[i],searchrow[i],good_win)
        else:
            updateGoodStat(searchcolumn[i],searchrow[i],good_win)
        i += 1
    Total = sheet['H5'].value
    print(Total)

bot.run('MTMwMzcxOTAwOTc5NTA0NzUyNg.GWHdIn.4qN6deiWx2QhX6rsC-YBQlPTeAjOKcMK90dbqM')