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


class MyCog(commands.Cog):
  def __init__(self, bot: commands.Bot) -> None:
    self.bot = bot

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
    

@bot.tree.command(name="personalaverage")
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
        case "Oliver":
            return "I"
        case "Sophie":
            return "N"
        case "Ethan":
            return "S"
        case "Richard":
            return "X"
        case "Dan":
            return "AC"
        case "Callum":
            return "AH"
        case "Ryan":
            return "AM"
        case "Stuart":
            return "AR" 
        case "Ant":
            return "AW"
        case "Ronnie":
            return "BB"
        case "Ian":
            return "BG"
        case "Reuben":
            return "BL"
        case "Findlay":
            return "BQ"
        case "Daniel":
            return "BV"
        case "Emily":
            return "CA"
        case "Laurence":
            return "CF"
        case "BenN":
            return "CK"
        case "BenS":
            return "CP"
        case "Andy":
            return "CU"
        case "Ben3":
            return "CZ"
        case "Carrick":
            return "DE"
        case "Connor":
            return "DJ"
        case "Drystan":
            return "DO"
        case "Heather":
            return "DT"
        case "Etienne":
            return "DY"
        case "Liv":
            return "ED"
        case "Rory":
            return "EI"
        case "Scott":
            return "EN"
        case "TJ":
            return "ES"
        case _: #Error if not found player
            return "ERROR"
        
def FindRole(Role: str) -> int:
    match Role:
        #Townsfolk
        case "Acrobat":
            return 5
        case "Alchemist":
            return 6
        case "Alsaahir":
            return 7
        case "Amnesiac":
            return 8
        case "Artist":
            return 9
        case "Atheist":
            return 10
        case "Ballonist":
            return 11
        case "Banshee":
            return 12
        case "Bounty Hunter":
            return 13
        case "Cannibal":
            return 14
        case "Chambermaid":
            return 15
        case "Chef":
            return 16
        case "Choirboy":
            return 17
        case "Clockmaker":
            return 18
        case "Courtier":
            return 19
        case "Cult Leader":
            return 20
        case "Dreamer":
            return 21
        case "Empath":
            return 22
        case "Engineer":
            return 23
        case "Exorcist":
            return 24
        case "Farmer":
            return 25
        case "Fisherman":
            return 26
        case "Flowergirl":
            return 27
        case "Fool":
            return 28
        case "Fortune Teller":
            return 29
        case "Gambler":
            return 30
        case "General":
            return 31
        case "Gossip":
            return 32
        case "Grandmother":
            return 33
        case "High Priestess":
            return 34
        case "Huntsman":
            return 35
        case "Innkeeper":
            return 36
        case "Investigator":
            return 37
        case "Juggler":
            return 38
        case "King":
            return 39
        case "Knight":
            return 40
        case "Librarian":
            return 41
        case "Lycanthrope":
            return 42
        case "Magician":
            return 43
        case "Mathmetician":
            return 44
        case "Mayor":
            return 45
        case "Minstrel":
            return 46
        case "Monk":
            return 47
        case "Nightwatchman":
            return 48
        case "Noble":
            return 49
        case "Oracle":
            return 50
        case "Pacifist":
            return 51
        case "Philosopher":
            return 52
        case "Pixie":
            return 53
        case "Poppy Grower":
            return 54
        case "Preacher":
            return 55
        case "Professor":
            return 56
        case "Ravenkeeper":
            return 57
        case "Sage":
            return 58
        case "Sailor":
            return 59
        case "Savant":
            return 60
        case "Seamstress":
            return 61
        case "Shugenja":
            return 62
        case "Slayer":
            return 63
        case "Snake Charmer":
            return 64
        case "Soldier":
            return 65
        case "Steward":
            return 66
        case "Tea Lady":
            return 67
        case "Town Crier":
            return 68
        case "Undertaker":
            return 69
        case "Village Idiot":
            return 70
        case "Virgin":
            return 71
        case "Washerwoman":
            return 72
        case "Townsfolk":
            return 73
        #Outsiders
        case "Barber":
            return 77
        case "Butler":
            return 78
        case "Damsel":
            return 79
        case "Drunk":
            return 80
        case "Golem":
            return 81
        case "Goon":
            return 82
        case "Hatter":
            return 83
        case "Heretic":
            return 84
        case "Klutz":
            return 85
        case "Lunatic":
            return 86
        case "Moonchild":
            return 87
        case "Mutant":
            return 88
        case "Ogre":
            return 89
        case "Plague Doctor":
            return 90
        case "Politician":
            return 91
        case "Puzzlemaster":
            return 92
        case "Recluse":
            return 93
        case "Saint":
            return 94
        case "Snitch":
            return 95
        case "Sweetheart":
            return 96
        case "Tinker":
            return 97
        case "Zealot":
            return 98
        case "Outsider":
            return 99
        #Minions
        case "Assassin":
            return 106
        case "Baron":
            return 107
        case "Boffin":
            return 108
        case "Boomdandy":
            return 109
        case "Cerenovus":
            return 110
        case "Devil's Advocate":
            return 111
        case "Evil Twin":
            return 112
        case "Fearmonger":
            return 113
        case "Goblin":
            return 114
        case "Godfather":
            return 115
        case "Harpy":
            return 116
        case "Marionette":
            return 117
        case "Mastermind":
            return 118
        case "Mezepheles":
            return 119
        case "Organ Grinder":
            return 120
        case "Pit-Hag":
            return 121
        case "Poisoner":
            return 122
        case "Psychopath":
            return 123
        case "Scarlet Woman":
            return 124
        case "Spy":
            return 125
        case "Summoner":
            return 126
        case "Vizier":
            return 127
        case "Widow":
            return 128
        case "Witch":
            return 129
        case "Minion":
            return 130
        #Demons
        case "Al-Hadikhia":
            return 134
        case "Fang Gu":
            return 134
        case "Imp":
            return 134
        case "Kazali":
            return 134
        case "Legion":
            return 134
        case "Leviathan":
            return 134
        case "Lil'Monsta":
            return 134
        case "Lleech":
            return 134
        case "Lord of Typhon":
            return 134
        case "No Dashii":
            return 134
        case "Ojo":
            return 134
        case "Po":
            return 134
        case "Pukka":
            return 134
        case "Riot":
            return 134
        case "Shabaloth":
            return 134
        case "Vigormortis":
            return 134
        case "Vortox":
            return 134
        case "Yaggababble":
            return 134
        case "Zombuul":
            return 134
        case "Demon":
            return 135
        #Totals
        case "Total Good":
            return 102
        case "Total Evil":
            return 156
        case "Total":
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