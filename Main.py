import openpyxl 
workbook = openpyxl.load_workbook('BotC-Stats.xlsx')
sheet = workbook.active
def main() -> None: 
    return None

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
            return "Ant"
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
        #Error case if not found
        case _:
            return 0
        
        

        

def updateGoodStat(column: chr, row: int, good_win: bool) -> None:
    if(not(good_win)): #since good has not won increment lost instead of win
        column=chr(ord(column)+1)
    cell = str(column) + str(row)
    incremented_data = sheet.cell().value + 1
    sheet[cell] = incremented_data

def updateEvilStat(column: chr, row: int, good_win: bool) -> None:
    if(good_win): #since good has won evil has not and thus must increment lost instead
        column=chr(ord(column)+1)
    cell = str(column) + str(row)
    incremented_data = sheet.cell().value + 1
    sheet[cell] = incremented_data


def saveAndClose() -> None:
    workbook.save('BotC-Stats.xlsx')
    
