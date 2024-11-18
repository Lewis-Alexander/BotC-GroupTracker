def GetPlayerFromDiscord(name: str) -> str:
    match name:
        case "rainbowhead":
            return "AH"
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
        case "total":
            return "D"
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
        case "homebrew townsfolk":
            return 73
        case "townsfolk":
            return 74
        #Outsiders
        case "barber":
            return 78
        case "butler":
            return 79
        case "damsel":
            return 80
        case "drunk":
            return 81
        case "golem":
            return 82
        case "goon":
            return 83
        case "hatter":
            return 84
        case "heretic":
            return 85
        case "klutz":
            return 86
        case "lunatic":
            return 87
        case "moonchild":
            return 88
        case "mutant":
            return 89
        case "ogre":
            return 90
        case "plague doctor":
            return 91
        case "politician":
            return 92
        case "puzzlemaster":
            return 93
        case "recluse":
            return 94
        case "saint":
            return 95
        case "snitch":
            return 96
        case "sweetheart":
            return 97
        case "tinker":
            return 98
        case "zealot":
            return 99
        case "homebrew outsider":
            return 100
        case "outsider":
            return 101
        #Minions
        case "assassin":
            return 108
        case "baron":
            return 109
        case "boffin":
            return 110
        case "boomdandy":
            return 111
        case "cerenovus":
            return 112
        case "devil's advocate":
            return 113
        case "evil twin":
            return 114
        case "fearmonger":
            return 115
        case "goblin":
            return 116
        case "godfather":
            return 117
        case "harpy":
            return 118
        case "marionette":
            return 119
        case "mastermind":
            return 120
        case "mezepheles":
            return 121
        case "organ grinder":
            return 122
        case "pit-hag":
            return 123
        case "poisoner":
            return 124
        case "psychopath":
            return 125
        case "scarlet woman":
            return 126
        case "spy":
            return 127
        case "summoner":
            return 128
        case "vizier":
            return 129
        case "widow":
            return 130
        case "witch":
            return 131
        case "homebrew minion":
            return 132
        case "minion":
            return 133
        #Demons
        case "al-hadikhia":
            return 137
        case "fang gu":
            return 138
        case "imp":
            return 139
        case "kazali":
            return 140
        case "legion":
            return 141
        case "leviathan":
            return 142
        case "lil' monsta":
            return 143
        case "lleech":
            return 144
        case "lord of typhon":
            return 145
        case "no dashii":
            return 146
        case "ojo":
            return 147
        case "po":
            return 148
        case "pukka":
            return 149
        case "riot":
            return 150
        case "shabaloth":
            return 151
        case "vigormortis":
            return 152
        case "vortox":
            return 153
        case "yaggababble":
            return 154
        case "zombuul":
            return 155
        case "homebrew demon":
            return 156
        case "demon":
            return 157
        #Totals
        case "total good":
            return 104
        case "total evil":
            return 160
        case "total":
            return 161
        #Error case if not found
        case _:
            return 0