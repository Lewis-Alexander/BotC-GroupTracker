import discord
import Helper
import Pairwise
import random
from Token import token
from discord import app_commands
from discord.ui import View, Button
from discord.ext import commands
from pathlib import Path
from Spreadsheetclass import spreadsheetValues

bot = commands.Bot(command_prefix='!', intents = discord.Intents.all())

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
    

@bot.tree.command(name="personal_average", description="Check your average winrates for sides")
async def personal_average(interaction: discord.Interaction):
    column = Helper.find_player_username(interaction.user.name)
    column = Helper.increment_col(column)
    column = Helper.increment_col(column)
    cell = str(column)+str(spreadsheetValues.average_good)
    TotalGood = Helper.sheet[cell].value
    cell = str(column)+str(spreadsheetValues.average_evil)
    TotalEvil = Helper.sheet[cell].value
    cell = str(column)+str(spreadsheetValues.average_total)
    Total = Helper.sheet[cell].value
    cell = str(column)+str(spreadsheetValues.total_played)
    Played = Helper.sheet[cell].value
    await interaction.response.send_message(f"Over {str(Played)} games played your average winrate was {str(Total)} consisting of {str(TotalGood)} whilst good and {str(TotalEvil)} whilst evil!", ephemeral=True)



@bot.tree.command(name="personal_role_stats", description="Check any your stats for a particular role")
@app_commands.describe(role = 'Role you would like to check (Townsfolk for Townsfolk total and Total Good for all good)')
async def personal_role_stats(interaction: discord.Interaction, role: str):
    column = Helper.find_player_username(interaction.user.name)
    column = Helper.decrement_col(column)
    row = Helper.find_role(role.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True)
    else:
        data = []
        for i in range(4):     
            cell = str(column)+str(row)
            data.append(Helper.sheet[cell].value)
            column = Helper.increment_col(column)
        await interaction.response.send_message(f'you have played the {role} {data[0]} times, of those you won {data[1]} and lost {data[2]} which makes your winrate {data[3]}.')

@bot.tree.command(name="player_average", description="Check any players average winrates")
@app_commands.describe(player = 'Player you would like to see the averages for')
async def player_average(interaction: discord.Interaction, player: str):
    column = Helper.find_player(player.lower())
    if(column == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column = Helper.increment_col(column)
        column = Helper.increment_col(column)
        data = []
        rows = [spreadsheetValues.average_good,spreadsheetValues.average_evil,spreadsheetValues.average_total]
        for i in range(3):       
            cell = str(column)+str(rows[i])
            data.append(Helper.sheet[cell].value)
        await interaction.response.send_message(f'{player}\'s average winrates are as follows: Good-{data[0]} Evil-{data[1]} Total-{data[2]}')


@bot.tree.command(name="role_total_stats", description="Check the stats for a particular role")
@app_commands.describe(role = 'Role you would like to check, (Townsfolk for Townsfolk total and Total Good for all good)')
async def role_total_stats(interaction: discord.Interaction, role: str):
    row = Helper.find_role(role.lower())
    if(row == "ERROR"):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True) 
    else: 
        column = 'C'
        data = []
        for i in range(4):       
                cell = str(column)+str(row)
                data.append(Helper.sheet[cell].value)
                column = Helper.increment_col(column)
        await interaction.response.send_message(f'The {role} has been played {data[0]} times of those they won {data[1]} and lost {data[2]} which makes their winrate {data[3]}.')

@bot.tree.command(name='player_role_stats', description="Check any players stats for a particular role")
@app_commands.describe(role = 'Role you would like to check, (Townsfolk for Townsfolk total and Total Good for all good)')
@app_commands.describe(player = 'Player you would like to check (same name as on spreadsheet)')
async def player_role_stats(interaction: discord.Interaction, role: str, player: str):
    row = Helper.find_role(role.lower())
    column = Helper.find_player(player.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True)
    elif(column == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        data = []
        for i in range(4):       
                cell = str(column)+str(row)
                data.append(Helper.sheet[cell].value)
                column = Helper.increment_col(column)
        await interaction.response.send_message(f'{player} has played {role} {data[0]} times of those they won {data[1]} and lost {data[2]} which makes their winrate {data[3]}.', ephemeral=True)   
    
        

@bot.tree.command(name="upload_spreadsheet", description="Uploads the current entire spreadsheet to be perused at your pleasure")
async def upload_spreadsheet(interaction: discord.Interaction):
    await interaction.response.send_message(file=discord.File(r'BotC-Stats.xlsx'))

@commands.is_owner()  # Prevent other people from using the command
@bot.tree.command(name="update_spreadsheet", description="if program has a csv file it uses it to update the spreadsheet")
async def update_spreadsheet(interaction: discord.Interaction):
    await interaction.response.defer(ephemeral=True)
    my_file = Path("Results.csv")
    if my_file.is_file():
        data = Helper.separate_file()
        Helper.update_stats(data)
        await interaction.followup.send(f'spreadsheet updated', ephemeral=True)
        
@commands.is_owner()  # Prevent other people from using the command
@bot.tree.command(name="update_matchups", description="if program has a csv file it uses it to update the spreadsheet")
async def update_matchups(interaction: discord.Interaction):
    await interaction.response.defer(ephemeral=True)
    my_file = Path("Results.csv")
    if my_file.is_file():
        data = Helper.separate_file()
        Helper.update_matchups(data)
        await interaction.followup.send(f'matchups updated', ephemeral=True)      

@bot.tree.command(name="update_role", description="auto update personal role from database")
async def update_role(interaction : discord.Interaction):
    server = bot.get_guild(1192193425739612281)
    role0 = server.get_role(1294069983151919134)
    role10 = server.get_role(1451683827897597973)
    role20 = server.get_role(1294281073043177535)
    role40 = server.get_role(1294068836764614788)
    role60 = server.get_role(1294069566409801760)
    role80 = server.get_role(1294068801448574996)
    role100 = server.get_role(1294068667801276426)
    member = interaction.user
    column = Helper.find_player_username(member.name)
    column = Helper.increment_col(column)
    column = Helper.increment_col(column)
    cell = column + str(spreadsheetValues.total_played)
    cellval = Helper.sheet[cell].value
    if(cellval < 10):
        await member.add_roles(role0)
        await interaction.response.send_message(f'You have been assigned the 0+ games role as you have not yet attended enough to increase you currently have {cellval} games played.', ephemeral=True)
    elif(cellval < 20):
        await member.add_roles(role10)
        await member.remove_roles(role0)
        await interaction.response.send_message(f'You have been assigned the 10+ games role as you have attended enough to increase you currently have {cellval} games played.', ephemeral=True)
    elif(cellval < 40):
        await member.add_roles(role20)
        await member.remove_roles(role10)
        await interaction.response.send_message(f'You have been assigned the 20+ games role as you have attended enough to increase you currently have {cellval} games played.', ephemeral=True)
    elif(cellval < 60):
        await member.add_roles(role40)
        await member.remove_roles(role20)
        await interaction.response.send_message(f'You have been assigned the 40+ games role as you have attended enough to increase you currently have {cellval} games played. ', ephemeral=True)
    elif(cellval < 80):
        await member.add_roles(role60)
        await member.remove_roles(role40)
        await interaction.response.send_message(f'You have been assigned the 60+ games role as you have attended enough to increase you currently have {cellval} games played.', ephemeral=True)
    elif(cellval < 100):
        await member.add_roles(role80)
        await member.remove_roles(role60)
        await interaction.response.send_message(f'You have been assigned the 80+ games role as you have attended enough to increase you currently have {cellval} games played.', ephemeral=True)
    else:
        await member.add_roles(role100)
        await member.remove_roles(role80)
        await interaction.response.send_message(f'You have been assigned the 100+ games role as you have attended enough to increase you currently have {cellval} games played.', ephemeral=True)
        
@bot.tree.command(name="highest_role_winrate", description="highest winrate for each role")
@app_commands.describe(role = 'Role you would like to check (Townsfolk for Townsfolk total and Total Good for all good)')
async def highest_role_winrate(interaction : discord.Interaction, role : str):
    column = spreadsheetValues.starting_player_percentage_column
    row = Helper.find_role(role.lower())
    if(row == 0):
        await interaction.response.send_message(f'Role not found please check spelling', ephemeral=True)
    data = []
    for player in range (spreadsheetValues.playercount):
        cell = str(column)+str(row)
        tempdata = Helper.sheet[cell].value
        tempdata = tempdata[:-1]
        data.append(float(tempdata))
        for i in range(5):
            column = Helper.increment_col(column)
    unsorteddata = data.copy()
    data.sort(reverse=True)
    i = 0
    highestindex = []
    for entry in data:
        if(data[0] == unsorteddata[i]):
            highestindex.append(i)
        i += 1
    name = []
    for value in highestindex:
        column = spreadsheetValues.starting_player_win_column
        for index in range(value):
            for i in range(5):
                column = Helper.increment_col(column)
        cell = str(column) + str(1)
        name.append(Helper.sheet[cell].value)
    if(len(name) > 1):
        players = f'Multiple players have a tied winrate of {data[0]}% on the {role} those being '
        for entry in name:
            if(entry != name[-1]):
                players += f'{entry}, '
            else:
                players += f'and {entry}.'
        await interaction.response.send_message(f'{players}')
    else:
        await interaction.response.send_message(f'{name[0]} has the highest winrate as the/a {role} which is {data[0]}%')
    
@commands.is_owner()  # Prevent other people from using the command
@bot.tree.command(name="update_user_role", description="Update the role for a given user")
@app_commands.describe(user="The user whose role you want to update")
async def update_user_role(interaction: discord.Interaction, user: discord.Member):
    try:
        # Retrieve the column for the user
        column = Helper.find_player_username(user.name)
        column = Helper.increment_col(column)
        column = Helper.increment_col(column)
        cell = column + str(spreadsheetValues.total_played)
        cellval = Helper.sheet[cell].value

        # Define roles based on the algorithm in update_role
        server = interaction.guild
        role0 = server.get_role(1294069983151919134)
        role10 = server.get_role(1451683827897597973)
        role20 = server.get_role(1294281073043177535)
        role40 = server.get_role(1294068836764614788)
        role60 = server.get_role(1294069566409801760)
        role80 = server.get_role(1294068801448574996)
        role100 = server.get_role(1294068667801276426)

        # Assign roles based on the total games played
        if cellval < 10:
            await user.add_roles(role0)
            await interaction.response.send_message(f"{user.mention} has been assigned the 0+ games role. they have {cellval} games played.", ephemeral=True)
        elif cellval < 20:
            await user.add_roles(role10)
            await user.remove_roles(role0)
            await interaction.response.send_message(f"{user.mention} has been assigned the 10+ games role. they have {cellval} games played.", ephemeral=True)
        elif cellval < 40:
            await user.add_roles(role20)
            await user.remove_roles(role10)
            await interaction.response.send_message(f"{user.mention} has been assigned the 20+ games role. they have {cellval} games played.", ephemeral=True)
        elif cellval < 60:
            await user.add_roles(role40)
            await user.remove_roles(role20)
            await interaction.response.send_message(f"{user.mention} has been assigned the 40+ games role. they have {cellval} games played.", ephemeral=True)
        elif cellval < 80:
            await user.add_roles(role60)
            await user.remove_roles(role40)
            await interaction.response.send_message(f"{user.mention} has been assigned the 60+ games role. they have {cellval} games played.", ephemeral=True)
        elif cellval < 100:
            await user.add_roles(role80)
            await user.remove_roles(role60)
            await interaction.response.send_message(f"{user.mention} has been assigned the 80+ games role. they have {cellval} games played.", ephemeral=True)
        else:
            await user.add_roles(role100)
            await user.remove_roles(role80)
            await interaction.response.send_message(f"{user.mention} has been assigned the 100+ games role. they have {cellval} games played.", ephemeral=True)
    except Exception as e:
        # Handle any errors that occur
        await interaction.response.send_message(f"Failed to update role: {str(e)}", ephemeral=True)

@bot.tree.command(name="player_to_player_matchup_evil", description="Check two players matchup stats when on the evil team together first is the reference player")
@app_commands.describe(player1 = 'Player you would like to check (same name as on spreadsheet)')
@app_commands.describe(player2 = 'Player you would like to check (same name as on spreadsheet)')
async def player_to_player_matchup_evil(interaction: discord.Interaction, player1: str, player2: str):
    column = Helper.find_player(player1.lower())
    row = Helper.find_player_matchup(player2.lower())
    if(column == "ERROR" or row == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column = Helper.increment_col(column)
        column = Helper.increment_col(column)
        row = row + 1 #start with minion info
        cell = str(column)+str(row)
        data = []
        for i in range(4):       
                cell = str(column)+str(row)
                data.append(Helper.sheet[cell].value)
                row = row + 1
        await interaction.response.send_message(f'when {player1} and {player2} are both minions their winrate is {data[0]}, if the second player is the demon or they both are demons their winrate is {data[1]}, if the first player is the demon and the second is one of the minions their winrate is {data[2]}, this means their average winrate as an evil team is {data[3]}.')
        
@bot.tree.command(name="player_to_player_matchup_good", description="Check two players matchup stats when on the good team together first is the reference player")
@app_commands.describe(player1 = 'Reference Player you would like to check (same name as on spreadsheet)')
@app_commands.describe(player2 = 'Secondary Player you would like to check (same name as on spreadsheet)')
async def player_to_player_matchup_good(interaction: discord.Interaction, player1: str, player2: str):
    column = Helper.find_player(player1.lower())
    row = Helper.find_player_matchup(player2.lower())
    if(column == "ERROR" or row == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column = Helper.increment_col(column)
        column = Helper.increment_col(column)
        row = row + 5 #start with correct info
        cell = str(column)+str(row)    
        data = Helper.sheet[cell].value
        row = row + 1
        await interaction.response.send_message(f'when {player1} and {player2} are both on the good team their winrate is {data}.')
        
@bot.tree.command(name="player_to_player_matchup_total", description="Check two players matchup stats when on the same team, first is the reference player")
@app_commands.describe(player1 = 'Reference Player you would like to check (same name as on spreadsheet)')
@app_commands.describe(player2 = 'Secondary Player you would like to check (same name as on spreadsheet)')
async def player_to_player_matchup_total(interaction: discord.Interaction, player1: str, player2: str):
    column = Helper.find_player(player1.lower())
    row = Helper.find_player_matchup(player2.lower())
    if(column == "ERROR" or row == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column = Helper.increment_col(column)
        column = Helper.increment_col(column)
        row = row + 6 #start with correct info
        cell = str(column)+str(row)    
        data = Helper.sheet[cell].value
        row = row + 1
        await interaction.response.send_message(f'when {player1} and {player2} are both on the same team their winrate is {data}.')
        

@bot.tree.command(name="player_to_player_winrate_delta", description="Check two players winrate delta when playing together, first is the reference player")
@app_commands.describe(player1 = 'Reference Player you would like to check (same name as on spreadsheet)')
@app_commands.describe(player2 = 'Secondary Player you would like to check (same name as on spreadsheet)')
async def player_to_player_winrate_delta(interaction: discord.Interaction, player1: str, player2: str):
    column = Helper.find_player(player1.lower())
    row = Helper.find_player_matchup(player2.lower())
    if(column == "ERROR" or row == "ERROR"):
        await interaction.response.send_message(f'Player not found please check spelling', ephemeral=True)
    else:
        column = Helper.increment_col(column)
        column = Helper.increment_col(column)
        row = row + 7 #start with correct info
        cell = str(column)+str(row)
        data = []
        for i in range(3):       
                cell = str(column)+str(row)
                data.append(Helper.sheet[cell].value)
                row = row + 1
        data[2] = data[2]*100
        await interaction.response.send_message(f'when {player1} and {player2} are on the evil team together the first players default evil winrate changes by {data[0]} (absolute), if they are together on the good team their winrate changes by {data[1]} (absolute), on average if both players are on the same team the first players average winrate delta is an absolute value of {data[2]}%.')

# reused for both fun and strength comparisons
async def send_new_comparison(interaction: discord.Interaction, is_initial: bool, category: str, message_text: str):
        role1, role2 = random.sample(spreadsheetValues.role_list, 2)

        async def button_callback(interaction: discord.Interaction, selected_role: str):
            Pairwise.save_pairwise_comparison(role1, role2, selected_role, category=category)
            await send_new_comparison(interaction, is_initial=False, category=category, message_text=message_text)

        async def skip_callback(interaction: discord.Interaction):
            await send_new_comparison(interaction, is_initial=False, category=category, message_text=message_text)

        role1_image = Helper.get_role_image(role1)
        role2_image = Helper.get_role_image(role2)

        button1 = Button(label=role1, style=discord.ButtonStyle.primary)
        button2 = Button(label=role2, style=discord.ButtonStyle.primary)
        skip_button = Button(label="Do Not Know", style=discord.ButtonStyle.secondary)

        button1.callback = lambda i: button_callback(i, role1)
        button2.callback = lambda i: button_callback(i, role2)
        skip_button.callback = skip_callback

        view = View()
        view.add_item(button1)
        view.add_item(button2)
        view.add_item(skip_button)

        files = []
        content = message_text
        if role1_image:
            files.append(discord.File(role1_image, filename="role1.png"))
            content += f"\nRole 1: {role1}"
        if role2_image:
            files.append(discord.File(role2_image, filename="role2.png"))
            content += f"\nRole 2: {role2}"

        if is_initial:
            await interaction.response.send_message(
                content=content, view=view, ephemeral=True, files=files
            )
        else:
            await interaction.followup.send(
                content=content, view=view, files=files, ephemeral=True
            )
    
@bot.tree.command(name="role_fun_comparison", description="Randomly select two roles and ask which is more fun to play as/against")
async def role_fun_comparison(interaction: discord.Interaction):
    await send_new_comparison(interaction, is_initial=True, category="fun", message_text="Which role is more fun to play as/against?")
    
@bot.tree.command(name="role_strength_comparison", description="Randomly select two roles and ask which is in your view stronger")
async def role_strength_comparison(interaction: discord.Interaction):
    await send_new_comparison(interaction, is_initial=True, category="strength", message_text="Which role is stronger?")

@bot.tree.command(name="role_fun_ranking", description="Display a ranking of roles based on pairwise comparisons")
async def role_fun_ranking(interaction: discord.Interaction):
    ranked_roles = Pairwise.generate_role_ranking(category="fun")
    if not ranked_roles:
        await interaction.response.send_message("No ranking data available.", ephemeral=True)
        return

    ranking_message = "\n".join(
        [f"{i + 1}. {role} (Score: {score})" for i, (role, score) in enumerate(ranked_roles)]
    )
    
    await interaction.response.send_message(
        f"Role Ranking:\n{ranking_message}", ephemeral=True
    )

@bot.tree.command(name="role_strength_ranking", description="Display a ranking of roles based on pairwise comparisons")
async def role_strength_ranking(interaction: discord.Interaction):
    ranked_roles = Pairwise.generate_role_ranking(category="strength")
    
    if not ranked_roles:
        await interaction.response.send_message("No ranking data available.", ephemeral=True)
        return
    
    ranking_message = "\n".join(
        [f"{i + 1}. {role} (Score: {score})" for i, (role, score) in enumerate(ranked_roles)]
    )
    
    await interaction.response.send_message(
        f"Role Ranking:\n{ranking_message}", ephemeral=True
    )

@bot.event
async def on_ready():
    Helper.setup_class()
    #await bot.get_channel(1302801362517622874).send(f"Bot is running and has loaded all current players and roles from the spreadsheet currently {len(spreadsheetValues.username_list)} players and {spreadsheetValues.rolecount} roles.")
    

bot.run(token)
