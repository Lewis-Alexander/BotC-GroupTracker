***Explanation***

This is a discord bot that allows a group that runs Blood on the Clocktower session on the app or in person to track their stats and matchups between players.

***Set-Up***

  - Clone the github by clicking on code then copy either https or ssh then in cmd cd into the desired location and run
```git clone (paste in whichever you copied)```
  - Create a new discord bot to get the token required following this guide:
```https://www.writebots.com/discord-bot-token/```
  - Create a new file in the base directory called token.py and copy this example then replace the temp text ensure you have developer mode on discord then right click on the server/role/channel and select copy id:
```
  from os import path
  token = 'copytokenhere'
  path = r'pathtoBotCStats.xls'
  bot_channel_id = channelforbotspaminserver
  server_id = serveryouwanttoaddbotinto
  role_0_id = roleinserverfor0games
  role_10_id = roleinserverfor10games
  role_20_id = roleinserverfor20games
  role_40_id = roleinserverfor40games
  role_60_id = roleinserverfor60games
  role_80_id = roleinserverfor80games
  role_100_id = roleinserverfor100games
```
  - Open BotcStats.xls then replace all but one of the template sections with your players make sure to paste in the players usernames below their names in the spreadsheet, when you need to add a new player copy the remaining template then insert before that template column you will also need to repeat for the template column in the matchup section (By default i assume that you will move the empty space for self to self matchups if you dont want to bother with that fill in the empty slot in the template). You will also need to check that the total column has properly continued the parttern (it should but excel doesnt like being consistent sometimes).
  - Save the excel sheet then close it

***Usage Explanation***

Run Bot.py, the bot should specify when it is ready in the channel you specified in token.py.
The bot will also open excel it needs to be open for the code to work.
Type / then the options should show in the menu if not run !sync and ctrl r to force refresh the discord app.
When you want to push the spreadsheet make sure to close it as otherwise you will be pushing a temporary copy.

**Add games played**
  - When you finish a game fill in results.csv with the winning team (1 for good wins 0 for evil wins) the players and the roles (ensure to place these in quotation marks) then save it.
  - run /update_spreadsheet it should show a bot thinking, then show update successful Finally it will prompt you whether to save the results in an already existing Session directory or whether to create a new one.
  - you can check whether everything was updated succesfully in the excel sheet.
  - upload_all_session_csvs (allows a user to see all games in a particular session)

**Other owner exclusive commands**
  - sync (mainly for development but if commands dont turn up try !sync)
  - update matchups (uses results.csv but only updates the matchups section of the spreadsheet)
  - update_user_role (allows you to force set someones role to the correct for games played)
  - copy_results (copies results.csv exactly the same as is done in update_spreadsheet but this allows for just adding in games you dont want to track for the overall stats such as teensyville)

**Generic commands**
  - Player to player stats (matchup good,evil,total and delta) uses the matchup section.
  - role stats (player_role_stats, role_total_stats, personal_role_stats) uses the stats section of the spreadsheet
  - average winrate (personal_average, player_average)
  - get_role (displays typed roles rule text)
  - upload_spreadsheet (uploads the xls for users to read)
  - update_role (updates own role by using username)

**Pairwise selection**
  - if you would like the roles to show an image along with the rules text just add png's to a new folder called role-images with the same name as the txt file in role-rules.
  - run either /role_fun_comparison or /role_strength_comparison.
  - the selections will then be saved in their own respective csv files.
  - the rankings can then be shown by using /role_ranking


 
