class spreadsheetValuesClass:
    def __init__(self): #current values from spreadsheet but this allows for it to be changed by the code automatically on startup
        self._average_good = 106
        self._average_evil = 165
        self._average_total = 166
        self._townsfolk = 75
        self._outsider = 103
        self._minion = 138
        self._demon = 163
        self._total_played = 167
        self._win_column_total = 'D' #should never change so no need to include in setters for now
        self._starting_player_win_column = 'I' #should never change so no need to include in setters for now
        self._total_percentage_column = 'F' #should never change so no need to include in setters for now
        self._starting_player_percentage_column = 'K' #should never change so no need to include in setters for now
        self._playercount = 36
        self._rolecount = 0 #dont actually know current number of roles, set during setup
        self._matchup_row_start = 171
        self._matchup_gap = 11
        self._player_list = []
        self._player_col_list = []
        self._role_list = []
        self._role_list_idx = []
        

    @property
    def average_good(self):
        return self._average_good

    @average_good.setter
    def average_good(self, value):
        self._average_good = value

    @property
    def average_evil(self):
        return self._average_evil

    @average_evil.setter
    def average_evil(self, value):
        self._average_evil = value

    @property
    def average_total(self):
        return self._average_total

    @average_total.setter
    def average_total(self, value):
        self._average_total = value

    @property
    def townsfolk(self):
        return self._townsfolk

    @townsfolk.setter
    def townsfolk(self, value):
        self._townsfolk = value

    @property
    def outsider(self):
        return self._outsider

    @outsider.setter
    def outsider(self, value):
        self._outsider = value

    @property
    def minion(self):
        return self._minion

    @minion.setter
    def minion(self, value):
        self._minion = value

    @property
    def demon(self):
        return self._demon

    @demon.setter
    def demon(self, value):
        self._demon = value

    @property
    def total_played(self):
        return self._total_played

    @total_played.setter
    def total_played(self, value):
        self._total_played = value

    @property
    def win_column_total(self):
        return self._win_column_total

    @win_column_total.setter
    def win_column_total(self, value):
        self._win_column_total = value

    @property
    def starting_player_win_column(self):
        return self._starting_player_win_column

    @starting_player_win_column.setter
    def starting_player_win_column(self, value):
        self._starting_player_win_column = value

    @property
    def total_percentage_column(self):
        return self._total_percentage_column

    @total_percentage_column.setter
    def total_percentage_column(self, value):
        self._total_percentage_column = value

    @property
    def starting_player_percentage_column(self):
        return self._starting_player_percentage_column

    @starting_player_percentage_column.setter
    def starting_player_percentage_column(self, value):
        self._starting_player_percentage_column = value

    @property
    def playercount(self):
        return self._playercount

    @playercount.setter
    def playercount(self, value):
        self._playercount = value
        
    @property
    def rolecount(self):
        return self._rolecount
    
    @rolecount.setter
    def rolecount(self, value):
        self._rolecount = value

    @property
    def matchup_row_start(self):
        return self._matchup_row_start

    @matchup_row_start.setter
    def matchup_row_start(self, value):
        self._matchup_row_start = value

    @property
    def matchup_gap(self):
        return self._matchup_gap

    @matchup_gap.setter
    def matchup_gap(self, value):
        self._matchup_gap = value
        
    @property
    def player_list(self):
        return self._player_list
    
    @player_list.setter
    def player_list(self, value):
        self._player_list = value
        
    @property
    def player_col_list(self):
        return self._player_col_list
    
    @player_col_list.setter
    def player_col_list(self, value):
        self._player_col_list = value
        
    @property
    def role_list(self):
        return self._role_list
    
    @role_list.setter
    def role_list(self, value):
        self._role_list = value
        
    @property
    def role_list_idx(self):
        return self._role_list_idx
    
    @role_list_idx.setter
    def role_list_idx(self, value):
        self._role_list_idx = value

spreadsheetValues = spreadsheetValuesClass()
# role_stats.average_good = 110  # Setting a value
# print(role_stats.average_good)  # Getting a value
