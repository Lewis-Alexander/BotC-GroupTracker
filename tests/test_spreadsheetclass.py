#Honestly not neccessary to test at all but cant hurt

import pytest
import importlib
from unittest.mock import MagicMock, patch


@pytest.fixture
def fresh_values():
    import Spreadsheetclass
    importlib.reload(Spreadsheetclass)
    return Spreadsheetclass.spreadsheetValues


class TestSpreadsheetValues:
    def test_spreadsheet_values_initialization(self, fresh_values):
        sv = fresh_values
        assert hasattr(sv, 'player_list')
        assert hasattr(sv, 'role_list')
        assert hasattr(sv, 'playercount')
        assert hasattr(sv, 'rolecount')
        from Spreadsheetclass import spreadsheetValuesClass
        assert isinstance(sv, spreadsheetValuesClass)

    def test_spreadsheet_values_default_values(self, fresh_values):
        sv = fresh_values
        # Numeric defaults
        assert sv.average_good == 106
        assert sv.average_evil == 165
        assert sv.average_total == 166
        assert sv.townsfolk == 75
        assert sv.outsider == 103
        assert sv.minion == 138
        assert sv.demon == 163
        assert sv.total_played == 167
        assert sv.playercount == 36
        assert sv.rolecount == 0
        assert sv.matchup_row_start == 171
        assert sv.matchup_gap == 11
        # String defaults
        assert sv.win_column_total == 'D'
        assert sv.starting_player_win_column == 'I'
        assert sv.total_percentage_column == 'F'
        assert sv.starting_player_percentage_column == 'K'
        # List defaults
        assert sv.player_list == []
        assert sv.username_list == []
        assert sv.player_col_list == []
        assert sv.role_list == []
        assert sv.role_list_idx == []

    def test_spreadsheet_values_lists_are_lists(self, fresh_values):
        sv = fresh_values
        assert isinstance(sv.player_list, list)
        assert isinstance(sv.role_list, list)
        assert isinstance(sv.player_col_list, list)
        assert isinstance(sv.username_list, list)
        assert isinstance(sv.role_list_idx, list)

    def test_spreadsheet_values_row_indices_types(self, fresh_values):
        sv = fresh_values
        # These should be integers representing row positions
        assert isinstance(sv.townsfolk, int) or sv.townsfolk is None
        assert isinstance(sv.outsider, int) or sv.outsider is None
        assert isinstance(sv.minion, int) or sv.minion is None
        assert isinstance(sv.demon, int) or sv.demon is None

    def test_setters_numeric_properties(self, fresh_values):
        sv = fresh_values
        sv.average_good = 200
        sv.average_evil = 300
        sv.average_total = 500
        sv.townsfolk = 80
        sv.outsider = 104
        sv.minion = 140
        sv.demon = 170
        sv.total_played = 180
        sv.playercount = 40
        sv.rolecount = 10
        sv.matchup_row_start = 200
        sv.matchup_gap = 15
        assert sv.average_good == 200
        assert sv.average_evil == 300
        assert sv.average_total == 500
        assert sv.townsfolk == 80
        assert sv.outsider == 104
        assert sv.minion == 140
        assert sv.demon == 170
        assert sv.total_played == 180
        assert sv.playercount == 40
        assert sv.rolecount == 10
        assert sv.matchup_row_start == 200
        assert sv.matchup_gap == 15

    def test_setters_string_columns(self, fresh_values):
        sv = fresh_values
        sv.win_column_total = 'Z'
        sv.starting_player_win_column = 'Y'
        sv.total_percentage_column = 'X'
        sv.starting_player_percentage_column = 'W'
        assert sv.win_column_total == 'Z'
        assert sv.starting_player_win_column == 'Y'
        assert sv.total_percentage_column == 'X'
        assert sv.starting_player_percentage_column == 'W'

    def test_setters_list_properties(self, fresh_values):
        sv = fresh_values
        sv.player_list = ['alice', 'bob']
        sv.username_list = ['alice123', 'bob456']
        sv.player_col_list = [2, 3]
        sv.role_list = ['townsfolk', 'demon']
        sv.role_list_idx = [75, 163]
        assert sv.player_list == ['alice', 'bob']
        assert sv.username_list == ['alice123', 'bob456']
        assert sv.player_col_list == [2, 3]
        assert sv.role_list == ['townsfolk', 'demon']
        assert sv.role_list_idx == [75, 163]
