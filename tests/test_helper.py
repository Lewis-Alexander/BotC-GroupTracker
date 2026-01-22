import pytest
from unittest.mock import MagicMock, patch, mock_open
from pathlib import Path


class TestFindRole:
    @patch('Helper.spreadsheetValues')
    def test_find_role_townsfolk(self, mock_values):
        mock_values.townsfolk = 2
        mock_values.role_list = []
        
        from Helper import find_role
        result = find_role("townsfolk")
        assert result == 2
    
    @patch('Helper.spreadsheetValues')
    def test_find_role_outsider(self, mock_values):
        mock_values.outsider = 3
        mock_values.role_list = []
        
        from Helper import find_role
        result = find_role("outsider")
        assert result == 3
    
    @patch('Helper.spreadsheetValues')
    def test_find_role_not_found(self, mock_values):
        mock_values.role_list = ['townsfolk', 'outsider']
        mock_values.role_list_idx = [2, 3]
        
        from Helper import find_role
        result = find_role("nonexistent")
        assert result == "ERROR"

    @patch('Helper.spreadsheetValues')
    def test_find_role_specific_role_lookup(self, mock_values):
        mock_values.role_list = ['artist', 'savant']
        mock_values.role_list_idx = [80, 200]
        from Helper import find_role
        assert find_role('ARTIST') == 80
        assert find_role('savant') == 200

    @patch('Helper.spreadsheetValues')
    def test_find_role_totals(self, mock_values):
        mock_values.average_good = 10
        mock_values.average_evil = 20
        mock_values.average_total = 30
        from Helper import find_role
        assert find_role('total good') == 10
        assert find_role('total evil') == 20
        assert find_role('total') == 30


class TestFindPlayer:
    @patch('Helper.spreadsheetValues')
    def test_find_player_exists(self, mock_values):
        mock_values.player_list = ['alice', 'bob', 'charlie']
        mock_values.player_col_list = [2, 3, 4]
        
        from Helper import find_player
        result = find_player("alice")
        assert result == 2
    
    @patch('Helper.spreadsheetValues')
    def test_find_player_case_insensitive(self, mock_values):
        mock_values.player_list = ['alice', 'bob']
        mock_values.player_col_list = [2, 3]
        
        from Helper import find_player
        result = find_player("ALICE")
        assert result == 2
    
    @patch('Helper.spreadsheetValues')
    def test_find_player_not_found(self, mock_values):
        mock_values.player_list = ['alice', 'bob']
        mock_values.player_col_list = [2, 3]
        
        from Helper import find_player
        result = find_player("dave")
        assert result == "ERROR"

class TestFindPlayerUsername:
    @patch('Helper.spreadsheetValues')
    def test_find_player_username_exists(self, mock_values):
        mock_values.username_list = ['alice123', 'bob456']
        mock_values.player_col_list = [2, 3]
        from Helper import find_player_username
        assert find_player_username('alice123') == 2
        assert find_player_username('bob456') == 3

    @patch('Helper.spreadsheetValues')
    def test_find_player_username_not_found(self, mock_values):
        mock_values.username_list = ['alice123']
        mock_values.player_col_list = [2]
        from Helper import find_player_username
        assert find_player_username('charlie789') == 'ERROR'

class TestFindPlayerMatchup:
    @patch('Helper.spreadsheetValues')
    def test_find_player_matchup_row_calc(self, mock_values):
        mock_values.player_list = ['alice', 'bob', 'charlie']
        mock_values.matchup_row_start = 171
        mock_values.matchup_gap = 11
        from Helper import find_player_matchup
        assert find_player_matchup('alice') == 171
        assert find_player_matchup('bob') == 182
        assert find_player_matchup('charlie') == 193
    
    @patch('Helper.spreadsheetValues')
    def test_find_player_matchup_not_found(self, mock_values):
        mock_values.player_list = ['alice']
        mock_values.matchup_row_start = 100
        mock_values.matchup_gap = 10
        from Helper import find_player_matchup
        assert find_player_matchup('bob') == 'ERROR'

class TestFindRoleMatchup:
    @patch('Helper.spreadsheetValues')
    def test_find_role_matchup_categories(self, mock_values):
        mock_values.outsider = 103
        mock_values.minion = 138
        mock_values.demon = 163
        from Helper import find_role_matchup
        assert find_role_matchup(75) == 1  # townsfolk
        assert find_role_matchup(103) == 1  # outsider boundary
        assert find_role_matchup(120) == 2  # minion range
        assert find_role_matchup(163) == 3  # demon boundary
        assert find_role_matchup(200) == 'ERROR'


class TestGetRoleImage:
    def test_get_role_image_exists(self, tmp_path):
        role_images_dir = tmp_path / "role-images"
        role_images_dir.mkdir()
        image_file = role_images_dir / "townsfolk.png"
        image_file.touch()
        
        with patch('Helper.Path') as mock_path:
            def path_side_effect(arg):
                if arg == "role-images":
                    return role_images_dir
                return Path(arg)
            mock_path.side_effect = path_side_effect
            from Helper import get_role_image
            image_path = get_role_image("townsfolk")
            assert image_path is not None
            assert image_path.exists()
            assert image_path.name == "townsfolk.png"
    
    def test_get_role_image_sanitization(self, tmp_path):
        role_images_dir = tmp_path / "role-images"
        role_images_dir.mkdir()
        image_file = role_images_dir / "eviltwin.png"
        image_file.touch()
        
        with patch('Helper.Path') as mock_path:
            def path_side_effect(arg):
                if arg == "role-images":
                    return role_images_dir
                return Path(arg)
            mock_path.side_effect = path_side_effect
            from Helper import get_role_image
            image_path = get_role_image("Evil Twin")
            assert image_path is not None
            assert image_path.exists()
            assert image_path.name == "eviltwin.png"


class TestGetRoleRules:
    def test_get_role_rules_exists(self, tmp_path):
        # Create a mock rules file for "savant"
        rules_dir = tmp_path / "role-rules"
        rules_dir.mkdir()
        rules_file = rules_dir / "savant.txt"
        rules_file.write_text("Savant rules here.")

        with patch('Helper.Path') as mock_path:
            def path_side_effect(arg):
                if arg == "role-rules":
                    return rules_dir
                return Path(arg)
            mock_path.side_effect = path_side_effect
            from Helper import get_role_rules
            rules_path = get_role_rules("savant")
            assert rules_path is not None
            assert rules_path.exists()
            assert rules_path.suffix == ".txt"

    def test_get_role_rules_sanitization(self, tmp_path):
        # Create a mock rules file for "eviltwin"
        rules_dir = tmp_path / "role-rules"
        rules_dir.mkdir()
        rules_file = rules_dir / "eviltwin.txt"
        rules_file.write_text("Evil Twin rules here.")

        with patch('Helper.Path') as mock_path:
            def path_side_effect(arg):
                if arg == "role-rules":
                    return rules_dir
                return Path(arg)
            mock_path.side_effect = path_side_effect
            from Helper import get_role_rules
            rules_path = get_role_rules("Evil Twin")
            assert rules_path is not None
            assert rules_path.exists()


class TestSeparateFile:
    def test_separate_file_reads_results_csv(self, tmp_path):
        csv_content = "1\nalice\ntownsfolk\nbob\nminion\n"
        m = mock_open(read_data=csv_content)
        with patch('Helper.open', m):
            from Helper import separate_file
            result = separate_file()
            assert isinstance(result, list)
            assert result == [['1'], ['alice'], ['townsfolk'], ['bob'], ['minion']]


class TestReplaceArrays:
    @patch('Helper.find_player')
    def test_replace_player_array(self, mock_find_player):
        mock_find_player.side_effect = [2, 3, 4]
        
        from Helper import replace_player_array
        players = [['alice'], ['bob'], ['charlie']]
        result = replace_player_array(players)
        assert result == [2, 3, 4]
    
    @patch('Helper.find_role')
    def test_replace_role_array(self, mock_find_role):
        mock_find_role.side_effect = [2, 3, 4]
        
        from Helper import replace_role_array
        roles = [['townsfolk'], ['minion'], ['demon']]
        result = replace_role_array(roles)
        assert result == [2, 3, 4]

class TestUpdateStatsFunctions:
    def _mock_sheet(self):
        sheet = MagicMock()
        cell_mock = MagicMock()
        cell_mock.value = None
        sheet.__getitem__.return_value = cell_mock
        return sheet, cell_mock

    def test_update_good_stat_increments_win_cell(self):
        import Helper
        sheet, cell = self._mock_sheet()
        Helper.sheet_edit = sheet
        Helper.workbook_edit = MagicMock()
        Helper.update_good_stat(column=5, row=10, good_win=['1'])
        # Good won => no column increment
        assert sheet.__getitem__.called
        assert cell.value == 1
        Helper.workbook_edit.save.assert_called_once()

    def test_update_good_stat_increments_loss_cell_when_good_lost(self):
        import Helper
        sheet, cell = self._mock_sheet()
        Helper.sheet_edit = sheet
        Helper.workbook_edit = MagicMock()
        Helper.update_good_stat(column=5, row=10, good_win=['0'])
        # Good lost => column incremented internally
        assert sheet.__getitem__.called
        assert cell.value == 1
        Helper.workbook_edit.save.assert_called_once()

    def test_update_evil_stat_increments_correct_cell(self):
        import Helper
        sheet, cell = self._mock_sheet()
        Helper.sheet_edit = sheet
        Helper.workbook_edit = MagicMock()
        Helper.update_evil_stat(column=6, row=20, good_win=['1'])
        # Good won => evil lost, column increment
        assert sheet.__getitem__.called
        assert cell.value == 1
        Helper.workbook_edit.save.assert_called_once()

    def test_update_player_matchup_same_team_and_evil_update(self):
        import Helper
        sheet = MagicMock()
        first_cell = MagicMock(); first_cell.value = None
        second_cell = MagicMock(); second_cell.value = None
        # Read and write for first and second cell accesses
        sheet.__getitem__.side_effect = [first_cell, first_cell, second_cell, second_cell]
        Helper.sheet_edit = sheet
        Helper.workbook_edit = MagicMock()
        # col_player=2 (minion), row_player=3 (demon) => gap=2, evil=2
        Helper.update_player_matchup(column=5, row=100, col_player=2, row_player=3, won=['1'])
        assert first_cell.value == 1
        assert second_cell.value == 1
        assert Helper.workbook_edit.save.call_count == 2

    def test_update_player_matchup_different_teams_early_return(self):
        import Helper
        Helper.sheet_edit = MagicMock()
        Helper.workbook_edit = MagicMock()
        # Different teams => early return
        Helper.update_player_matchup(column=5, row=100, col_player=1, row_player=2, won=['1'])
        Helper.sheet_edit.__getitem__.assert_not_called()
        Helper.workbook_edit.save.assert_not_called()

    @patch('Helper.replace_player_array')
    @patch('Helper.replace_role_array')
    @patch('Helper.replace_player_array_matchup')
    @patch('Helper.replace_role_array_matchup')
    @patch('Helper.update_good_stat')
    @patch('Helper.update_evil_stat')
    @patch('Helper.update_player_matchup')
    def test_update_stats_flow(self, mock_update_player_matchup, mock_update_evil, mock_update_good,
                               mock_role_matchup, mock_player_matchup, mock_replace_role, mock_replace_player):
        from Helper import update_stats
        Data = [['1'], ['alice'], ['townsfolk'], ['bob'], ['minion']]
        mock_replace_player.return_value = [10, 20]
        mock_replace_role.return_value = [75, 120]
        mock_player_matchup.return_value = [171, 182]
        mock_role_matchup.return_value = [1, 2]
        update_stats(Data)
        mock_update_good.assert_called()
        mock_update_evil.assert_called()
        assert mock_update_player_matchup.call_count >= 1

    @patch('Helper.replace_player_array')
    @patch('Helper.replace_role_array')
    @patch('Helper.replace_player_array_matchup')
    @patch('Helper.replace_role_array_matchup')
    @patch('Helper.update_player_matchup')
    def test_update_matchups_flow(self, mock_update_player_matchup, mock_role_matchup, mock_player_matchup,
                                  mock_replace_role, mock_replace_player):
        from Helper import update_matchups
        Data = [['1'], ['alice'], ['townsfolk'], ['bob'], ['minion']]
        mock_replace_player.return_value = [10, 20]
        mock_replace_role.return_value = [75, 120]
        mock_player_matchup.return_value = [171, 182]
        mock_role_matchup.return_value = [1, 2]
        update_matchups(Data)
        assert mock_update_player_matchup.call_count >= 1

class TestWorkbookOperations:
    def test_open_and_close_workbook_edit(self):
        with patch('Helper.openpyxl.load_workbook') as mock_load:
            mock_wb = MagicMock()
            mock_wb.__getitem__.return_value = MagicMock()
            mock_load.return_value = mock_wb
            from Helper import open_workbook_edit, close_workbook_edit, workbook_edit, sheet_edit
            open_workbook_edit()
            assert workbook_edit is not None
            assert sheet_edit is not None
            close_workbook_edit()
            # After close, globals should be None
            from Helper import workbook_edit as wb_after, sheet_edit as sh_after
            assert wb_after is None
            assert sh_after is None

    def test_refresh_data_workbook(self):
        with patch('Helper.openpyxl.load_workbook') as mock_load:
            mock_wb = MagicMock(); mock_wb.__getitem__.return_value = MagicMock()
            mock_load.return_value = mock_wb
            from Helper import refresh_data_workbook, workbook, sheet
            # Set existing globals to mocks
            refresh_data_workbook()
            from Helper import workbook as wb_after, sheet as sh_after
            assert wb_after is not None
            assert sh_after is not None

    def test_recalculate_and_cache_workbook_windows(self):
        with patch('Helper.platform.system', return_value='Windows'), \
             patch('Helper.win32') as mock_win32:
            excel = MagicMock()
            excel.Workbooks.Open.return_value = MagicMock()
            mock_win32.gencache.EnsureDispatch.return_value = excel
            from Helper import recalculate_and_cache_workbook
            recalculate_and_cache_workbook()
            assert excel.Workbooks.Open.called
            assert excel.CalculateFull.called
            assert excel.Quit.called
