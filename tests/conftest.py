"""Pytest configuration and fixtures."""
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch


@pytest.fixture
def mock_spreadsheet_values():
    """Mock spreadsheetValues for testing."""
    mock = MagicMock()
    mock.player_list = ['player1', 'player2', 'player3']
    mock.player_col_list = [2, 3, 4]
    mock.username_list = ['user1', 'user2', 'user3']
    mock.playercount = 3
    mock.role_list = ['townsfolk', 'outsider', 'minion', 'demon']
    mock.role_list_idx = [2, 3, 4, 5]
    mock.rolecount = 4
    mock.townsfolk = 2
    mock.outsider = 3
    mock.minion = 4
    mock.demon = 5
    mock.average_good = 6
    mock.average_evil = 7
    mock.average_total = 8
    mock.total_played = 9
    mock.matchup_row_start = 10
    mock.matchup_gap = 5
    return mock


@pytest.fixture
def test_data_dir(tmp_path):
    """Create a temporary directory for test files."""
    return tmp_path
