import pytest
from unittest.mock import MagicMock, patch, mock_open
from pathlib import Path
import csv


class TestSavePairwiseComparison:
    def test_save_pairwise_comparison_fun_category(self, tmp_path):
        import Pairwise
        
        # Create temporary files
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        # Patch the file paths
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            Pairwise.save_pairwise_comparison("townsfolk", "minion", "townsfolk", "fun")
            
            # Verify the data
            with fun_file.open("r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)
                assert len(rows) == 2
                assert rows[0] == ["Role1", "Role2", "SelectedRole"]
                assert rows[1] == ["townsfolk", "minion", "townsfolk"]
    
    def test_save_pairwise_comparison_strength_category(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        # Patch the file paths
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            Pairwise.save_pairwise_comparison("demon", "outsider", "demon", "strength")
            
            with strength_file.open("r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)
                assert len(rows) == 2
                assert rows[0] == ["Role1", "Role2", "SelectedRole"]
                assert rows[1] == ["demon", "outsider", "demon"]
    
    def test_save_multiple_comparisons(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"

        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            Pairwise.save_pairwise_comparison("role1", "role2", "role1", "fun")
            Pairwise.save_pairwise_comparison("role3", "role4", "role4", "fun")
            Pairwise.save_pairwise_comparison("role5", "role6", "role5", "fun")
            
            with fun_file.open("r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)
                assert len(rows) == 4
                assert rows[1] == ["role1", "role2", "role1"]
                assert rows[2] == ["role3", "role4", "role4"]
                assert rows[3] == ["role5", "role6", "role5"]


class TestGenerateRoleRanking:
    def test_generate_role_ranking_basic_scoring(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
            writer.writerow(["townsfolk", "minion", "townsfolk"])
            writer.writerow(["townsfolk", "demon", "townsfolk"])
            writer.writerow(["minion", "demon", "minion"])
        
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            result = Pairwise.generate_role_ranking("fun")
            
            # Verify structure
            assert isinstance(result, list)
            assert all(isinstance(item, tuple) and len(item) == 2 for item in result)
            
            roles = [r for r, _ in result]
            assert "townsfolk" in roles
            assert "minion" in roles
            assert "demon" in roles
            
            # Verify scoring
            role_dict = dict(result)
            assert role_dict["townsfolk"] > role_dict["minion"]
            assert role_dict["townsfolk"] > role_dict["demon"]
            assert role_dict["minion"] > role_dict["demon"]
    
    def test_generate_role_ranking_strength_category(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
            writer.writerow(["demon", "minion", "demon"])
            writer.writerow(["demon", "outsider", "demon"])
            writer.writerow(["minion", "outsider", "minion"])
        
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            result = Pairwise.generate_role_ranking("strength")
            
            roles = [r for r, _ in result]
            assert "demon" in roles
            assert "minion" in roles
            assert "outsider" in roles
            
            role_dict = dict(result)
            assert role_dict["demon"] > role_dict["minion"]
            assert role_dict["demon"] > role_dict["outsider"]
    
    def test_generate_role_ranking_equal_scores(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        
        # Create circular preferences: A>B, B>C, C>A
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
            writer.writerow(["roleA", "roleB", "roleA"])
            writer.writerow(["roleB", "roleC", "roleB"])
            writer.writerow(["roleC", "roleA", "roleC"])
        
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            result = Pairwise.generate_role_ranking("fun")
            
            # All roles should have equal base scores (1 win, 1 loss each)
            assert len(result) == 3
            roles = [r for r, _ in result]
            assert "roleA" in roles
            assert "roleB" in roles
            assert "roleC" in roles
    
    def test_generate_role_ranking_empty_file(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            result = Pairwise.generate_role_ranking("fun")
            
            # Should return empty list
            assert result == []
    
    def test_generate_role_ranking_transitive_weighting(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        
        # A beats B, B beats C thus A should get transitive bonus
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
            writer.writerow(["roleA", "roleB", "roleA"])
            writer.writerow(["roleB", "roleC", "roleB"])
        
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            result = Pairwise.generate_role_ranking("fun")
            
            role_dict = dict(result)
            assert "roleA" in role_dict
            assert "roleB" in role_dict
            assert "roleC" in role_dict
            
            # roleA should have highest score due to transitive bonus
            assert role_dict["roleA"] > role_dict["roleB"]
            assert role_dict["roleB"] > role_dict["roleC"]
    
    def test_generate_role_ranking_weighted_scores(self, tmp_path):
        import Pairwise
        
        fun_file = tmp_path / "pairwise_fun_comparisons.csv"
        strength_file = tmp_path / "pairwise_strength_comparisons.csv"
        
        # roleA: 2 wins out of 2 comparisons (score: 1.0)
        # roleB: 1 win out of 3 comparisons (score: 0.33)
        # roleC: 0 wins out of 2 comparisons (score: -1.0)
        with fun_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
            writer.writerow(["roleA", "roleB", "roleA"])
            writer.writerow(["roleA", "roleC", "roleA"])
            writer.writerow(["roleB", "roleC", "roleB"])
        
        with strength_file.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Role1", "Role2", "SelectedRole"])
        
        with patch.object(Pairwise, 'PAIRWISE_FUN_FILE', fun_file), \
             patch.object(Pairwise, 'PAIRWISE_STRENGTH_FILE', strength_file):
            
            result = Pairwise.generate_role_ranking("fun")
            
            role_dict = dict(result)
            
            assert role_dict["roleA"] > role_dict["roleB"]
            assert role_dict["roleB"] > role_dict["roleC"]


class TestPairwiseFileInitialization:
    def test_pairwise_files_structure(self):
        import Pairwise
        
        assert hasattr(Pairwise, 'PAIRWISE_FUN_FILE')
        assert hasattr(Pairwise, 'PAIRWISE_STRENGTH_FILE')
        assert isinstance(Pairwise.PAIRWISE_FUN_FILE, Path)
        assert isinstance(Pairwise.PAIRWISE_STRENGTH_FILE, Path)
