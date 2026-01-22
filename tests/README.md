***Tests***

**Overview**

This folder contains unit tests for the project.

**Layout**
- test_bot.py: Discord bot commands, error handling, role assignment tiers, role comparisons, session/CSV management.
- test_helper.py: Lookup functions, file I/O, workbook operations, and stat update logic.
- test_spreadsheetclass.py: Initialization defaults and property setters.
- test_pairwise.py: Pairwise comparison save/append and ranking logic.

**Setup**

Install project deps, then the testing tools:

```powershell
pip install -r requirements.txt
pip install -U pytest pytest-asyncio pytest-mock
```

**Run**
- All tests:
```
python -m pytest -v
```
- Single file:
```
python -m pytest tests/test_helper.py -v
```

**Test Counts**
- Bot: 44 tests
- Helper: 30 tests
- Spreadsheetclass: 7 tests
- Pairwise: 10 tests