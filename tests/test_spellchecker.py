import pytest
import os
import json
import src.utils.spellchecker as sc

def test_spellchecker_engine_basic(tmp_path):
    # Mock file paths
    dict_file = tmp_path / "es_dictionary.txt"
    user_dict_file = tmp_path / "user_dictionary.json"
    
    # Write some words to dict
    dict_file.write_text("hola\nmundo\nespañol\na veces\nmaestro\nconducta", encoding="utf-8")
    
    # Override module globals
    original_dict = sc.DICT_FILE
    original_user_dict = sc.USER_DICT_FILE
    sc.DICT_FILE = str(dict_file)
    sc.USER_DICT_FILE = str(user_dict_file)
    
    try:
        # Instantiate engine
        engine = sc.SpellCheckerEngine()
        
        # Test basic checking
        assert engine.is_correct("hola") is True
        assert engine.is_correct("holas") is False
        
        # Test autocorrect mapping
        assert engine.correction("habeses") == "a veces"
        
        # Test suggest / candidates
        cands = engine.candidates("holas")
        assert "hola" in cands or len(cands) > 0
        
        # Test adding to user dictionary
        engine.add_to_user_dict("holas")
        assert engine.is_correct("holas") is True
        
        # Verify user dictionary saved
        assert os.path.exists(sc.USER_DICT_FILE)
        with open(sc.USER_DICT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            assert "holas" in data
            
    finally:
        # Restore globals
        sc.DICT_FILE = original_dict
        sc.USER_DICT_FILE = original_user_dict
