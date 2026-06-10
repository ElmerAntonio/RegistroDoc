import os
import json
import re
import tkinter as tk
import customtkinter as ctk
from collections import Counter
from config import BASE_DIR

# File paths
DICT_FILE = os.path.join(BASE_DIR, "utils", "es_dictionary.txt")
USER_DICT_FILE = os.path.join(BASE_DIR, "..", "user_dictionary.json")

# Typo map for automatic correction as they type
TYPO_MAP = {
    "habeses": "a veces",
    "aveses": "a veces",
    "indiciplina": "indisciplina",
    "excetera": "etcétera",
    "asi": "así",
    "tambien": "también",
    "desinteres": "desinterés",
    "reunion": "reunión",
    "citacion": "citación",
    "telefono": "teléfono",
    "dia": "día",
    "atencion": "atención",
    "educacion": "educación",
    "correcion": "corrección",
    "evaluacion": "evaluación",
    "apreciacion": "apreciación",
    "examenes": "exámenes",
    "mas": "más",
    "ademas": "además",
    "tardansas": "tardanzas",
    "tardansa": "tardanza",
    "groseria": "grosería",
    "groserias": "groserías",
    "empujo": "empujó",
    "golpeo": "golpeó",
    "distraido": "distraído",
    "participacion": "participación",
    "academico": "académico",
    "esfuerso": "esfuerzo",
    "esfuersa": "esfuerza",
    "ase": "hace",
    "aser": "hacer",
    "aciendo": "haciendo",
    "echo": "hecho",
    "hahecho": "ha hecho",
    "agresion": "agresión",
    "bocabulario": "vocabulario",
    "salon": "salón",
    "entrego": "entregó",
    "noentrego": "no entregó",
    "conducta excetera": "conducta, etcétera",
}

class SpellCheckerEngine:
    def __init__(self):
        self.words = Counter()
        self.user_words = set()
        self.ignored_words = set()
        self.load_dictionary()
        self.load_user_dictionary()

    def load_dictionary(self):
        if os.path.exists(DICT_FILE):
            try:
                with open(DICT_FILE, "r", encoding="utf-8") as f:
                    for line in f:
                        w = line.strip().lower()
                        if w:
                            self.words[w] += 1
            except Exception as e:
                print("Error loading dictionary:", e)
        
        # Add common terms in report cards to ensure they are never flagged
        reports_words = [
            "indisciplina", "inasistencia", "tardanza", "conducta", "consejería", 
            "mención", "acudiente", "premedia", "primaria", "boleta", "tutor", 
            "trimestre", "calificación", "evaluación", "apreciación", 
            "disciplina", "excelente", "comportamiento", "compañerismo", "celular",
            "atención", "esfuerzo", "participación", "citación", "promedio", 
            "dificultad", "educativo", "deberes", "respetuoso", "irrespetuoso"
        ]
        for w in reports_words:
            self.words[w] += 10

    def load_user_dictionary(self):
        if os.path.exists(USER_DICT_FILE):
            try:
                with open(USER_DICT_FILE, "r", encoding="utf-8") as f:
                    self.user_words = set(json.load(f))
            except Exception:
                pass

    def save_user_dictionary(self):
        try:
            with open(USER_DICT_FILE, "w", encoding="utf-8") as f:
                json.dump(list(self.user_words), f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def add_to_user_dict(self, word):
        w = word.strip().lower()
        if w:
            self.user_words.add(w)
            self.save_user_dictionary()

    def ignore_word(self, word):
        w = word.strip().lower()
        if w:
            self.ignored_words.add(w)

    def is_correct(self, word):
        w = word.lower()
        if w.isdigit():
            return True
        # Allow words with hyphens or apostrophes if their parts are correct
        if "-" in w:
            return all(self.is_correct(part) for part in w.split("-"))
        # Standard dictionary checks
        return (w in self.words or w in self.user_words or w in self.ignored_words)

    # Peter Norvig's candidates and correction logic
    def P(self, word): 
        return self.words[word] / sum(self.words.values()) if sum(self.words.values()) > 0 else 0

    def correction(self, word): 
        return max(self.candidates(word), key=self.P)

    def candidates(self, word): 
        word = word.lower()
        if self.is_correct(word):
            return {word}
        if word in TYPO_MAP:
            return {TYPO_MAP[word]}
        
        candidates_set = (self.known([word]) or self.known(self.edits1(word)) or self.known(self.edits2(word)) or {word})
        return candidates_set

    def known(self, words): 
        return set(w for w in words if w in self.words or w in self.user_words)

    def edits1(self, word):
        letters    = 'abcdefghijklmnopqrstuvwxyzáéíóúñ'
        splits     = [(word[:i], word[i:])    for i in range(len(word) + 1)]
        deletes    = [L + R[1:]               for L, R in splits if R]
        transposes = [L + R[1] + R[0] + R[2:] for L, R in splits if len(R)>1]
        replaces   = [L + c + R[1:]           for L, R in splits if R for c in letters]
        inserts    = [L + c + R               for L, R in splits for c in letters]
        return set(deletes + transposes + replaces + inserts)

    def edits2(self, word): 
        return (e2 for e1 in self.edits1(word) for e2 in self.edits1(e1))


class SpellcheckTextbox(ctk.CTkTextbox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.spellchecker = SpellCheckerEngine()
        self._spellcheck_job = None
        
        # Access underlying tkinter Text widget safely
        self.tk_text = getattr(self, "_textbox", None) or getattr(self, "textbox", None)
        if self.tk_text:
            self.tk_text.tag_config("misspelled", underline=True, foreground="#EF4444")
            
        # Bind events
        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<Button-3>", self._show_suggestions_menu) # Windows right click
        self.bind("<Button-2>", self._show_suggestions_menu) # macOS/Linux right click
        
        # Initial check
        self._spellcheck_job = self.after(1000, self.check_spelling)

    def _on_key_release(self, event):
        # Trigger autocorrect on spacing or punctuation keys
        if event.keysym in ('space', 'Return', 'period', 'comma', 'exclam', 'question', 'semicolon'):
            self._perform_autocorrect()
            
        # Debounce spelling check
        if self._spellcheck_job:
            self.after_cancel(self._spellcheck_job)
        self._spellcheck_job = self.after(500, self.check_spelling)

    def _perform_autocorrect(self):
        if not self.tk_text:
            return
        try:
            cursor_idx = self.tk_text.index("insert")
            line_num = int(cursor_idx.split('.')[0])
            line_text = self.tk_text.get(f"{line_num}.0", cursor_idx)
            
            # Find the last word in this line segment
            matches = list(re.finditer(r'[a-zA-ZáéíóúñÁÉÍÓÚÑüÜ]+', line_text))
            if matches:
                last_match = matches[-1]
                word = last_match.group(0)
                word_lower = word.lower()
                
                if word_lower in TYPO_MAP:
                    correction = TYPO_MAP[word_lower]
                    if word[0].isupper():
                        correction = correction[0].upper() + correction[1:]
                    
                    start_idx = f"{line_num}.{last_match.start()}"
                    end_idx = f"{line_num}.{last_match.end()}"
                    
                    # Temporarily store cursor and text status
                    self.tk_text.delete(start_idx, end_idx)
                    self.tk_text.insert(start_idx, correction)
        except Exception as e:
            print("Autocorrect exception:", e)

    def check_spelling(self):
        if not self.tk_text or not self.winfo_exists():
            return
        try:
            text = self.tk_text.get("1.0", "end - 1c")
            self.tk_text.tag_remove("misspelled", "1.0", "end")
            
            for match in re.finditer(r'[a-zA-ZáéíóúñÁÉÍÓÚÑüÜ]+', text):
                word = match.group(0)
                if len(word) <= 2:
                    continue
                    
                if not self.spellchecker.is_correct(word):
                    start_idx = f"1.0 + {match.start()} chars"
                    end_idx = f"1.0 + {match.end()} chars"
                    self.tk_text.tag_add("misspelled", start_idx, end_idx)
        except Exception as e:
            print("Spellcheck exception:", e)

    def _get_word_at_index(self, index):
        if not self.tk_text:
            return None, None, None
        try:
            line_start = self.tk_text.index(f"{index} linestart")
            line_end = self.tk_text.index(f"{index} lineend")
            line_text = self.tk_text.get(line_start, line_end)
            
            char_offset = int(self.tk_text.index(index).split('.')[1])
            
            matches = list(re.finditer(r'[a-zA-ZáéíóúñÁÉÍÓÚÑüÜ]+', line_text))
            for match in matches:
                if match.start() <= char_offset <= match.end():
                    word = match.group(0)
                    word_start = f"{line_start.split('.')[0]}.{match.start()}"
                    word_end = f"{line_start.split('.')[0]}.{match.end()}"
                    return word, word_start, word_end
        except Exception:
            pass
        return None, None, None

    def _show_suggestions_menu(self, event):
        if not self.tk_text:
            return
            
        click_idx = self.tk_text.index(f"@{event.x},{event.y}")
        word, start, end = self._get_word_at_index(click_idx)
        
        menu = tk.Menu(self, tearoff=0)
        
        if word and not self.spellchecker.is_correct(word):
            # Misspelled word: show corrections
            candidates = self.spellchecker.candidates(word)
            has_sugg = False
            
            for cand in list(candidates)[:5]:
                if cand != word:
                    # Capitalize suggestion if original word was capitalized
                    display_cand = cand
                    if word[0].isupper():
                        display_cand = cand[0].upper() + cand[1:]
                    
                    menu.add_command(
                        label=display_cand,
                        command=lambda c=display_cand, ws=start, we=end: self._replace_word(ws, we, c)
                    )
                    has_sugg = True
            
            if not has_sugg:
                menu.add_command(label="(No hay sugerencias)", state="disabled")
                
            menu.add_separator()
            menu.add_command(label=f"Agregar '{word}' al diccionario", command=lambda w=word: self._add_to_dict(w))
            menu.add_command(label="Omitir una vez", command=lambda w=word: self._ignore_word(w))
        else:
            # Normal word or empty space: show standard text options
            menu.add_command(label="Cortar", command=lambda: self.tk_text.event_generate("<<Cut>>"))
            menu.add_command(label="Copiar", command=lambda: self.tk_text.event_generate("<<Copy>>"))
            menu.add_command(label="Pegar", command=lambda: self.tk_text.event_generate("<<Paste>>"))
            menu.add_separator()
            menu.add_command(label="Seleccionar Todo", command=self._select_all)
            
        menu.post(event.x_root, event.y_root)
        return "break"

    def _replace_word(self, start, end, new_word):
        if not self.tk_text:
            return
        try:
            self.tk_text.delete(start, end)
            self.tk_text.insert(start, new_word)
            self.check_spelling()
        except Exception:
            pass

    def _add_to_dict(self, word):
        self.spellchecker.add_to_user_dict(word)
        self.check_spelling()
        
    def _ignore_word(self, word):
        self.spellchecker.ignore_word(word)
        self.check_spelling()

    def _select_all(self):
        if not self.tk_text:
            return
        try:
            self.tk_text.tag_add("sel", "1.0", "end")
        except Exception:
            pass
