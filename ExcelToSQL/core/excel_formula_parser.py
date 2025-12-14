import re
from typing import List, Dict, Optional, Any, Tuple, Set
from dataclasses import dataclass
from enum import Enum

class TokenType(Enum):
    NUMBER = "NUMBER"
    STRING = "STRING"
    CELL_REF = "CELL_REF"
    RANGE_REF = "RANGE_REF"
    TABLE_REF = "TABLE_REF"
    FUNCTION = "FUNCTION"
    OPERATOR = "OPERATOR"
    COMMA = "COMMA"
    SHEET_REF = "SHEET_REF"
    PERCENT = "PERCENT"
    WHITESPACE = "WHITESPACE"
    EOF = "EOF"
    SEMICOLON = "SEMICOLON"
    COLON = "COLON"
    LBRACKET = "LBRACKET"
    RBRACKET = "RBRACKET"
    LPAREN = "LPAREN"
    RPAREN = "RPAREN"
    BOOLEAN = "BOOLEAN"
    ERROR = "ERROR"
    NAMED_RANGE = "NAMED_RANGE"
    HASH = "HASH"
    AT_SIGN = "AT_SIGN"
    EXCLAMATION = "EXCLAMATION"
    DOLLAR = "DOLLAR"
    AMPERSAND = "AMPERSAND"
    SPACE = "SPACE"
    UNKNOWN = "UNKNOWN"

@dataclass
class Token:
    type: TokenType
    value: str
    position: int

class ExcelFormulaTokenizer:
    def __init__(self, formula: str):
        self.formula = formula.lstrip('=')
        self.position = 0
        self.length = len(self.formula)

    def tokenize(self) -> List[Token]: 
        """i here converts formula string into list of tokens"""
        tokens = []

        while self.position < self.length:
            if self.formula[self.position].isspace():
                self.position += 1
                continue
            # using a chain of function calls connected by or wont be a bad idea 
            # cleaner than just if else 
            token = (
                self._match_number() or
                self._match_string() or
                self._match_cell_ref() or
                self._match_range_ref() or
                self._match_table_ref() or
                self._match_function() or
                self._match_operator() or
                self._match_comma() or
                self._match_sheet_ref() or
                self._match_percent() or
                self._match_semicolon() or
                self._match_colon() or
                self._match_lbracket() or
                self._match_rbracket() or
                self._match_lparen() or
                self._match_rparen() or
                self._match_boolean() or
                self._match_error() or
                self._match_named_range() or
                self._match_hash() or
                self._match_at_sign() or
                self._match_exclamation() or
                self._match_dollar() or
                self._match_ampersand() or
                self._match_space()
            )
            if token:
                tokens.append(token)
            else:
                tokens.append(Token(TokenType.UNKNOWN, self.formula[self.position], self.position))
                self.position += 1
        tokens.append(Token(TokenType.EOF, "", self.position)) # end of file marker, so parser could know where to stop
        return tokens
    
    # Placeholder for now 
    def _match_number(self) -> Optional[Token]:

        # pattern 
        # \d+\.?\d*  → 123, 123.45, 123.
        # \.\d+      → .5, .123
        # Both can be followed by [eE][+-]?\d+ for scientific notation
        pattern = r'(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?'
        match = re.match(pattern, self.formula[self.position:])
        
        if not match:
            return None
        
        start_pos = self.position
        number_str = match.group(0)
        self.position += len(number_str)
        
        # checking for percentage sign
        is_percent = False
        if self.position < self.length and self.formula[self.position] == '%':
            is_percent = True
            self.position += 1
        
        # i convert to float
        try:
            value = float(number_str)
            if is_percent:
                value = value / 100
            
            return Token(TokenType.NUMBER, value, start_pos)
        
        except ValueError:
            # should not really happen with my regex, but safety first
            self.position = start_pos  # reset pos
            return None
        

    def _match_string(self) -> Optional[Token]:
        if self.position >= self.length or self.formula[self.position] != '"':
            return None
        
        start_pos = self.position
        self.position += 1  
        value = ""
        while self.position < self.length:
            char = self.formula[self.position]
            if char == '"':
                if self.position + 1 < self.length and self.formula[self.position + 1] == '"':
                    value += '"'
                    self.position += 2 
                else:
                    self.position += 1 
                    return Token(TokenType.STRING, value, start_pos)
            else:
                value += char
                self.position += 1
        
        return Token(TokenType.STRING, value, start_pos)
    
    def _match_cell_ref(self) -> Optional[Token]:
        """
        match cell references like .. A1, B2, $A$1, $B2, A$1, XFD1048576
        supports absolute references with $ signs
        erxcel columns.. A-Z, AA-ZZ, AAA-XFD (max)
        excel rows.. 1-1048576 (max)
        """
        # Pattern.. optional$, letters, optional$, digits
        # \$?      → optional $ before column
        # [A-Z]+   → one or more letters (column)
        # \$?      → optional $ before row
        # \d+      → one or more digits (row)
        pattern = r'\$?[A-Z]+\$?\d+'
        match = re.match(pattern, self.formula[self.position:], re.IGNORECASE)
        if not match:
            return None
        start_pos = self.position
        cell_ref = match.group(0).upper() 
        self.position += len(match.group(0))
        return Token(TokenType.CELL_REF, cell_ref, start_pos)    


    def _match_range_ref(self) -> Optional[Token]:
        return None   
    def _match_table_ref(self) -> Optional[Token]:
        return None   
    def _match_function(self) -> Optional[Token]:
        return None   
    def _match_operator(self) -> Optional[Token]:
        return None   
    def _match_comma(self) -> Optional[Token]:
        return None   
    def _match_sheet_ref(self) -> Optional[Token]:
        return None   
    def _match_percent(self) -> Optional[Token]:
        return None   
    def _match_semicolon(self) -> Optional[Token]:
        return None   
    def _match_colon(self) -> Optional[Token]:
        return None   
    def _match_lbracket(self) -> Optional[Token]:
        return None   
    def _match_rbracket(self) -> Optional[Token]:
        return None  
    def _match_lparen(self) -> Optional[Token]:
        return None  
    def _match_rparen(self) -> Optional[Token]:
        return None  
    def _match_boolean(self) -> Optional[Token]:
        return None  
    def _match_error(self) -> Optional[Token]:
        return None  
    def _match_named_range(self) -> Optional[Token]:
        return None 
    def _match_hash(self) -> Optional[Token]:
        return None 
    def _match_at_sign(self) -> Optional[Token]:
        return None 
    def _match_exclamation(self) -> Optional[Token]:
        return None  
    def _match_dollar(self) -> Optional[Token]:
        return None 
    def _match_ampersand(self) -> Optional[Token]:
        return None  
    def _match_space(self) -> Optional[Token]:
        return None  
    
