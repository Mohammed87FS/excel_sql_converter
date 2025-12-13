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
        return None   
    def _match_string(self) -> Optional[Token]:
        return None   
    def _match_cell_ref(self) -> Optional[Token]:
        return None       
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
    
