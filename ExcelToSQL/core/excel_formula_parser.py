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
    LPAREN = "LPAREN"
    RPAREN = "RPAREN"
    COMMA = "COMMA"
    SHEET_REF = "SHEET_REF"
    PERCENT = "PERCENT"
    EOF = "EOF"


@dataclass
class Token:
    type: TokenType
    value: Any
    position: int


class ExcelTokenizer:
    
    def __init__(self, formula: str):
        self.formula = formula.lstrip('=').strip()
        self.pos = 0
        self.length = len(self.formula)
        
    def tokenize(self) -> List[Token]:
        tokens = []
        
        while self.pos < self.length:
            if self.formula[self.pos].isspace():
                self.pos += 1
                continue
            
            token = (
                self._match_string() or
                self._match_table_reference() or
                self._match_sheet_reference() or
                self._match_range() or
                self._match_cell_reference() or
                self._match_number() or
                self._match_function() or
                self._match_operator() or
                self._match_punctuation()
            )
            
            if token:
                tokens.append(token)
            else:
                self.pos += 1
        
        tokens.append(Token(TokenType.EOF, None, self.pos))
        return tokens
    
    def _match_string(self) -> Optional[Token]:
        if self.formula[self.pos] != '"':
            return None
        
        start = self.pos
        self.pos += 1
        value = ""
        
        while self.pos < self.length:
            if self.formula[self.pos] == '"':
                if self.pos + 1 < self.length and self.formula[self.pos + 1] == '"':
                    value += '"'
                    self.pos += 2
                else:
                    self.pos += 1
                    return Token(TokenType.STRING, value, start)
            else:
                value += self.formula[self.pos]
                self.pos += 1
        
        return Token(TokenType.STRING, value, start)
    
    def _match_table_reference(self) -> Optional[Token]:
        if not re.match(r'[A-Za-z_]\w*\[', self.formula[self.pos:]):
            return None
        
        start = self.pos
        
        table_match = re.match(r'([A-Za-z_]\w*)', self.formula[self.pos:])
        if not table_match:
            return None
        
        table_name = table_match.group(1)
        self.pos += len(table_name)
        
        if self.pos >= self.length or self.formula[self.pos] != '[':
            self.pos = start
            return None
        
        bracket_count = 0
        column_start = self.pos
        while self.pos < self.length:
            if self.formula[self.pos] == '[':
                bracket_count += 1
            elif self.formula[self.pos] == ']':
                bracket_count -= 1
                if bracket_count == 0:
                    self.pos += 1
                    break
            self.pos += 1
        
        column_part = self.formula[column_start:self.pos]
        full_match = self.formula[start:self.pos]
        
        return Token(TokenType.TABLE_REF, {
            'table': table_name,
            'column': column_part,
            'full': full_match
        }, start)
    
    def _match_sheet_reference(self) -> Optional[Token]:
        start = self.pos
        
        if self.formula[self.pos] == "'":
            self.pos += 1
            sheet_name = ""
            while self.pos < self.length and self.formula[self.pos] != "'":
                sheet_name += self.formula[self.pos]
                self.pos += 1
            if self.pos < self.length and self.formula[self.pos] == "'":
                self.pos += 1
                if self.pos < self.length and self.formula[self.pos] == "!":
                    self.pos += 1
                    return Token(TokenType.SHEET_REF, sheet_name, start)
        
        match = re.match(r'([A-Za-z_]\w*)!', self.formula[self.pos:])
        if match:
            sheet_name = match.group(1)
            self.pos += len(match.group(0))
            return Token(TokenType.SHEET_REF, sheet_name, start)
        
        self.pos = start
        return None
    
    def _match_range(self) -> Optional[Token]:
        match = re.match(r'\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+', self.formula[self.pos:])
        if match:
            start = self.pos
            value = match.group(0)
            self.pos += len(value)
            return Token(TokenType.RANGE_REF, value, start)
        return None
    
    def _match_cell_reference(self) -> Optional[Token]:
        match = re.match(r'\$?[A-Z]+\$?\d+', self.formula[self.pos:])
        if match:
            start = self.pos
            value = match.group(0)
            self.pos += len(value)
            return Token(TokenType.CELL_REF, value, start)
        return None
    
    def _match_number(self) -> Optional[Token]:
        match = re.match(r'\d+\.?\d*|\.\d+', self.formula[self.pos:])
        if match:
            start = self.pos
            value = match.group(0)
            self.pos += len(value)
            
            if self.pos < self.length and self.formula[self.pos] == '%':
                self.pos += 1
                return Token(TokenType.NUMBER, float(value) / 100, start)
            
            return Token(TokenType.NUMBER, float(value), start)
        return None
    
    def _match_function(self) -> Optional[Token]:
        match = re.match(r'[A-Z_][A-Z0-9_.]*(?=\()', self.formula[self.pos:], re.IGNORECASE)
        if match:
            start = self.pos
            value = match.group(0).upper()
            self.pos += len(value)
            return Token(TokenType.FUNCTION, value, start)
        return None
    
    def _match_operator(self) -> Optional[Token]:
        ops = ['<>', '<=', '>=', '<', '>', '=', '+', '-', '*', '/', '^', '&']
        for op in ops:
            if self.formula[self.pos:self.pos+len(op)] == op:
                start = self.pos
                self.pos += len(op)
                return Token(TokenType.OPERATOR, op, start)
        return None
    
    def _match_punctuation(self) -> Optional[Token]:
        char = self.formula[self.pos]
        if char == '(':
            token = Token(TokenType.LPAREN, char, self.pos)
        elif char == ')':
            token = Token(TokenType.RPAREN, char, self.pos)
        elif char == ',':
            token = Token(TokenType.COMMA, char, self.pos)
        else:
            return None
        
        self.pos += 1
        return token


class ASTNode:
    pass


@dataclass
class NumberNode(ASTNode):
    value: float


@dataclass
class StringNode(ASTNode):
    value: str


@dataclass
class CellRefNode(ASTNode):
    cell: str
    sheet: Optional[str] = None


@dataclass
class RangeRefNode(ASTNode):
    range: str
    sheet: Optional[str] = None


@dataclass
class TableRefNode(ASTNode):
    table: str
    column: str
    full_ref: str


@dataclass
class FunctionNode(ASTNode):
    name: str
    args: List[ASTNode]


@dataclass
class BinaryOpNode(ASTNode):
    operator: str
    left: ASTNode
    right: ASTNode


@dataclass
class UnaryOpNode(ASTNode):
    operator: str
    operand: ASTNode


class ExcelParser:
    
    def __init__(self, tokens: List[Token]):
        self.tokens = tokens
        self.pos = 0
        self.current_sheet = None
    
    def parse(self) -> ASTNode:
        return self._parse_expression()
    
    def _current_token(self) -> Token:
        if self.pos < len(self.tokens):
            return self.tokens[self.pos]
        return self.tokens[-1]
    
    def _advance(self):
        self.pos += 1
    
    def _parse_expression(self) -> ASTNode:
        return self._parse_comparison()
    
    def _parse_comparison(self) -> ASTNode:
        left = self._parse_concat()
        
        while self._current_token().type == TokenType.OPERATOR:
            op = self._current_token().value
            if op not in ['=', '<>', '<', '>', '<=', '>=']:
                break
            self._advance()
            right = self._parse_concat()
            left = BinaryOpNode(op, left, right)
        
        return left
    
    def _parse_concat(self) -> ASTNode:
        left = self._parse_addition()
        
        while self._current_token().type == TokenType.OPERATOR and self._current_token().value == '&':
            self._advance()
            right = self._parse_addition()
            left = BinaryOpNode('&', left, right)
        
        return left
    
    def _parse_addition(self) -> ASTNode:
        left = self._parse_multiplication()
        
        while self._current_token().type == TokenType.OPERATOR:
            op = self._current_token().value
            if op not in ['+', '-']:
                break
            self._advance()
            right = self._parse_multiplication()
            left = BinaryOpNode(op, left, right)
        
        return left
    
    def _parse_multiplication(self) -> ASTNode:
        left = self._parse_power()
        
        while self._current_token().type == TokenType.OPERATOR:
            op = self._current_token().value
            if op not in ['*', '/']:
                break
            self._advance()
            right = self._parse_power()
            left = BinaryOpNode(op, left, right)
        
        return left
    
    def _parse_power(self) -> ASTNode:
        left = self._parse_unary()
        
        if self._current_token().type == TokenType.OPERATOR and self._current_token().value == '^':
            self._advance()
            right = self._parse_power()
            return BinaryOpNode('^', left, right)
        
        return left
    
    def _parse_unary(self) -> ASTNode:
        token = self._current_token()
        
        if token.type == TokenType.OPERATOR and token.value in ['-', '+']:
            self._advance()
            operand = self._parse_unary()
            return UnaryOpNode(token.value, operand)
        
        return self._parse_primary()
    
    def _parse_primary(self) -> ASTNode:
        token = self._current_token()
        
        if token.type == TokenType.NUMBER:
            self._advance()
            return NumberNode(token.value)
        
        if token.type == TokenType.STRING:
            self._advance()
            return StringNode(token.value)
        
        sheet = None
        if token.type == TokenType.SHEET_REF:
            sheet = token.value
            self._advance()
            token = self._current_token()
        
        if token.type == TokenType.CELL_REF:
            self._advance()
            return CellRefNode(token.value, sheet)
        
        if token.type == TokenType.RANGE_REF:
            self._advance()
            return RangeRefNode(token.value, sheet)
        
        if token.type == TokenType.TABLE_REF:
            self._advance()
            return TableRefNode(
                token.value['table'],
                token.value['column'],
                token.value['full']
            )
        
        if token.type == TokenType.FUNCTION:
            return self._parse_function()
        
        if token.type == TokenType.LPAREN:
            self._advance()
            expr = self._parse_expression()
            if self._current_token().type == TokenType.RPAREN:
                self._advance()
            return expr
        
        self._advance()
        return NumberNode(0)
    
    def _parse_function(self) -> FunctionNode:
        func_name = self._current_token().value
        self._advance()
        
        if self._current_token().type != TokenType.LPAREN:
            return FunctionNode(func_name, [])
        self._advance()
        
        args = []
        while self._current_token().type != TokenType.RPAREN and self._current_token().type != TokenType.EOF:
            args.append(self._parse_expression())
            
            if self._current_token().type == TokenType.COMMA:
                self._advance()
            elif self._current_token().type == TokenType.RPAREN:
                break
        
        if self._current_token().type == TokenType.RPAREN:
            self._advance()
        
        return FunctionNode(func_name, args)


class ExcelToSQLConverter:
    
    def __init__(self, column_mappings: Dict[str, str], sheet_mappings: Dict[str, str]):
        self.column_mappings = column_mappings
        self.sheet_mappings = sheet_mappings
        self.current_sheet = None
    
    def convert(self, ast: ASTNode, current_sheet: str) -> str:
        self.current_sheet = current_sheet
        return self._convert_node(ast)
    
    def _convert_node(self, node: ASTNode) -> str:
        if isinstance(node, NumberNode):
            return str(node.value)
        
        elif isinstance(node, StringNode):
            escaped = node.value.replace("'", "''")
            return f"'{escaped}'"
        
        elif isinstance(node, CellRefNode):
            return self._convert_cell_ref(node)
        
        elif isinstance(node, RangeRefNode):
            return self._convert_range_ref(node)
        
        elif isinstance(node, TableRefNode):
            return self._convert_table_ref(node)
        
        elif isinstance(node, BinaryOpNode):
            return self._convert_binary_op(node)
        
        elif isinstance(node, UnaryOpNode):
            return self._convert_unary_op(node)
        
        elif isinstance(node, FunctionNode):
            return self._convert_function(node)
        
        return "NULL"
    
    def _convert_cell_ref(self, node: CellRefNode) -> str:
        col_letter = re.sub(r'[$\d]', '', node.cell)
        
        sql_col = self.column_mappings.get(col_letter, col_letter.lower())
        
        if node.sheet and node.sheet != self.current_sheet:
            table_name = self.sheet_mappings.get(node.sheet, node.sheet.lower())
            return f"{table_name}.{sql_col}"
        
        return sql_col
    
    def _convert_range_ref(self, node: RangeRefNode) -> str:
        return f"/* RANGE: {node.range} */"
    
    def _convert_table_ref(self, node: TableRefNode) -> str:
        this_row_match = re.search(r'\[\[#This Row\],\[([^\]]+)\]\]', node.column)
        if this_row_match:
            col_name = this_row_match.group(1)
            sql_col = re.sub(r'[^\w]', '_', col_name).strip('_').lower()
            return sql_col
        
        simple_match = re.search(r'^\[([^\]]+)\]$', node.column)
        if simple_match:
            col_name = simple_match.group(1)
            sql_col = re.sub(r'[^\w]', '_', col_name).strip('_').lower()
            return sql_col
        
        col_name = node.column.strip('[]')
        col_name = re.sub(r'#\w+\s+Row', '', col_name).strip(',[] ')
        sql_col = re.sub(r'[^\w]', '_', col_name).strip('_').lower()
        return sql_col if sql_col else node.table.lower()
    
    def _convert_binary_op(self, node: BinaryOpNode) -> str:
        left = self._convert_node(node.left)
        right = self._convert_node(node.right)
        
        op_map = {
            '+': '+',
            '-': '-',
            '*': '*',
            '/': '/',
            '^': 'POWER',
            '&': '||',
            '=': '=',
            '<>': '!=',
            '<': '<',
            '>': '>',
            '<=': '<=',
            '>=': '>='
        }
        
        sql_op = op_map.get(node.operator, node.operator)
        
        if sql_op == 'POWER':
            return f"POWER({left}, {right})"
        else:
            return f"({left} {sql_op} {right})"
    
    def _convert_unary_op(self, node: UnaryOpNode) -> str:
        operand = self._convert_node(node.operand)
        if node.operator == '-':
            return f"(-{operand})"
        return operand
    
    def _convert_function(self, node: FunctionNode) -> str:
        func_name = node.name.upper()
        args = [self._convert_node(arg) for arg in node.args]
        
        if func_name in ['SUM', 'AVG', 'COUNT', 'MIN', 'MAX', 'ROUND', 'ABS', 'SQRT']:
            return f"{func_name}({', '.join(args)})"
        
        if func_name == 'IF':
            if len(args) >= 2:
                condition = args[0]
                true_val = args[1]
                false_val = args[2] if len(args) > 2 else 'NULL'
                return f"CASE WHEN {condition} THEN {true_val} ELSE {false_val} END"
        
        if func_name == 'IFERROR':
            if len(args) >= 2:
                value = args[0]
                fallback = args[1]
                return f"COALESCE({value}, {fallback})"
        
        if func_name == 'AND':
            return f"({' AND '.join(args)})"
        
        if func_name == 'OR':
            return f"({' OR '.join(args)})"
        
        if func_name == 'NOT':
            return f"NOT ({args[0]})" if args else "NOT (TRUE)"
        
        if func_name in ['INDEX', 'MATCH', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP']:
            return f"/* TODO: {func_name}({', '.join(args)}) - Convert to JOIN */"
        
        if func_name in ['COUNTIFS', 'SUMIFS', 'AVERAGEIFS', 'COUNTIF', 'SUMIF']:
            return f"/* TODO: {func_name}({', '.join(args)}) - Convert to WHERE clause */"
        
        if func_name == 'LEN':
            return f"LENGTH({args[0]})" if args else "LENGTH(NULL)"
        
        if func_name in ['CONCAT', 'CONCATENATE']:
            return f"CONCAT({', '.join(args)})"
        
        if func_name in ['TRIM', 'UPPER', 'LOWER']:
            return f"{func_name}({args[0]})" if args else f"{func_name}(NULL)"
        
        return f"/* {func_name}({', '.join(args)}) */"


def convert_excel_formula_to_sql(formula: str, column_mappings: Dict[str, str], 
                                 sheet_mappings: Dict[str, str], current_sheet: str) -> str:
    try:
        tokenizer = ExcelTokenizer(formula)
        tokens = tokenizer.tokenize()
        
        parser = ExcelParser(tokens)
        ast = parser.parse()
        
        converter = ExcelToSQLConverter(column_mappings, sheet_mappings)
        sql = converter.convert(ast, current_sheet)
        
        return sql
    
    except Exception as e:
        return f"/* ERROR converting formula: {formula} - {str(e)} */"


if __name__ == "__main__":
    test_formulas = [
        "=A2*B2",
        "=SUM(C2:C10)",
        "=IF(D2>100, 'High', 'Low')",
        "=ROUND(E2/F2, 2)",
        "=(G2-H2)/H2",
        "=I2*(1+J2)",
        "=IFERROR(K2/L2, 0)",
        "=Table[[#This Row],[Amount]]*Table[[#This Row],[Price]]",
        "=Sheet2!A2+Sheet2!B2",
        "=IF(AND(M2>0, N2<100), M2*0.1, 0)",
    ]
    
    col_map = {chr(65+i): f"col_{chr(97+i)}" for i in range(26)}
    sheet_map = {"Sheet2": "other_table"}
    
    print("="*80)
    print("Excel Formula to SQL Converter - Test Suite")
    print("="*80)
    
    for formula in test_formulas:
        print(f"\nExcel: {formula}")
        sql = convert_excel_formula_to_sql(formula, col_map, sheet_map, "Sheet1")
        print(f"SQL:   {sql}")
