import re
import openpyxl 

file_path = "a.xlsx"
sheet_name = "a"

if len(sheet_name)>31:
    raise ValueError('the name that you type has more than 31 characters')
elif len(sheet_name) ==0:
    raise ValueError('the sheet name is blank')  
elif re.search(r'[\/*/?/]', sheet_name) :
    raise ValueError('error sheet name') 



#     #check cell name
# cell_name = cell.coordinate    
# if not re.match(r'[A-Za-z]+\d+' , cell_name):
#     raise ValueError( ' syntax error : cell name')


def tokenize(expression):
    tokens = re.findall(r'=|[A-Za-z]\d+|\d+|\+|-|\*|/|\(|\)', expression)
    return tokens

def read_excel_cell(file_path, sheet_name, cell):
    try:
        if not re.match(r'[A-Za-z]+\d+' , cell):
            raise ValueError( ' syntax error : cell name')
        wb = openpyxl.load_workbook(file_path)

        sheet = wb[sheet_name]
        value = sheet[cell].value
        return value 
    except Exception as e:
        print(f'Error : {e}')
        return 0
     
def evaluate_cell(cell):
    return read_excel_cell(file_path,sheet_name,cell)
    
def parse_expression(tokens):
    # check for empty string
    if not tokens:
        raise ValueError("Invalid expression: Empty input")
        
    # check for first token : must be =
    # here first token not =
    current_token = tokens.pop(0)
    if current_token not in ['=']:
        print("Invalid expression: Excel Expressions Must Starts with =")
        return
    # here first token is =
    else:
        next_token = tokens.pop(0)
        temp = next_token
        # here must check if next_token is a cell
        if re.match(r'[a-zA-Z]+[0-9]+', next_token):
            temp = evaluate_cell(next_token)
            
         
        result = parse_term(int(temp), tokens)
        
        while tokens and tokens[0] in ['+', '-']:
            operator = tokens.pop(0)
            next_token = tokens.pop(0)
            temp = next_token
            
            # here must check if next_token is a cell
            if re.match(r'[a-zA-Z]+[0-9]+', next_token):
                temp = evaluate_cell(next_token)
                
            if operator == '+':
                result += parse_term(int(temp), tokens)
            elif operator == '-':
                result -= parse_term(int(temp), tokens)

    return result
        
def parse_term(token, tokens):

    result = int(token)

    while tokens and tokens[0] in ['*', '/']:
        operator = tokens.pop(0)
        next_token = tokens.pop(0)
        temp = next_token
        
        # here must check if next_token is a cell
        if re.match(r'[a-zA-Z]+[0-9]+', next_token):
            temp = evaluate_cell(next_token)

        if operator == '*':
            result *= int(temp)
        elif operator == '/':
            result /= int(temp)

    return result
        
  
e = "= A1 - 5 * 100"
t = tokenize(e)
r = parse_expression(t)
print(r)
