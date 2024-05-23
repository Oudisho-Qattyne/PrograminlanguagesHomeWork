# in this file we will build simple interpreter to deal with 
# 1] SUM Function
# 2] MIN Function

import re
import openpyxl

# Excel File Path and Spreedsheat Name
file_path = "a.xlsx"
sheet_name = "a"


# Read Excel Cell Value 
def read_excel_cell(file_path, sheet_name, cell):
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[sheet_name]
        value = sheet[cell].value
        return value
    except Exception as e:
        print(f'Error : {e}')
        return 0

# sum function
def sum_values(tokens):
    # check syntax error
    if tokens[0] != '(' or tokens[-1] != ')':
        return -1
    
    # empty range => return 0
    if len (tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')':
        print('State 1')
        return 0
    
    # range => sum the values
    # TODO : = SUM(A1;A3;A4:A5)
    if len (tokens) > 2 and tokens[0] == '(' and tokens[-1] == ')':
        # range or some cells
        # range
        if re.match(r'[A-Za-z]\d+',str(tokens[1])) and tokens[2] == ':' and re.match(r'[A-Za-z]\d+',str(tokens[-2])):
            range_end_ref_num = re.findall(r'\d+',str(tokens[-2]))
            range_start_ref_num = re.findall(r'\d+',str(tokens[1]))
            range_ref_name = re.findall(r'[A-Za-z]',str(tokens[1]))
            range_cells_names = []
            for i in range(int(range_start_ref_num[0]),int(range_end_ref_num[0])+1):
                range_cells_names.append(str(range_ref_name[0])+str(i)) 
            range_values = []
            for cell in range_cells_names:
                temp = read_excel_cell(file_path,sheet_name,cell)
                range_values.append(temp)
            total = 0
            for val in range_values:
                total = total + val

            return total
        # some cells
        else:
            total = 0
            for tok in tokens:
                if re.match(r'[A-Za-z]\d+',str(tok)):
                    val = read_excel_cell(file_path,sheet_name,str(tok))
                    total = total+val
            return total
                       
# min function
# no range => error
# range => min
# cells => the smallest one
def min_values(tokens):
    # check syntax error
    if tokens[0] != '(' or tokens[-1] != ')':
        return -1
    
    # empty range => error
    if len (tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')':
        return -1
    
    if len (tokens) > 2 and tokens[0] == '(' and tokens[-1] == ')':
        # range or some cells
        # range
        if re.match(r'[A-Za-z]\d+',str(tokens[1])) and tokens[2] == ':' and re.match(r'[A-Za-z]\d+',str(tokens[-2])):
            range_end_ref_num = re.findall(r'\d+',str(tokens[-2]))
            range_start_ref_num = re.findall(r'\d+',str(tokens[1]))
            range_ref_name = re.findall(r'[A-Za-z]',str(tokens[1]))
            range_cells_names = []
            for i in range(int(range_start_ref_num[0]),int(range_end_ref_num[0])+1):
                range_cells_names.append(str(range_ref_name[0])+str(i)) 
            range_values = []
            for cell in range_cells_names:
                temp = read_excel_cell(file_path,sheet_name,cell)
                range_values.append(temp)
            m_min = min(range_values)
            return m_min
        # some cells
        else:
            m_values = []
            for tok in tokens:
                if re.match(r'[A-Za-z]\d+',str(tok)):
                    val = read_excel_cell(file_path,sheet_name,str(tok))
                    m_values.append(val)
            m_min = min(m_values)
            return m_min


# Lexical Analyzer
def lexical_analyzer(expression):
    tokens = re.findall(r'=|[A-Za-z]\d+|[A-Za-z]+|\(|\)|\:|\;', expression);
    return tokens;

# Syntax Analyzer
# check the syntax :
# 1] starts with =
# 2] confirm if the function exists in interpreter library
# 3] correct order => correct function name, ( ,range :not exist => 0 else => sum cells , )
def syntax_analyzer(tokens):
    
    # starts with = or not
    equal_token = tokens.pop(0)
    if equal_token not in ['='] :
        raise ValueError('Syntax Error : Excel Expressions must Starts with =')
    else:
        # search if exists in interpreter library
        # note that function name not case sensitive
        
        function_name_token = tokens.pop(0)
        function_exists = False
        # function does not exist
        if function_name_token not in ['SUM','MIN']:
            raise ValueError('Undefined Function !')
        
        # function exist
        else:
            function_exists = True
            function_name = function_name_token
            if function_name in ['SUM']:
                # call sum function
                res = sum_values(tokens)
                if  res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    print('sum result : ',res)
            if function_name in ['MIN'] :
                # call MIN 
                res = min_values(tokens)
                if res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    print('min result : ',res)

exp = '=MIN(A1;A5)'
t = lexical_analyzer(exp)
syntax_analyzer(t)
