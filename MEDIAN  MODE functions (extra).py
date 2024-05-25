# in this file we will build simple interpreter to deal with 
# 1] SUM Function
# 2] MIN Function

import re
import openpyxl
import math
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
            print("s,df,oe,lc,")
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

def calc_median(values):
    if(len(values)%2 == 0):
        index = len(values)/2
        values.sort()
        first_number = values[int(index)-1]
        second_number = values[int(index)]
        median = (first_number + second_number)/2
        return median
    else:
        index = len(values)/2
        values.sort()
        median = values[math.floor(index)]
        return median
    
def MEDIAN(tokens):
    if tokens[0] != '(' or tokens[-1] != ')':
        return -1
    
    # empty range => return 0
    if len (tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')':
        print('State 1')
        return 0
    # range => sum the values
    # TODO : = SUM(A1;A3;A4:A5)
    if len (tokens) > 2 and tokens[0] == '(' and tokens[-1] == ')':
        
        
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
          
            
        else:
            
            range_values = []
            for tok in tokens:
                if re.match(r'[A-Za-z]\d+',str(tok)):
                    temp = read_excel_cell(file_path,sheet_name,str(tok))
                    range_values.append(temp)
                    
        median = calc_median(range_values)
        
        return median
        
        
def calc_modes(range_values):
    dict = {}
    for value in range_values:
        if value in dict.keys():
            dict[value] += 1
        else:
            dict[value] = 1
            
    max_frequency = max(dict.values())
    modes = [key for key, count in dict.items() if count == max_frequency]
    
    if len(modes) == len(dict.values()):
        return 'no modes' 
    
    else:
        return modes  
    
    
def MODE(tokens):
    if tokens[0] != '(' or tokens[-1] != ')':
        return -1

    if len (tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')':
        print('State 1')
        return 0
    
    if len (tokens) > 2 and tokens[0] == '(' and tokens[-1] == ')':
        
        
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
                
        
        else:
            
            range_values = []
            for tok in tokens:
                if re.match(r'[A-Za-z]\d+',str(tok)):
                    temp = read_excel_cell(file_path,sheet_name,str(tok))
                    range_values.append(temp)
                    
        modes = calc_modes(range_values)
        
        return modes
        

# Lexical Analyzer
def lexical_analyzer(expression):
    tokens = re.findall(r'=|[A-Za-z]\d+|[A-Za-z]+|\(|\)|\:|\;|\*|/|\+|-', expression)
    return tokens

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
        if function_name_token not in ['MEDIAN' , 'MODE']:
            raise ValueError('Undefined Function !')
        
        # function exist
        else:
            function_exists = True
            function_name = function_name_token
            if function_name in ['MEDIAN']:
                # call sum function
                res = MEDIAN(tokens)
                if  res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    print('MEDIAN result : ',res)
            if function_name in ['MODE']:
                # call sum function
                res = MODE(tokens)
                if  res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    print('MODE result : ',res)
            
exp = '=MODE(A1;A4;A2)'
t = lexical_analyzer(exp)
syntax_analyzer(t)
