# in this file we will build simple interpreter to deal with Logical Functions :
# 1] IF Function
# 2] AND Function

import re
import openpyxl

# Excel File Path and Spreedsheat Name
file_path = "a.xlsx"
sheet_name = "a"

def evaluate_condition_3(condition,truepart,falsepart):
    # number state
    if re.match(r'\d+',condition[0]):
        print('Number Operation Number')
        operation = re.findall(r'\>|\<',condition[1])
        if operation[0] == '>':
            if condition[0] > condition[2]:
                return truepart[0]
        else:
            return falsepart[0]
        
        if operation[0] == '<':
            if condition[0] < condition[2]:
                return truepart[0]
        else:
            return falsepart[0]
        
    # cell state
    elif re.match(r'[A-Za-z]+\d+',condition[0]):
        print('Cell Operation Number')
        val = read_excel_cell(file_path,sheet_name,condition[0])
        print(f'The Value of {condition[0]} Cell : {val}')
        operation = re.findall(r'\>|\<',condition[1])
        if operation[0] == '>':
            if val > int(condition[2]):
                return truepart[0]
        else:
            return falsepart[0]
        if operation[0] == '<':
            if val < int(condition[2]):
                return truepart[0]
        else:
            return falsepart[0]
        
    # function state
    else:
        print('Function Operation Number')
        function_name_token = condition.pop(0)
        if function_name_token not in ['SUM']:
            raise ValueError('Undefined Function !')
        else:
            if function_name_token in ['SUM']:
                # call sum function
                res = sum_values(condition)
                if  res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    if res:
                        return truepart[0]
                    else:
                        return falsepart[0]

def evaluate_condition_2(condition,true_part):
    # number state
    if re.match(r'\d+',condition[0]):
        print('Number Operation Number')
        operation = re.findall(r'\>|\<',condition[1])
        if operation[0] == '>':
            if condition[0] > condition[2]:
                return true_part[0]
        else:
            return False
        
        if operation[0] == '<':
            if condition[0] < condition[2]:
                return true_part[0]
        else:
            return False
        
    # cell state
    elif re.match(r'[A-Za-z]+\d+',condition[0]):
        print('Cell Operation Number')
        val = read_excel_cell(file_path,sheet_name,condition[0])
        print(f'The Value of {condition[0]} Cell : {val}')
        operation = re.findall(r'\>|\<',condition[1])
        if operation[0] == '>':
            if val > int(condition[2]):
                return true_part[0]
        else:
            return False
        if operation[0] == '<':
            if val < int(condition[2]):
                return true_part[0]
        else:
            return False
        
    # function state
    else:
        print('Function Operation Number')
        function_name_token = condition.pop(0)
        if function_name_token not in ['SUM']:
            raise ValueError('Undefined Function !')
        else:
            if function_name_token in ['SUM']:
                # call sum function
                res = sum_values(condition)
                if  res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    if res:
                        return true_part[0]
                    else:
                        return False

# consider just one state : SUM(A1:A5)
def sum_values(tokens):
    open_bracket = 0
    close_bracket = 0
    for tok in tokens:
        if tok == '(':
            open_bracket = tokens.index('(')
            break
        
    for tok in tokens:
        if tok == ')':
            close_bracket = tokens.index(')')
            break
            
    start_range = re.findall(r'\d+',tokens[open_bracket+1])
    end_range   = re.findall(r'\d+',tokens[close_bracket-1])
    cell_name   = re.findall(r'[A-Za-z]+',tokens[open_bracket+1])
    
    range_cells_names = []
    range_values = []
    total = 0
    
    for i in range(int(start_range[0]),int(end_range[0])+1):
        range_cells_names.append(str(cell_name[0])+str(i)) 
        
    for cell in range_cells_names:
        temp = read_excel_cell(file_path,sheet_name,cell)
        range_values.append(temp)

    for val in range_values:
        total = total + val
        
    print(f'The Values of : {range_cells_names} are {range_values}')
    print(f'Total is : {total}')

    operation = re.findall(r'\>|\<',tokens[close_bracket+1])
    if operation[0] == '>':
        if total > int(tokens[close_bracket+2]):
            return True
        else:
            return False
    if operation[0] == '<':
        if total < int(tokens[close_bracket+2]):
            return True
        else:
            return False
    
# we just take the sum function
def function_comparision(tokens,cond_component):
    function_name_token = tokens.pop(0)
    if function_name_token not in ['SUM']:
        raise ValueError('Undefined Function !')
    else:
            if function_name_token in ['SUM']:
                # call sum function
                res = sum_values(tokens)
                if  res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    return res

def cell_comparision(tokens,cond_component):
    value = read_excel_cell(file_path,sheet_name,cond_component[0])
    print(f'The Value of {cond_component[0]} Cell : {value}')
    operation = re.findall(r'\>|\<',cond_component[1])
    if operation[0] == '>':
        if value > int(cond_component[2]):
            return True
        else:
            return False
    if operation[0] == '<':
        if value < int(cond_component[2]):
            return True
        else:
            return False
    
def number_comparision(tokens,cond_component):
    operation = re.findall(r'\>|\<',cond_component[1])
    
    if operation[0] == '>':
        if cond_component[0] > cond_component[2]:
            return True
        else:
            return False
    if operation[0] == '<':
        if cond_component[0] < cond_component[2]:
            return True
        else:
            return False

# we will consider 3 states:
# 1] number operation number
# 2] cell operation number
# 3] function operation number
def evaluate_condition(tokens,cond_component):
    
    if re.match(r'\d+',cond_component[0]):
        print('Number Operation Number')
        return number_comparision(tokens,cond_component)
        
    elif re.match(r'[A-Za-z]+\d+',cond_component[0]):
        print('Cell Operation Number')
        return cell_comparision(tokens,cond_component)
    
    else:
        print('Function Operation Number')
        return function_comparision(tokens,cond_component) 

# IF Function
# Syntax : IF(Condition -> required, True_Value -> required , False_Value -> Optional)
# Condition Missing ==> Error --------------------------------------------------------------------
# if(T_V,F_V)       ==> Error --------------------------------------------------------------------
# if(cond)          ==> T     | F       ==> 0 ----------------------------------------------------
# if(cond,T_V)      ==> T_V   | F       ==> 1 ----------------------------------------------------
# if(cond,T_V,F_V)  ==> T_V   | F_V     ==> 2 ----------------------------------------------------
# if(cond,,F_V)     ==> 0     | F_V     ==> 2
# if(cond,T_V,)     ==> 0     | T_V     ==> 2

def if_function(tokens):
    opening_brackets = tokens.pop(0)
    closing_brackets = tokens.pop(-1)
    
    # Syntax Error
    if opening_brackets != '(' or closing_brackets != ')':
        return -1
    
    # Correct Syntax
    else:
        # extract the condition , true value and false value
        spliter_indices = [i for i, x in enumerate(tokens) if x == ","]
        
        # no comma ==> just condition | no condition 
        if len(spliter_indices) == 0:
            print('State A :')
            print('**********')
            print('either IF() or IF(condition)')
            cond_component = []
            for tok in tokens:
                cond_component.append(tok)
                
                
            # Condition Missing ==> Error
            if len(cond_component) == 0:
                print('IF() : no Condition will Cause an Error')
                return -1
            
            # if(cond) ==> T | F  
            else:
                print('IF(condition) : will evaluate the Condition to True or False')
                result = evaluate_condition(tokens,cond_component)
                return result
        
        # one comma ==> if(T_V,F_V) ==> Error | if(cond,T_V) ==> T_V | F
        if len(spliter_indices) == 1:
            print('State B :')
            print('**********')
            print('either IF(condition,true_value) or IF(true_value,false_value)')
            comma_index = tokens.index(',')
            temp_1 = []
            temp_2 = []
            
            for i in range(comma_index):
                temp_1.append(tokens[i])

            for i in range(comma_index+1,len(tokens)):
                temp_2.append(tokens[i])
            # if(T_V,F_V) ==> Error
            if (len(temp_1) != 0 and len(temp_2) != 0)  and ('>' not in temp_1 and '>' not in temp_2) and ('<' not in temp_1 and '<' not in temp_2 ):
                print('IF(true_value,false_value) will Cause an Error')
                raise ValueError('Error')
                return -1
            
            # if(cond,T_V) ==> T_V | F
            else:
                print('IF(condition,true_value) will Evalute to true_value or False')
                return evaluate_condition_2(temp_1,temp_2)

        # two comma ==> if(cond,T_V,F_V)  ==> T_V or F_V | if(cond,,F_V) or if(cond,T_V,) ==> 0
        if len(spliter_indices) == 2:
            print('State C :')
            print('**********')
            print('either IF(condition,true_value,false_value) or IF(condition,,false_value) or IF(condition,true_value,)')
            t1 = []
            t2 = []
            t3 = []
            for i in range(0,spliter_indices[0]):
                t1.append(tokens[i])
            for i in range(spliter_indices[0]+1,spliter_indices[1]):
                t2.append(tokens[i])
            for i in range(spliter_indices[1]+1,len(tokens)):
                t3.append(tokens[i])
            return evaluate_condition_3(t1,t2,t3)

# syntax (cond1,cond2)
def and_function(tokens):
    opening_brackets = tokens.pop(0)
    closing_brackets = tokens.pop(-1)
    
    # Syntax Error
    if opening_brackets != '(' or closing_brackets != ')':
        return -1
    
    # Correct Syntax
    else:
        # extract the condition , true value and false value
        spliter_indices = [i for i, x in enumerate(tokens) if x == ","]
        cond1 = []
        cond2 = []
        for i in range(0,spliter_indices[0]):
            cond1.append(tokens[i])
        for i in range(spliter_indices[0]+1,len(tokens)):
            cond2.append(tokens[i])
        r1 = evaluate_condition(cond1,cond1)
        r2 = evaluate_condition(cond2,cond2)
        if r1 and r2:
            return True
        else:
            return False

        
    
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
    
# Lexical Analyzer
def lexical_analyzer(expression):
    print('Stage One <Lexical Analyzer> :')
    print('-------------------------------')
    print('Input Expression :\n',expression)
    tokens = re.findall(r'=|[A-Za-z]\d+|[A-Za-z]+|\(|\)|\:|\;|\,|\>|\<|\d+|\!|\=', expression)
    print('Tokens :\n',tokens)
    return tokens

# Syntax Analyzer
# check the syntax :
# 1] starts with =
# 2] confirm if the function exists in interpreter library
def syntax_analyzer(tokens,e):
    print('\nStage Two <Syntax Analyzer> :')
    print('-------------------------------')
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
        if function_name_token not in ['IF','AND','if','and']:
            raise ValueError('Undefined Function !')
        
        # function exist
        else:
            function_exists = True
            function_name = function_name_token
            if function_name in ['IF','if']:
                # call IF function
                print('IF() Function Have been Called !\n')
                print('We Considre Number,Cell and Function Comparsison')
                res = if_function(tokens)
                if  res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    print(e,' will be Evaluated to : ',res)
                    
            if function_name in ['AND','and'] :
                # call AND 
                res = and_function(tokens)
                if res == -1:
                    raise ValueError('Error: Syntax Error')
                else:
                    print('AND result : ',res)

#testing 
e = '= if(SUM(A1:A4)<2,"Hello","Goodbye")'
# e = '=and(5>2,6<3)'
t = lexical_analyzer(e)
syntax_analyzer(t,e)