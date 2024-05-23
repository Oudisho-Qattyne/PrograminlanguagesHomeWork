# This File Used to Build Simple Interpreter Using Python to Deal with 2 Range of Cells in Excel 
# Like this = (A1:A3) * (B1:B3) 

# Importing Modules to Deal with Regular Expressions and Excel Files
import re
import openpyxl 

# Excel File Path and Spreedsheat Name
file_path = "a.xlsx"
sheet_name = "a"

# multiply ranges
def ranges_times(first_range,first_range_values,second_range,second_range_values):
    for i in range(len(first_range)):
        print(f'{first_range[i]} * {second_range[i]} <==> {first_range_values[i]} * {second_range_values[i]} = {first_range_values[i] * second_range_values[i]}')

# division ranges
def ranges_divide(first_range,first_range_values,second_range,second_range_values):
     for i in range(len(first_range)):
        print(f'{first_range[i]} / {second_range[i]} <==> {first_range_values[i]} / {second_range_values[i]} = {first_range_values[i] / second_range_values[i]}')

# add ranges
def ranges_add(first_range,first_range_values,second_range,second_range_values):
    for i in range(len(first_range)):
        print(f'{first_range[i]} + {second_range[i]} <==> {first_range_values[i]} + {second_range_values[i]} = {first_range_values[i] + second_range_values[i]}')

# subtract ranges
def ranges_minus(first_range,first_range_values,second_range,second_range_values):
    for i in range(len(first_range)):
        print(f'{first_range[i]} - {second_range[i]} <==> {first_range_values[i]} - {second_range_values[i]} = {first_range_values[i] - second_range_values[i]}')

# Lexical Analyzer
def tokenize(expression):
    tokens = re.findall(r'=|[A-Za-z]\d+|\d+|\+|-|\*|/|\(|\)|\:', expression)
    return tokens

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

# to get the Number of Brackets in a Single Range
def get_brackets_number(range_token):
    total_left_bracket = range_token.count('(')
    total_right_brackets = range_token.count(')')
    return total_left_bracket,total_right_brackets

# Dealing with Ranges
# Input  : List of Tokenes
# Output : Ranges of Cells, Operation between Ranges
def deal_with_range(tokens):
    left_range = []
    right_range = []
    all_ranges = []
    all_ranges_cells_name = []

    # get the Operation Index
    token_string = str(tokens)
    operation_token = re.findall(r'\+|-|\*|/', token_string)
    operation_index = tokens.index(operation_token[0])

    # get the Left Range 
    for i in range(0,operation_index):
        left_range.append(tokens[i])
    
    # get the Right Range
    for i in range(operation_index+1,len(tokens)):
        right_range.append(tokens[i])
    
    # Combine Left and Right Ranges in one List
    all_ranges.append(left_range)
    all_ranges.append(right_range)
    
    # if Number of ( != Number of ) then Raise Syntax Error
    for my_range in all_ranges:
        total_left_bracket,total_right_brackets = get_brackets_number(my_range)

        if total_left_bracket > total_right_brackets:
            raise ValueError('Syntax Error : Missing Closing Bracket')
            return
    
        elif total_right_brackets > total_left_bracket:
            raise ValueError('Syntax Error : Missing Opening Bracket')
            return
    
        # range syntax correct
        else:
            # Search for Colon, Previous Operand and the Next Operand to Defin Range Start and End
            # Colon Token index
            colon_token_index = my_range.index(':')
            
            # Starting and Ending of the Range
            range_start = my_range[colon_token_index-1]
            range_end = my_range[colon_token_index+1]
            
            # Get the Starting, Ending and Name of Range Cells
            range_end_ref_num = re.findall(r'\d+',range_end)
            range_start_ref_num = re.findall(r'\d+',range_start)
            range_ref_name = re.findall(r'[A-Za-z]',range_end)
        
            # Get the Cells between Start and End of Range ex: a1:a3 => a1 a2 a3
            range_cells_names = []
            for i in range(int(range_start_ref_num[0]),int(range_end_ref_num[0])+1):
                range_cells_names.append(str(range_ref_name[0])+str(i)) 
                
            all_ranges_cells_name.append(range_cells_names)
            
    return all_ranges_cells_name,operation_token[0]
            
# Syntax Analyzer  
def parser(tokens):
    equal_token = tokens.pop(0)
    if equal_token not in ['='] :
        raise ValueError('Syntax Error : Excel Expressions must Starts with =')
    else:
        cells_range,operation = deal_with_range(tokens)
        first_range = cells_range[0]
        second_range = cells_range[1]
        first_range_values = []
        second_range_values = []
        for cell in first_range:
            temp = read_excel_cell(file_path,sheet_name,cell)
            first_range_values.append(temp)
            
        for cell in second_range:
            temp = read_excel_cell(file_path,sheet_name,cell)
            second_range_values.append(temp)
            
        if operation in ['*']:
            ranges_times(first_range,first_range_values,second_range,second_range_values)
        elif operation in ['/']:
            ranges_divide(first_range,first_range_values,second_range,second_range_values)
        elif operation in ['+']:
            ranges_add(first_range,first_range_values,second_range,second_range_values)
        else :
            ranges_minus(first_range,first_range_values,second_range,second_range_values)


# testing
e = "= ( A1 : A5 ) - (B1 : B5)"
t = tokenize(e)
parser(t)
