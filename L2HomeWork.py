import re
def tokenize(expression) :
    tokens = re.findall(r'\d+|\+|\-|\*|/|\(|\)|\^|\=' , expression)
    return tokens



def parse_expression(tokens):
    if not tokens:
        raise ValueError("Invalid expression: Empty input")
    current_token = tokens.pop(0)
    if current_token != '=':
        raise ValueError("Invalid expression: did't start with = (not exel RE)")
    current_token = tokens.pop(0)
    result = parse_term(current_token, tokens)
    while tokens and tokens[0] in ['+' , '-']:
        operator = tokens.pop(0)
        if not tokens:
            raise ValueError("InvaIid expression: Unexpected end of input")
        if operator == '+':
            result += parse_term(tokens.pop(0), tokens)
        elif operator == '-':
            result -= parse_term(tokens.pop(0), tokens)
    return result


def parse_term(token, tokens) :
    if not token.isdigit():
            raise ValueError("invalid expression: Expected number, got '{token} '")
    result = int(token)
    while tokens and tokens[0] in ['*' , '/' , '^']:
        operator = tokens.pop(0)
        if not tokens :
            raise ValueError("InvaIid expression: Unexpected end of input")
        next_token = tokens.pop(0)
        if not next_token.isdigit():
            raise ValueError(f"Inva1id expression: Expected number, got i {next_token}")
        if operator == '*':
            result *= int(next_token)
        elif operator == '/':
            result /= int(next_token)
        elif operator == '^':
            result = pow(result , int(next_token))
    return result



def evaluate_expression(tokens):
    try:
        result = parse_expression(tokens.copy())
        print(result)
    except ValueError as e:
        print(f"Error: {str(e)}")
        
try:
    expression = '=4 ^ 2'
    tokens = tokenize(expression)
    evaluate_expression(tokens.copy())
except ValueError as e:
    print(f"Error: {str(e)}")
        
