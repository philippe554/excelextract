
def applyTokenReplacement(value, tokens):
    if not isinstance(value, str):
        return value
    
    for token, tokenValue in tokens.items():
        value = value.replace(f"%%{token}%%", str(tokenValue))

    return value