
def applyTokenReplacement(value, tokens):
    for token, tokenValue in tokens.items():
        value = value.replace(f"%%{token}%%", str(tokenValue))

    return value