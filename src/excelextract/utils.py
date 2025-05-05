
def cleanConfig(config):
    new_dict = {}
    for key, value in config.items():
        new_key = key.lower()
        if isinstance(value, dict):
            new_dict[new_key] = cleanConfig(value)
        elif isinstance(value, list):
            new_list = []
            for item in value:
                if isinstance(item, dict):
                    new_list.append(cleanConfig(item))
                else:
                    new_list.append(item)
            new_dict[new_key] = new_list
        else:
            new_dict[new_key] = value
    return new_dict