

def regex_ignore_case(text):
    out = ''
    for s in text:
        out += f'[{s.upper()}{s.lower()}]'

    return out
