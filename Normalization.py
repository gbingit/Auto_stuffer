def to_normalize(text):
    return str(text).lower().replace(' ', '').replace(',', '')

def to_normalize_address(address):
    return str(address).lower().replace(' ','').replace(',', '').replace('Attn:', '')

def to_normalize_cfg(cfg):
    return str(cfg).lower().replace(' ', '').replace(',', '').replace('-', '')

def to_lsplit_meaning(text):
    tem_body = str(text).lower().replace(',', ' ').replace('*', ' ')
    return tem_body.split()