def to_normalize(text):
    return str(text).lower().replace(' ', '').replace(',', '')

def to_normalize_address(address):
    return str(address).lower().replace(' ','').replace(',', '').replace('Attn:', '')

def to_normalize_cfg(cfg):
    return str(cfg).lower().replace(' ', '').replace(',', '').replace('-', '')