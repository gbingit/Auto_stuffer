def to_normalize(text):
    return str(text).lower().replace(' ', '').replace(',', '')

def to_normalize_address(address):
    return str(address).lower().replace(' ','').replace(',', '').replace('Attn:', '')

def to_normalize_cfg(cfg):
    return str(cfg).lower().replace(' ', '').replace(',', '').replace('-', '')

def to_lsplit_meaning(text):
    tem_body = str(text).lower().replace(',', ' ').replace('*', ' ')
    return tem_body.split() #这里split是做什么的❓
# ✅基础示例
# s = "Hello World 你好 世界"
# result = s.split()
# print(result)  # 输出：['Hello', 'World', '你好', '世界']
