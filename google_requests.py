import re


vr_pattern = r'(?i)вр\s*-\s*(\d{8})'
ansvr_pattern = r'(?i)ответ на вр-\s*(\d{8})'
srok_pattern =  r'(?i)срок\s*(-\s*до\s+|до\s+|:\s+)?(сегодня\s*|(\d{2}\.\d{2}\.\d{4}))\s*(до)?\s*(\d{1,2}:\d{2})?'

def wrap_enterletter(text):
    try:
        text = text.strip()
        vr_match = re.search(vr_pattern, text)
        srok_match = re.search(srok_pattern, text)

        vr_value = vr_match.group(1)
        if vr_value:
            text = text.replace(vr_match.group(), '')

        srok_value = srok_match.group(2) if srok_match else None

        text = re.sub(r'n+', 'n', text)
        text = text.strip()
        return text, vr_value, srok_value
    except Exception:
        return False

def process_text(text):
    lines = text.splitlines()
    result_lines = []
    for line in lines:
        if "Маршрут" in line:
            result_lines.append(line[:line.index("Маршрут")])
            break
        elif "_" in line:
            result_lines.append(line[:line.index("_")])
            break
        else:
            result_lines.append(line)
    lines = result_lines
    lines = [line for line in lines if not line.startswith('#')]
    lines = [line for line in lines if not any(sub in line for sub in ['Вр-', 'вр-', 'вр -', 'Вр -'])]
    lines = [line for line in lines if not line.startswith('@')]
    return '\n'.join(lines)

def wrap_outerletter(text,data=None):
    text = re.sub(r'n+', 'n', text)
    text = text.strip()
    f_text = process_text(text)
    if not data:
        vrs = re.findall(vr_pattern,text)
        ansvr = re.search(ansvr_pattern,text)
        if ansvr:
            vrs = [v for v in vrs if v != ansvr.group()[-8:]]
        return f_text, [vrs[0] if vrs else False, ansvr.group()[-8:] if ansvr else False]
    else:
        if not data[1][0]:
            data[1][0] = re.findall(vr_pattern, text)
            return data[0]+' '+f_text, data[1]
