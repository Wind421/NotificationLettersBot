import re

vr_pattern = r'(?i)вр-(\d{8})'
srok_pattern =  r'(?i)срок\s*(-\s*до\s+|до\s+|:\s+)?(\d{2}\.\d{2}\.\d{4})\s*(до)?\s*(\d{1,2}:\d{2})?'

def wrap_enter_letter(text):
    text = text.strip()
    vr_match = re.search(vr_pattern, text)
    srok_match = re.search(srok_pattern, text)

    vr_value = vr_match.group(1)
    srok_value = srok_match.group(2) if srok_match else None
    return text,vr_value,srok_value