import re

"""
Модуль для разбивки писем и запросов на составляющее
"""

vr_pattern = r'(?i)вр\s*-\s*(\d{8})'
ansvr_pattern = r'(?i)ответ на вр-\s*(\d{8})'
request_pattern = r'(?i)RP(\d{5})'
srok_pattern =  r'(?i)срок\s*(-\s*до\s+|до\s+|:\s+)?(сегодня\s*|(\d{2}\.\d{2}\.(\d{4}|\d{2})))\s*(до)?\s*(\d{1,2}:\d{2})?'

def wrap_change(text):
    vr = re.search(vr_pattern, text)
    if vr:
        return vr.group(1)
    else:
        raise ValueError('Неправильный формат Вр.')

def wrap_enterletter(text):
    """
    Метод разбивает текст письма на составляющие:
    Сам текст - убраны лишние пробелы
    Временный номер (ВР) - ищет вр, вр-, вр -
    Срок письма - ищет срок, срок до, срок - до, срок: и даты сегодня, формат 00.00.00 и 00.00.0000 и время 00:00
    """
    try:
        text = text.strip()
        vr_match = re.search(vr_pattern, text) #поиск по паттерну
        srok_match = re.search(srok_pattern, text) #поиск по паттерну

        vr_value = vr_match.group(1)
        if vr_value:
            text = text.replace(vr_match.group(), '') #убирает из текста ВР

        srok_value = srok_match.group(2) if srok_match else None

        if any([True for word in ['кампус', 'гагарина', 'гостин', 'корпус', 'дальняя', 'мирового', 'РИП'] if word in text.lower()]):
            project = 'кампус'
        elif any([True for word in ['терасс', 'парк', 'почаин'] if word in text.lower()]):
            project = 'почаина'
        elif any([True for word in ['черниг', 'берегоукр', 'набереж'] if word in text.lower()]):
            project = 'чернига'
        else:
            project = ''

        text = re.sub(r'n+', 'n', text) #убирает лишние переносы строк
        text = text.strip()
        return text, vr_value, srok_value, project
    except Exception:
        return False

def process_text(text):
    """
    Метод убирает все лишнее из текста
    """
    lines = text.splitlines() #Разбивает на строки
    result_lines = []
    for line in lines:
        if "Маршрут" in line: #Если найден маршрут, удаляем все после него
            result_lines.append(line[:line.index("Маршрут")])
            break
        elif "_" in line: #Аналогично для нижнего подчеркивания ___
            result_lines.append(line[:line.index("_")])
            break
        else:
            result_lines.append(line)

    lines = result_lines
    lines = [line for line in lines if not line.startswith('#')] #Удаляет строку с хэштегами
    lines = [line for line in lines if not any(sub in line for sub in ['Вр-', 'вр-', 'вр -', 'Вр -'])] #Удаляет строки с вр
    lines = [line for line in lines if not line.startswith('@')] #Удаляет строки с тэгами
    return '\n'.join(lines)

def wrap_outerletter(text,data=None):
    """
    Метод разбивает текст исходящего письма на составляющие:
    Сам текст - убраны лишние пробелы и все что в рамках process_text
    Временный номер (ВР) - ищет вр, вр-, вр -
    Временный номер ответа - ответ на вр-
    Срок письма - ищет срок, срок до, срок - до, срок: и даты сегодня, формат 00.00.00 и 00.00.0000 и время 00:00
    """
    text = re.sub(r'n+', 'n', text)
    text = text.strip()
    f_text = process_text(text)

    if any([True for word in ['кампус', 'гагарина', 'гостин', 'корпус', 'дальняя', 'мирового', 'РИП'] if
            word in text.lower()]):
        project = 'кампус'
    elif any([True for word in ['терасс', 'парк', 'почаин'] if word in text.lower()]):
        project = 'почаина'
    elif any([True for word in ['черниг', 'берегоукр', 'набереж'] if word in text.lower()]):
        project = 'чернига'
    else:
        project = ''

    if not data: #первое письмо без данных
        vrs = re.findall(vr_pattern,text)
        ansvr = re.search(ansvr_pattern,text)
        if ansvr:
            vrs = [v for v in vrs if v != ansvr.group()[-8:]] #собирает все вр не совпадающие с ответом на вр
        return f_text, [[vrs[0]] if vrs else '', ansvr.group()[-8:] if ansvr else ''],project #возвращает False или вр-ы
    else: #Второе письмо с данными
        if not data[1][0]: #Если нет Вр
            data[1][0] = re.findall(vr_pattern, text)
            return data[0]+' '+f_text, data[1], data[2]

def wrap_request(text):
    """
        Метод разбивает текст исходящего запроса на составляющие:
        Сам текст - убраны лишние пробелы
        Номер запроса - ищет RP
        Срок письма - ищет срок, срок до, срок - до, срок: и даты сегодня, формат 00.00.00 и 00.00.0000 и время 00:00
        """
    try:
        text = text.strip()
        request_match = re.search(request_pattern, text)
        srok_match = re.search(srok_pattern, text)

        request_value = request_match.group()
        if request_value:
            text = text.replace(request_match.group(), '')

        srok_value = srok_match.group(2) if srok_match else None
        return text, request_value, srok_value
    except Exception:
        return False