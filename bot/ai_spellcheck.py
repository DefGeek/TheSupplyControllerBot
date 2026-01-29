import requests
import logging
from typing import Optional


async def check_spelling(text: str) -> Optional[str]:
    """
    Проверяет орфографию с помощью внешнего API (например, Yandex Speller)
    Возвращает исправленный текст или None, если ошибок нет
    """
    try:
        # Используем Yandex Speller API
        url = "https://speller.yandex.net/services/spellservice.json/checkText"
        params = {
            "text": text,
            "lang": "ru",
            "options": 512  # игнорировать слова с цифрами
        }

        response = requests.get(url, params=params, timeout=5)

        if response.status_code == 200:
            suggestions = response.json()

            if suggestions:
                # Исправляем текст
                words = text.split()
                corrections = {}

                for suggestion in suggestions:
                    if suggestion.get('s'):
                        corrections[suggestion['word']] = suggestion['s'][0]

                corrected_words = []
                for word in words:
                    if word in corrections:
                        corrected_words.append(corrections[word])
                    else:
                        corrected_words.append(word)

                corrected_text = ' '.join(corrected_words)

                if corrected_text != text:
                    return corrected_text

        return None

    except Exception as e:
        logging.error(f"Error in spell check: {e}")
        return None


async def check_list_spelling(items: list) -> list:
    """
    Проверяет список позиций на орфографические ошибки
    """
    corrected_items = []

    for item in items:
        corrected = await check_spelling(item['product_name'])
        if corrected:
            item['corrected_name'] = corrected
            item['has_correction'] = True
        else:
            item['corrected_name'] = item['product_name']
            item['has_correction'] = False
        corrected_items.append(item)

    return corrected_items