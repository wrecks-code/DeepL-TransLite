import random

import deepl_pptx_translator.configHandler as configHandler

import deepl

translator = None
translations_cache = {}


def initialize_deepl_api():
    global translator
    translator = deepl.Translator(configHandler.DEEPL_API_KEY)


def translate_text_w_deepl(text):
    if text == "":
        return ""

    if not configHandler.USE_DEEPL_API:
        if text.endswith(".pptx"):
            return "TESTFILE-" + text
        return str(random.randint(1, 100000))

    if text in translations_cache:
        return translations_cache[text]

    # Check if the text has a trailing space
    has_trailing_space = text.endswith(" ")

    # Remove trailing space for translation, add it back later if necessary
    text_to_translate = text.rstrip()

    if not text_to_translate:
        return ""

    # Make a new translation request
    translated_text = translator.translate_text(
        text_to_translate, target_lang=configHandler.target_lang
    ).text

    # Add back the trailing space if it existed
    if has_trailing_space:
        translated_text += " "

    # Cache the translation for future use
    translations_cache[text] = translated_text

    return translated_text
