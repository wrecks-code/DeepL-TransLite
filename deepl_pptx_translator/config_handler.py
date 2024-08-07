import configparser
import os

CONFIG_FILE_PATH = "config.ini"

DEEPL_API_KEY, output_path, target_lang, marker_char, USE_DEEPL_API = (
    None,
    None,
    None,
    None,
    False,
)


def read_config():
    from deepl_pptx_translator import gui_handler

    global DEEPL_API_KEY, output_path, target_lang, translator, marker_char, USE_DEEPL_API

    if not os.path.exists(CONFIG_FILE_PATH):
        gui_handler.show_noconfig_error()

    # Create a ConfigParser object
    config = configparser.ConfigParser()

    # Read the configuration file
    config.read(CONFIG_FILE_PATH, encoding="utf-8")

    DEEPL_API_KEY = config.get("SETTINGS", "deepl_api_key")
    output_path = config.get("SETTINGS", "output_path")
    target_lang = config.get("SETTINGS", "target_lang")
    marker_char = config.get("DEVELOPMENT", "marker_char")
    USE_DEEPL_API = config.getboolean("DEVELOPMENT", "use_deepl_api")
    output_path = os.path.expandvars(output_path)

    print(
        f"Deepl API key: {DEEPL_API_KEY}, Output folder: {output_path}, Target language: {target_lang}, Use Deepl API: {USE_DEEPL_API}"
    )

    if not os.path.exists(output_path):
        # Create the folder
        os.makedirs(output_path)
        print(f"Folder '{output_path}' created successfully.")
