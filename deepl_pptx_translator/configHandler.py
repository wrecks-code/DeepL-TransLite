import configparser
import os

config_file_path = "config.ini"

DEEPL_API_KEY, output_path, target_lang, marker_char, USE_DEEPL_API = (
    None,
    None,
    None,
    None,
    False,
)


def read_config():
    from deepl_pptx_translator import guiHandler

    global DEEPL_API_KEY, output_path, target_lang, translator, marker_char, USE_DEEPL_API

    if not os.path.exists(config_file_path):
        guiHandler.show_noconfig_error()

    # Create a ConfigParser object
    config = configparser.ConfigParser()

    # Read the configuration file
    config.read(config_file_path, encoding="utf-8")

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
