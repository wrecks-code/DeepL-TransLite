from deepl_pptx_translator import api, config_handler, gui_handler


def main():
    config_handler.read_config()
    api.initialize_deepl_api()
    gui_handler.main_gui()


if __name__ == "__main__":
    main()
