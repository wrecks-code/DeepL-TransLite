from deepl_pptx_translator import apiHandler, configHandler, guiHandler


def main():
    configHandler.read_config()
    apiHandler.initialize_deepl_api()
    guiHandler.mainGUI()


if __name__ == "__main__":
    main()
