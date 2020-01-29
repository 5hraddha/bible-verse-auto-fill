import re
import requests
import pandas as pd
import configparser


# Set Environment Values
def set_env():
    # excel_url = "https://github.com/stephengeospy/bible-verse-auto-fill/bible_memory_verses.xls"
    cfg_file = "bible-verse-auto-fill.cfg"
    cfg_section_name = "bible-crossway"
    excel_url = "bible_memory_verses_api.xls"
    sheet_name = "crossway"
    return cfg_file, cfg_section_name, excel_url, sheet_name


def get_api_cfg(section):
    config = configparser.SafeConfigParser()
    config.read(cfg_file)
    cfg_dict = {}
    cfg_dict.update({'API_KEY': config.get(section, 'API_KEY')})
    cfg_dict.update({'API_URL': config.get(section, 'API_URL')})
    return cfg_dict


def set_api_params(cfg_dict):
    params = {
        'include-headings': False,
        'include-footnotes': False,
        'include-verse-numbers': False,
        'include-short-copyright': False,
        'include-passage-references': False
    }
    headers = {'Authorization': 'Token {}'.format(cfg_dict['API_KEY'])}
    cfg_dict.update({'PARAMS': params})
    cfg_dict.update({'HEADERS': headers})
    return cfg_dict


def get_esv_text(reference, cfg_dict):
    # cfg_dict['PARAMS']['q'] = reference
    cfg_dict['PARAMS'].update({'q': reference})
    response = requests.get(cfg_dict['API_URL'], params=cfg_dict['PARAMS'], headers=cfg_dict['HEADERS'])
    passages = response.json()['passages']
    return re.sub('\s\s+', ' ', passages[0].strip() if passages else 'Error: Passage not found')


if __name__ == '__main__':
    # Set Env Values
    cfg_file, cfg_section_name, excel_url, sheet_name = set_env()

    # Get API Keys using ConfigParser
    cfg_dict = get_api_cfg(cfg_section_name)

    # Create Request String for API call
    cfg_dict = set_api_params(cfg_dict)

    # Create Pandas DataFrame for Excel
    df = pd.read_excel(excel_url, sheet_name).fillna('N/A')

    # Iterate over each Reference to pull ESV Verse for those not found in excel
    for index in range(len(df)):
        reference, text = df.iloc[index, 3:5]
        if reference != 'N/A' and text == 'N/A':
            passage_text = get_esv_text(reference, cfg_dict)
            df.iloc[index, 4] = passage_text
            print(reference)
            print(passage_text)
        else:
            if reference == 'N/A':
                print("No Reference at {0}".format(index, text))

    # Write back to the excel sheet with updated info
    df.to_excel(excel_url, sheet_name="crossway", index=False)