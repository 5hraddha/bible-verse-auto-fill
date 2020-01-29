import sys
import requests
import numpy as np
import pandas as pd
import re


# GET THE XLS ##

# excel_url = "https://raw.githubusercontent.com/cs109/2014_data/master/countries.csv"
# excel_url = "https://github.com/stephengeospy/bible-verse-auto-fill/bible_memory_verses.xls"
excel_url = "bible_memory_verses.xls"
sheet_name = "crossway"
df = pd.read_excel(excel_url, sheet_name).fillna('N/A')

# Crossway Authorization: Token 38cc98eecd667273acf3862b26bb45a319fc9793 ##
API_KEY = '38cc98eecd667273acf3862b26bb45a319fc9793'
API_URL = 'https://api.esv.org/v3/passage/text/'


def get_esv_text(reference):
    params = {
        'q': reference,
        'include-headings': False,
        'include-footnotes': False,
        'include-verse-numbers': False,
        'include-short-copyright': False,
        'include-passage-references': False
    }

    headers = {'Authorization': 'Token %s' % API_KEY}
    response = requests.get(API_URL, params=params, headers=headers)
    passages = response.json()['passages']
    return re.sub('\s\s+', ' ', passages[0].strip() if passages else 'Error: Passage not found')


if __name__ == '__main__':

    for index in range(len(df)):
        reference, text = df.iloc[index, 3:5]
        if reference != 'N/A' and text == 'N/A':
            passage_text = get_esv_text(reference)
            df.iloc[index, 4] = passage_text
            print(reference)
            print(passage_text)
        else:
            if reference == 'N/A':
                print("No Reference at {0}".format(index, text))

    df.to_excel(excel_url, sheet_name="crossway", index=False)