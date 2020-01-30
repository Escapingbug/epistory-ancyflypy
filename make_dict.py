import argparse
import json
import sqlite3
from openpyxl import load_workbook
from generate import get_encodings_cache, gen_word_encoding

WORDS_PROFILE = [ # (title, start, ends)
    ('B3', 'B9', 'B33'),
    ('B3','B35', 'B60'),
    ('C3', 'C9', 'C10'),
    ('D3', 'D9', 'D23'),
    ('E3', 'E9', 'E83'),
    ('F3', 'F9', 'F139'),
    ('G3', 'G9', 'G594'),
    ('H3', 'H9', 'H1324'),
    ('I3', 'I9', 'I445'),
    ('J3', 'J9', 'J238'),
    ('K3', 'K9', 'K78')]


def extract_words(filename, output):
    """returns { title: [words] }
    """
    wb = load_workbook(filename=filename)
    chinese_sheet = wb['Dict - CHINESE']
    words = {}

    for profile in WORDS_PROFILE:
        title_at = profile[0]
        starts_at = profile[1]
        ends_at = profile[2]
        
        title = chinese_sheet[title_at].value

        if words.get(title) is None:
            words[title] = []

        starts_row = int(starts_at[1:])
        ends_row = int(ends_at[1:])

        for row in range(starts_row, ends_row):
            cell_at = starts_at[0] + str(row)

            value = chinese_sheet[cell_at].value
            if type(value) is int:
                word = str(value)
            else:
                word = value

            if '||' in word:
                seperated = word.split('||')
                word = seperated[0].split('/')[0] + '||' + seperated[1].split('/')[0]
            else:
                word = word.split('/')[0]

            words[title].append(word)

    return words


def encode_single_character(text, c):
    for (encodings,) in c.execute('select encodings from single_characters where character = ?', text):
        encodings = encodings.split(' ')
        for encoding in encodings:
            if '*' in encoding:
                continue
            else:
                return encoding


def encode_word(text, cache):
    max_word = ''
    for encoding in gen_word_encoding(text, cache):
        if len(encoding) > len(max_word):
            max_word = encoding
    return max_word


def deal(word, c, cache):
    try:
        to_int = int(word)
        return word
    except:
        pass

    if len(word) == 1:
        #return word + '/' + encode_single_character(word, c)
        return encode_single_character(word, c)
    else:
        #return word + '/' + encode_word(word, cache)
        return encode_word(word, cache)


def encode(words, dbname):
    with sqlite3.connect(dbname) as conn:
        cache = get_encodings_cache(conn)

        c = conn.cursor()

        for name, dict_words in words.items():
            if name == 'Letter':
                continue
           
            for i in range(len(dict_words)):
                word = dict_words[i]
                if '||' in word:
                    seperated = word.split('||')
                    dict_words[i] = deal(seperated[0], c, cache) + '||' + deal(seperated[1], c, cache)
                else:
                    dict_words[i] = deal(word, c, cache)
    return words


def main():
    parser = argparse.ArgumentParser(description='make Epistory dictionary')
    parser.add_argument('dict', help='Epistory original dictionary')
    parser.add_argument('db', help='ancyflypy cache database')
    parser.add_argument('output', help='output file')
    args = parser.parse_args()
    dictionary = args.dict
    output = args.output
    words = extract_words(dictionary, output)
    words = encode(words, args.db)

    output_dict = {
            'language': '中文_ancyflypy',
            'dictionaries': [
            ]
    }

    for k, v in words.items():
        output_dict['dictionaries'].append({
            'name': k,
            'words': v
        })

    with open(output, 'w') as f:
        json.dump(output_dict, f)

if __name__ == '__main__':
    main()
