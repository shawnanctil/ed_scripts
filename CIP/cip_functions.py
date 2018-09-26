'''Libraries'''

import pandas as pd
from pandas import DataFrame, Series
import nltk.data
from nltk.corpus import stopwords
from nltk.corpus import wordnet as wn
from nltk.tokenize import TreebankWordTokenizer
from fuzzywuzzy import fuzz,process

'''
Static variables
'''

tok = TreebankWordTokenizer()
english_stops = set(stopwords.words('english'))
french_stops = set(stopwords.words('french'))

'''
Functions
'''

# Gets synsets for a given term.

def get_synset(word):
    for word in wn.synsets(word):
        return word.name()

# Get definitions

def get_def(syn):
    return wn.synset(syn).definition()

# Get lemmas

def get_lemma(syn):
    for word in wn.synset(syn).lemmas():
        return word.name()

# Assess fuzzy matching ratio

def get_fuzz(string,sectormatrix):
    r = process.extract(string,sectormatrix,limit=10,scorer=fuzz.token_sort_ratio)
    results = [s for s in r if s[1] >70]
    return len(results)

# Matches against code and returns top 2 values

def cip_fuzz_value(string,sectormatrix):
    r = process.extract(string,sectormatrix,limit=2,scorer=fuzz.token_sort_ratio)
    results = [s for s in r]
    return results

# Creates a dataframe called sector_matrix based on another dataframe's column. Should be followed with an export.

def sector_tag_en(frame):
    tok_list = [tok.tokenize(w) for w in frame]
    split_words = [w.lower() for sub in tok_list for w in sub]
    clean_words = [w for w in split_words if w not in english_stops]
    synset = [get_synset(w) for w in clean_words]
    lemmas = [get_lemma(w) if w != None else '' for w in synset]
    sector_matrix = DataFrame({'Categories': clean_words,
                               'Synsets': synset,
                               'Lemmas': lemmas})
    sec_syn = list(sector_matrix['Synsets'])
    sector_matrix['Definition'] = [get_def(w) if w != None else '' for w in sec_syn]
    return sector_matrix


def sector_tag_b(frame,sector):
    tok_list = [tok.tokenize(w) for w in frame]
    split_words = [w.lower() for sub in tok_list for w in sub]
    clean_words_en = [w for w in split_words if w not in english_stops]
    clean_words = [w for w in clean_words_en if w not in french_stops]
    word_frame = DataFrame({'Words': clean_words,
                               'Sector': sector})
    return word_frame
					
# Create function to append an excel sheet with output of dataframes

def append_df_to_excel(filename,df,sheet_name='Sheet1',startrow=None,truncate_sheet=False,**to_excel_kwargs):
    if 'engine' in to_excel_kwargs:
        to_excel_kwargs.pop('engine')
    
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    
    try:
        writer.book = load_workbook(filename)
        
        if startrow is None and sheet_name in writer.book.sheetnames:
            startrow = writer.book[sheet_name].max_row
            
        if truncate_sheet and sheet_name in writer.book.sheetnames:
            idx = writer.book.sheetnames.index(sheet_name)
            writer.book.remove(writer.book.worksheets[idx])
            writer.book.create_sheet(sheet_name,idx)
        
        writer.sheets = {ws.title:ws for ws in writer.book.worksheets}
    except FileNotFoundError:
        pass
    
    if startrow is None:
        startrow = 0
    
    df.to_excel(writer,sheet_name,startrow=startrow,**to_excel_kwargs)
    
    writer.save()
