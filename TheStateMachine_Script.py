# -*- coding: utf-8 -*-
"""
Created on Mon Dec 11 21:57:59 2023

@author: lilza
"""
# %% importing modules
# from https://stackoverflow.com/questions/33666557/get-phonemes-from-any-word-in-python-nltk-or-other-modules

import nltk
import pandas as pd
import re

# sentiment analysis by word/line https://realpython.com/python-nltk-sentiment-analysis/

# %% building constants and loading key starter functions
nltk.download('cmudict')
arpabet = nltk.corpus.cmudict.dict()
input_file_name = "C:/Users/lilza/Documents/TheStateMachine.xlsx"

# %% reading in words and their phonemes
datafile = pd.read_excel(input_file_name, "UniqueWordStats")
# print(datafile['Word'])

words = datafile['Word'].tolist()
# print(words)

word_dict = {}
not_found_list = []
for word in words:
    try:
        word_dict[word] = arpabet[word][0]
        # print(arpabet[word][0])
    except Exception as e:
        word_dict[word] = "NOT_FOUND"
        not_found_list.append(word)
        # print(e)

# print(word_dict)

# %% adding words considered exceptions
num_not_found = 0
for word in words:
    val = word_dict[word]
    if val == "NOT_FOUND":
        num_not_found += 1

# print("Number of words that you need to do is: "+str(num_not_found))
print(not_found_list)
# can find punch, card, letter s...but how do doesnt??
print(arpabet['punch'][0])
print(arpabet['card'][0])
print(arpabet['sees'][0])
print(arpabet['does'][0])
print(arpabet["isn't"][0])

punchcard = arpabet['punch'][0]
for phon in arpabet['card'][0]:
    punchcard.append(phon)
punchcards = punchcard.copy()
punchcards.append('Z')
doesnt = arpabet['does'][0]
for phon in ('AH0', 'N', 'T'):
    doesnt.append(phon)
# punchcard.append(arpabet['card'][0])
# print(punchcard)
# print(punchcards)
# print(doesnt)

word_dict['punchcards'] = punchcards
word_dict['punchcard'] = punchcard
word_dict["doesn’t"] = doesnt
# print(word_dict["doesn't"])
# print(word_dict["doesn’t"])
# print(word_dict.keys())

# %% adding the arpabet phoneme list to datafile dataframe

# print(datafile.columns)
phonemic_pattern = []
num_syllables = datafile['Number of syllables'].tolist()
num_letters = []
num_phonemes = []

for word in words:
    phonemes = word_dict[word]
    phonemic_pattern.append(phonemes)
    num_phonemes.append(len(phonemes))
    num_letters.append(len(word))

# print(phonemic_pattern)
# print()
# print(num_syllables)
# print()
# print(num_letters)
# print()
# print(num_phonemes)

new_UniqueWordStats = pd.DataFrame(
    {'Word': words,
     'Phonemic Pattern': phonemic_pattern,
     'Number of syllables': num_syllables,
     'Number of letters': num_letters,
     'Number of phonemes': num_phonemes
    })

# %% building the phoneme_dim_table
# has columns line, phoneme, phoneme_num
# just for generating frequencies and stuff

# in order to get to phoeneme_dim_table need PoemWordStats
datafile2 = pd.read_excel(input_file_name, "PoemWordStats")
# print(datafile2.columns)
# print(datafile2['Word']) # can join words to this column

line_nums = datafile2['Line Number'].tolist()
words2 = datafile2['Word'].tolist()
num_datafile2_rows = len(line_nums)

line_nums_tbl = []
words_tbl = []
phoneme = []
phoneme_num = []

for i in range(0,num_datafile2_rows):
    # for row in datafile2
    # save line number, word
    lin = line_nums[i]
    word = words2[i]
    # get the phonemes for that word
    val = word_dict[word]
    phon_num = 0
    for phon in val:
        # for each phoneme in val
        # append info to a big list
        phon_num += 1
        phoneme_num.append(phon_num)
        line_nums_tbl.append(lin)
        words_tbl.append(word)
        phoneme.append(phon)

# print(line_nums_tbl[11:20])
# print()
# print(words_tbl[11:20])
# print()
# print(phoneme[11:20])
# print()
# print(phoneme_num[11:20])

new_PhonemeStats = pd.DataFrame(
    {'Line Number': line_nums_tbl,
     'Word': words_tbl,
     'Phoeneme': phoneme,
     'Phoneme Number': phoneme_num
    })

# %% then replace PunctuationDimTable alphabetic characters
# to just have punctuation

datafile3 = pd.read_excel(input_file_name, "PoemStats")
# line number, punctuation
# print(datafile3.columns)

line_nums = datafile3['Line Number'].tolist()
poem_lines = datafile3['Text'].tolist()
num_datafile3_rows = len(line_nums)

line_nums_tbl = []
punct_tbl = []
punct_num_tbl = []
punct_num_lin = []
punct_num = 0

for i in range(0,num_datafile3_rows):
    # for row in datafile3
    # save line number
    lin = line_nums[i]
    # do regex on text to remove alphanumeric characters
    s = re.sub('[0-9a-zA-Z\s]+', '', poem_lines[i])
    len_s = len(s)
    # get punctuation number
    if len_s > 0:
        punct_lin_num = 0
        for j in range(0,len_s):
            punct_num += 1
            punct_lin_num += 1
            line_nums_tbl.append(lin)
            punct_tbl.append(s[j])
            punct_num_tbl.append(punct_num)
            punct_num_lin.append(punct_lin_num)

# print(line_nums_tbl)
# print()
# print(punct_tbl)
# print()
# print(punct_num_tbl)
# print()
# print(punct_num_lin)
# print()

new_PunctStats = pd.DataFrame(
    {'Line Number': line_nums_tbl,
     'Punctuation': punct_tbl,
     'Punctuation Number Poem': punct_num_tbl,
     'Punctuation Number Line': punct_num_lin
    })

# %% writing all the results to the new data table

output_file_name = "C:/Users/lilza/Documents/TheStateMachine_2024.xlsx"

writer = pd.ExcelWriter(output_file_name, engine='openpyxl', mode='a')
new_UniqueWordStats.to_excel(writer, sheet_name='UniqueWordStats')
new_PhonemeStats.to_excel(writer, sheet_name='PhonemeStats')
new_PunctStats.to_excel(writer, sheet_name='PunctStats')
writer.save()
writer.close()