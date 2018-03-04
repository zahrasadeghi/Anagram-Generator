# -*- coding: utf-8 -*-
from itertools import permutations
from openpyxl import load_workbook


def fillDictionary(fileName, persian_words):
    wb = load_workbook(filename=fileName, read_only=True)
    ws = wb.active
    for row in ws.rows:
        if row[0].value:
            if row[0].value and row[0].value[0] in persian_words.keys():
                persian_words[row[0].value[0]].add(row[0].value.replace('‌', ' '))
            else:
                persian_words[row[0].value[0]] = set([row[0].value.replace('‌', ' ')])

        if row[1].value:
            if row[1].value[0] in persian_words.keys():
                persian_words[row[1].value[0]].add(row[1].value.replace('‌', ' '))
            else:
                persian_words[row[1].value[0]] = set([row[1].value.replace('‌', ' ')])

        if row[2].value:
            if row[2].value[0] in persian_words.keys():
                persian_words[row[2].value[0]].add(row[2].value.replace('‌', ' '))
            else:
                persian_words[row[2].value[0]] = set([row[2].value.replace('‌', ' ')])

        if row[3].value:
            if row[3].value[0] in persian_words.keys():
                persian_words[row[3].value[0]].add(row[3].value.replace('‌', ' '))
            else:
                persian_words[row[3].value[0]] = set([row[3].value.replace('‌', ' ')])

        if row[4].value:
            if row[4].value[0] in persian_words.keys():
                persian_words[row[4].value[0]].add(row[4].value.replace('‌', ' '))
            else:
                persian_words[row[4].value[0]] = set([row[4].value.replace('‌', ' ')])


def isPersian(word):
    if word in persian_words[word[0]]:
        return True
    return False


def cleanSpaces(word):
    while word[0] == ' ':
        word = word[1:]
    while word[-1] == ' ':
        word = word[:-1]
    return word


def findAnagrams(word):
    perms = [''.join(p) for p in permutations(word)]
    answer = set()
    for word in perms:
        word = cleanSpaces(word)
        if isPersian(word):
            answer.add(word)
    return answer

persian_words = {}
fillDictionary('PersianWords.xlsx', persian_words)

while True:
    word = input('enter your word : ')
    word = str(word.replace('‌', ' '))
    answer = findAnagrams(word)
    print(answer)
