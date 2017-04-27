"""
This is python2.7, because I haven't got the NLTK installed under
python3, and don't want to do so because if I do then I will almost
certainly get a muddle.

Not much should need to change, but the definition of "set" changes: I
can't do set1+set2, I have to do set1.union(set2)
"""

import sys, re

from nltk.corpus import wordnet

"""
Get all the places listed in your sets of cities, places an
countries. These are meaningful as potential keys.
"""

def getAllPlaces(placeFiles=["cities.txt", "places", "countries"], places=None):
    if places == None:
        places = set()
    if isinstance(placeFiles, str):
        for x in re.compile("\S+", re.DOTALL).finditer(open(placeFiles).read()):
            places.add(x.group(0).lower())
    else:
        for placeFile in placeFiles:
            getAllPlaces(placeFile, places)
    return places

def getBanks(f="banks.txt"):
    return [bank.strip().lower().split(" ") for bank in open(f)]

"""
Get all the terms that appear in any bank's name
"""
def getTerms(f="banks.txt"):
    terms = set()
    f = open(f).read().lower()
    for x in re.compile("[a-z]+").finditer(f):
        terms.add(x.group(0))
    return terms

"""
Useful things: place names plus terms that appear in some bank's name
that are not standard English words and are not in the stop list. We
have to explicitly add the place names because morphy *does* recognise
some place names (especially common ones like London, which are of
course particularly likely to occur as parts of bank names, so if we
are removing all the words that morphy recognises then we have to put
back the ones that are actually names). But we only need names that do
occur in the list of terms, because ones that don't occur in the list
of terms derived from bank names are irrelevant if what we are looking
at it is bank names, so we include terms.intersection(places).
"""
def unwords(terms, places=None, stopwords=None):
    if not places:
        places = set()
    if not stopwords:
        stopwords = set()
    potentialKeys = set()
    for t in terms:
        if not wordnet.morphy(t) and not t in stopwords:
            potentialKeys.add(t)
    return potentialKeys.union(terms.intersection(places))

STOPWORDS = {"banque", "banc", "banca"}

"""
OK: get all the terms, get all the place names, use a fixed
set of stopwords, do it
"""

def getInterestingTerms():
    return unwords(getTerms(), places=getAllPlaces(), stopwords=STOPWORDS)

import json

"""
To save it as .json, I have to turn it into a dictionary first.

So if you want to use it as a set after retrieving it, you'll have to
turn it back into one
"""

def saveInterestingTerms(terms, tfile="terms.json"):
    tfile = open(tfile, "w")
    json.dump({t:True for t in terms}, tfile)
    tfile.close()

def reloadInterestingTerms(tfile="terms.json"):
    return {x for x in json.load(open(tfile))}
