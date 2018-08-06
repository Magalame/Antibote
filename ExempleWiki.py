# -*- coding: utf-8 -*-
"""
Created on Thu Jun 21 20:47:11 2018


Télécharge des pages au hasard sur wikipedia puis les soumets à antidote.
Note, les modifications faite sur antidote ne sont pas écrites sur wiki. 
Il s'agit plus de visualiser les fautes et les corriger manuellement ensuite

"""
import requests
import mwparserfromhell
from collections import deque
import time

#quelconque avantage en utilisant un generateur? Je crois pas (pas d'avantage à "suivre" le gcontinue),mais pas d'inconvénient?

def bonTitre(titre,listeInterdits):
    if listeInterdits == []:
        return True
    else:
        return (listeInterdits[0] not in titre) and bonTitre(titre,listeInterdits[1:])

def queryGenerator(url,request={},suivreSeed=False): #from API:Query documentation
    
    
    global requestsRes
    lastContinue = {}
    while True:
        # Clone original request
        req = request.copy()
        # Modify it with the values returned in the 'continue' section of the last result.
        if suivreSeed:
            req.update(lastContinue)
        # Call API
        result = requests.get(url, params=req)
        
#        print(result.request.url,flush=True)
        result = result.json()
        requestsRes.append(result)
        
        if 'error' in result:
            raise ValueError("Error has been returned:"+str(result['error']))
        if 'warnings' in result:
            print(result['warnings'])
        if 'query' in result:
            
            idPage = list(result['query']['pages'].keys())[0]
            title = result['query']['pages'][idPage]['title']
            extract = result['query']['pages'][idPage]['extract']
            
#            print(title)
            
            if bonTitre(title,['Discussion utilisateur:','Discussion:','Catégorie:','Discussion Wikipédia:','Utilisateur:','Modèle:','Portail:','Wikipédia:','Module:','Fichier:','Discussion modèle:','Discussion fichier:','Projet:','Sujet:']):
                print(title)
                if extract != "":
                    yield extract,idPage,title
                    
        if 'continue' not in result:
            break
        
        lastContinue = result['continue']
        
def docGenerator(url="",queryGen=None):
    if not queryGen and not url:
        raise ValueError("Le générateur de requêtes par défaut requiert une url")
    if not queryGen:
        queryGen = queryGenerator(url)
    while True:
        extract,idPage,title = next(queryGen)
        yield ExtractToDoc(extract,idPage,title)
        
#t = query('https://fr.wikipedia.org/w/api.php?generator=random&prop=extracts&explaintext&exlimit=1&action=query&format=json')
queueDocuments = deque()
requestsRes = []

def docSource():
    global queueDocuments
    cor = Correcteur()
    
    def waiter(t,n):
        if len(queueDocuments) >= n:
            time.sleep(t)
            waiter(t,n)
   
    doc = docGenerator(url='https://fr.wikipedia.org/w/api.php?generator=random&prop=extracts&explaintext&exlimit=1&action=query&format=json')
    
    while True:
        waiter(50,100)
        cor.CorrigeDeMeme(next(doc))
         

def ExtractToDoc(extract,idDoc,title):
    wikicode = mwparserfromhell.parse(extract)
    sections = wikicode.filter_text()
    sections = sections[::2] #remove section titles
    return FillDoc(idDoc,sections,titre=title)
    
def FillDoc(idDoc,listeTextes,titre="TitreEx"):
    doc = Document(idDoc,titre=titre)
    for i,texte in enumerate(listeTextes):
        doc.DefinieZone(i,texte=texte)
    return doc

if __name__ == "__main__":
    docSource()