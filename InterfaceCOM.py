# -*- coding: utf-8 -*-
"""
Created on Mon Jun 18 22:07:16 2018

@author: noudy
"""
import win32com.server.register,win32com.client
from win32com.server import localserver
import pythoncom
from pythoncom import com_error
import sys
import pickle
import time
     
class Document:
    
    def __init__(self,idDoc,titre=""):
        self.idDoc = int(idDoc)
        self.titre = str(titre)
        self.dictZones = {}
        self.idZoneCourante = 0
        
    def DefinieZone(self,idZone,texte=""):
        self.dictZones[idZone] = self.ZoneDeTexte(idZone,texte)
        
    def printDoc(self):
        for idZone,Zone in self.dictZones.items():
            print(idZone,":",Zone.texte)
            
    def SupprimeZone(self,idZone):
        del self.dictZones[idZone]
    
    class ZoneDeTexte:
        def __init__(self,idZone,texte=""):
            self.idZone = int(idZone)
            self.texte = str(texte)
            
            
    
            
class Correcteur:
    def __init__(self, serveur=win32com.client.Dispatch("PythonDemos.Utilities"),antidote= win32com.client.Dispatch("Antidote.ApiOle"),langue="",outil="C",versionApi="2.0"):
        self.serveur = serveur
        self.antidote = antidote
        self.outil = outil
        self.langue = ""
        self.versionApi = versionApi
        
    def Televerse(self,doc):
        self.serveur.AjouteDoc(pickle.dumps(doc))
        
    def TeleverseDocs(self,docs):
        pickledDocs = []
        for doc in docs:
            pickledDocs.append(pickle.dumps(doc))
        self.serveur.AjouteDocs(pickledDocs)
        
    def Telecharge(self,idDoc):
        return pickle.loads(self.serveur.RecupereDoc(idDoc))
    
    def TelechargeDocs(self,idDocListe):
        unpickledDocs = []
        docs = self.serveur.RecupererDocs(idDocListe)
        for doc in docs:
            unpickledDocs.append(pickle.loads(doc))
        return unpickledDocs
    
    def Corrige(self,idDoc):
        self.serveur.DefinieDocCourant(idDoc)
        self.antidote.LanceOutilDispatch2(self.serveur,self.outil,self.langue,self.versionApi)
    
    def CorrigeDoc(self,idDoc,attendre=True,d=0.5,telecharger=False):
        if not attendre:
            self.Corrige(self,idDoc)
        else:
            self.AttendreActivationApp(self.Corrige,idDoc,d=d)
            if telecharger:#disponible seulement si l'app a deu le temps d'effectuer la correction
                return self.Telecharge(idDoc)
        
    def CorrigeDocs(self,idDocListe,attendre=True,d=0.5,telecharger=False):
        res = []
        for idDoc in idDocListe:
            res.append(self.CorrigeDoc(idDoc,attendre=attendre,d=d,telecharger=telecharger))
        return res
    
    def CorrigeDeMeme(self,doc,d=0.5):
        self.Televerse(doc)
        self.AttendreActivationApp(self.CorrigeDoc,doc.idDoc,d=d)
        return self.Telecharge(doc.idDoc)
    
    def CorrigeEtAttends(self,idDoc,d=0.5):
        self.AttendreActivationApp(self.CorrigeDoc,idDoc,d=d)
    
    def AttendreActivationApp(self,fonction,*args,d=0.5):
        tmp = self.serveur.compteActiveApplication
        res = fonction(*args)
        while(self.serveur.compteActiveApplication == tmp):
            time.sleep(d)
        return res
    
    def SupprimeDoc(self,idDoc):
        self.serveur.SupprimeDocs([int(idDoc)])
        
    def SupprimeDocs(self,idDocListe):
        self.serveur.SupprimeDocs(idDocListe)
        
def test():
    
    global monCOM
    monCOM = win32com.client.Dispatch("PythonDemos.Utilities")
    doc = Document(1,titre="TitreEx")
    doc.DefinieZone(1,"Texte")
    monCOM.AjouteDoc(pickle.dumps(doc))
    doc = Document(2,titre="TitreEx")
    doc.DefinieZone(1,"Texte")
    doc.DefinieZone(2,"Texte")
    monCOM.AjouteDoc(pickle.dumps(doc))
    monCOM.DefinieDocCourant(1)
    assert monCOM.DonneIdDocumentCourant() != 0
    print(monCOM.DonneTitreDocCourant())
    print(monCOM.DonneNbZonesDeTexte(1))
    monCOM.DefinieZoneCourante(1,1)
    monCOM.DefinieZoneCourante(2,1)
    
def run(fn,*args):
    try:
        fn(*args)
    except com_error as error:
        tmp = error
        hr,msg,exc,arg = tmp.args
        print(exc[2])
#HRESULT DonneIntervalle([in] LONG idDoc, [in] LONG idZone, [in] LONG leDebut, [in] LONG laFin,[out, retval] BSTR *retour);
#HRESULT DonnePolice([in] LONG idDoc, [in] LONG idZone, [out, retval] BSTR *retour);
#HRESULT RemplaceIntervalle([in] LONG idDoc, [in] LONG idZone, [in] LONG leDebut, [in] LONG laFin,
#[in] BSTR laChaine);
#HRESULT SelectionneIntervalle([in] LONG idDoc, [in] LONG idZone, [in] LONG leDebut, [in] LONG
#laFin);