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

class PythonUtilities:
    
  #attributs liés au serveur COM en soi
#  _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
  _public_methods_ = ['AjouteDocs','SupprimeDoc','SupprimeDocs','RecupereDoc','ActiveApplication','RemplaceIntervalle','SelectionneIntervalle','DonneIntervalle','DefinieDocCourant','DefinieZoneCourante','AfficheDocs','DonneLongueurZoneDeTexte','ActiveApplication','DonneIdZoneDeTexte','DonneTitreDocCourant','DonneIdDocumentCourant' ,'DonneNbZonesDeTexte','AjouteDoc','DonneDebutSelection','DonneFinSelection','DonneIdZoneDeTexteCourante']
  _public_attrs_ = ['compteActiveApplication','docsProtegesEnEcriture']
  _reg_progid_ = "PythonDemos.Utilities"
  # NEVER copy the following ID
  # Use "print pythoncom.CreateGuid()" to make a new one.
  _reg_clsid_ = "{D390AE78-D6A2-47CF-B462-E4F2DC9C70F5}"
  
  #autres attributs
  dictDocs = {'idDocCourant':1} #dict so that it is shared and accessible between instances, might change though
  
  def __init__(self):
      self.compteActiveApplication = 0
      self.docsProtegesEnEcriture = False
  
  def AjouteDocs(self,docListe):
      for doc in docListe:
          self.AjouteDoc(doc)
    
  def AjouteDoc(self,doc):
      
      doc = pickle.loads(doc)
      
      print(doc)
      print(doc.dictZones)
      
      if self.docsProtegesEnEcriture:
          try:
              self.dictDocs[int(doc.idDoc)]
          #if line above does not raise exception, it means the item already exists, so we raise error
              raise ValueError("Il existe déjà un document avec cet id (" + str(doc.idDoc) + ")")
          except KeyError:
              self.dictDocs[int(doc.idDoc)] = doc
      else:   
          
          self.dictDocs[int(doc.idDoc)] = doc
  
  def RecupereDoc(self,idDoc):
      return pickle.dumps(self.dictDocs[int(idDoc)])
    
  def RecupereDocs(self,idDocListe):
      docs = []
      for idDoc in idDocListe:
          docs.append(self.RecupereDoc(idDoc))
      return docs
          
    
  def AfficheDocs(self):
      return str(self.dictDocs)
  
  def DefinieZoneCourante(self,idDoc,idZone):
      print("DefinieZoneCourante",idDoc,idZone)
      self.dictDocs[int(idDoc)].idZoneCourante = int(idZone)
      
  def DefinieDocCourant(self,idDoc):
      print("DefinieDocCourant",idDoc)
      self.dictDocs['idDocCourant'] = int(idDoc)
      print("Resultat:",self.dictDocs['idDocCourant'])
  
  def SupprimeDoc(self,idDoc):
      print("SupprimeDoc",idDoc)
      del self.dictDocs[int(idDoc)]
    
  def SupprimeDocs(self,idDocListe):
      print("SupprimeDocs",idDocListe)
      for idDoc in idDocListe:
          self.SupprimeDoc(idDoc)
    

  #--------------------------------------------------
  
  def ActiveApplication(self):
      self.compteActiveApplication += 1
      pass
  
  def ActiveDocument(self,idDoc):
      pass
  
  def DonneDebutSelection(self,idDoc,idZone):
      print("DonneDebutSelection",idDoc,idZone)
      return 0
  
  def DonneFinSelection(self,idDoc,idZone):
      print("DonneFinSelection",idDoc,idZone)
      res = len(self.dictDocs[idDoc].dictZones[idZone].texte)
      return res
  
  def DonneIdDocumentCourant(self):
      print("DonneIdDocumentCourant retourne",int(self.dictDocs['idDocCourant']))
      return int(self.dictDocs['idDocCourant'])
  
  def DonneIdZoneDeTexte(self,idDoc,indice):
      print("DonneIdZoneDeTexte",idDoc,indice)
      res = list(self.dictDocs[idDoc].dictZones.keys())[indice-1]
      print("Resultat:",res)
      return res 
  
  def DonneIdZoneDeTexteCourante(self,idDoc):
      print("DonneIdZoneDeTexteCourante",idDoc)
      res = self.dictDocs[idDoc].idZoneCourante
      print("Resultat:",res)
      return res
  
  def DonneLongueurZoneDeTexte(self,idDoc,idZone):
      print("DonneLongueurZoneDeTexte",idDoc,idZone)
      res = len(self.dictDocs[idDoc].dictZones[idZone].texte)
      print("Resultat:",res)
      return res
  
  def DonneTitreDocCourant(self):
      res = self.dictDocs[self.dictDocs['idDocCourant']].titre
      print("Resultat:",res)
      return res
      
  def DonneNbZonesDeTexte(self,idDoc):
      print("DonneNbZonesDeTexte",idDoc)
      res = len(self.dictDocs[idDoc].dictZones)
      print("Resultat:",res)
      return res 
  
  def DonneIntervalle(self, idDoc,idZone,debut,fin):
      print("DonneIntervalle",idDoc,idZone,debut,fin)
      res = self.dictDocs[idDoc].dictZones[idZone].texte[debut:fin]
      print("Resultat:",res)
      return res

  def SelectionneIntervalle(self, idDoc,idZone,debut,fin):
      pass
  
  def RemplaceIntervalle(self, idDoc,idZone,debut,fin,laChaine):
      print("RemplaceIntervalle",idDoc,idZone,debut,fin,laChaine)
      orig = self.dictDocs[idDoc].dictZones[idZone].texte
      nouveauTexte = orig[:debut] + laChaine + orig[fin:]
      print("Resultat:",nouveauTexte)
      self.dictDocs[idDoc].dictZones[idZone].texte = nouveauTexte
      
      
      
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
            
        
  
if __name__=='__main__':

  if '--register' in sys.argv[1:]  or '--unregister' in sys.argv[1:]:
      print("Registering COM server…")
      win32com.server.register.UseCommandLine(PythonUtilities)
  else:
      # start the server.
      localserver.serve(['{D390AE78-D6A2-47CF-B462-E4F2DC9C70F5}'])
#  testC = win32com.client.Dispatch("PythonDemos.Utilities")
#  test = win32com.client.Dispatch("Antidote.ApiOle")
#  test.LanceOutil2("PythonDemos.Utilities","C","fr","2.0")
#  
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