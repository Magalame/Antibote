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

from sanstitre0 import Document

class ServeurCorrecteur:
    
  #attributs liés au serveur COM en soi
  _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER #https://msdn.microsoft.com/en-us/library/windows/desktop/ms693716(v=vs.85).aspx
  _public_methods_ = []
  pub = _public_methods_
  _public_attrs_ = ['compteActiveApplication','docsProtegesEnEcriture']
  _reg_progid_ = "Correcteur.Antidote"
  
  # NEVER copy the following ID
  # Use "print(pythoncom.CreateGuid())" to make a new one.
  _reg_clsid_ = "{D390AE78-D6A2-47CF-B462-E4F2DC9C70F5}"
  
  #autres attributs
  dictDocs = {'idDocCourant':0} #dict so that it is shared and accessible between instances, might change though
  idCour = 0

  def __init__(self):
      self.compteActiveApplication = 0
      self.docsProtegesEnEcriture = False
      
  def printinfo(function):
        def wrapper(*args):
            
            print(function.__name__,args)
            res = function(*args)
            print("Résultat:",res)
            
            return res
        return wrapper

  pub.append("Coucou") #juste une fonction test
  def Coucou(self):
      print("Coucou")
      return "Coucou"
  
  pub.append("AjouteDocs")
  @printinfo
  def AjouteDocs(self,docs):
      docs = pickle.loads(docs)
      for doc in docs:
          self.AjouteDoc(doc,pickled=False)
          
  pub.append("AjouteDoc")
  @printinfo
  def AjouteDoc(self,doc,pickled=True): #pickled= False est seulement pour le cas où une fonction DANS le serveur (cf AjouteDocs) l'appelle

      if pickled:
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
  
  pub.append("RecupereDoc")
  @printinfo
  def RecupereDoc(self,idDoc):
      return pickle.dumps(self.dictDocs[int(idDoc)])
    
  pub.append("RecupereDocs")
  @printinfo
  def RecupereDocs(self,idDocListe):
      docs = []
      for idDoc in pickle.loads(idDocListe):
          docs.append(self.dictDocs[int(idDoc)])
      return pickle.dumps(docs)
          
  pub.append("AfficheDocs")
  @printinfo
  def AfficheDocs(self):
      return str(self.dictDocs)
  
  pub.append("DefinieZoneCourante")
  @printinfo
  def DefinieZoneCourante(self,idDoc,idZone):
      self.dictDocs[int(idDoc)].idZoneCourante = int(idZone)
      return self.dictDocs[int(idDoc)].idZoneCourante
  
  pub.append("DefinieDocCourant")
  @printinfo
  def DefinieDocCourant(self,idDoc):
      self.dictDocs['idDocCourant'] = int(idDoc)
      return self.dictDocs['idDocCourant']
  
  pub.append("SupprimeDoc")
  @printinfo
  def SupprimeDoc(self,idDoc):
      del self.dictDocs[int(idDoc)]

  pub.append("SupprimeDocs")
  @printinfo
  def SupprimeDocs(self,idDocListe):
      for idDoc in idDocListe:
          self.SupprimeDoc(idDoc)

  #--------------------------------------------------
  
  pub.append("ActiveApplication")
  @printinfo
  def ActiveApplication(self):
      self.compteActiveApplication += 1
      pass
  
  pub.append("ActiveDocument")
  @printinfo
  def ActiveDocument(self,idDoc):
      pass
  
  pub.append("DonneDebutSelection")
  @printinfo
  def DonneDebutSelection(self,idDoc,idZone): #on a pas de GUI donc tout ce qui implique la selection a été ignoré
      return 0
  
  pub.append("DonneFinSelection")
  @printinfo
  def DonneFinSelection(self,idDoc,idZone):
      res = len(self.dictDocs[idDoc].dictZones[idZone].texte)
      return res
  
  pub.append("DonneIdDocumentCourant")
  @printinfo
  def DonneIdDocumentCourant(self):
      return int(self.dictDocs['idDocCourant'])
  
  pub.append("DonneIdZoneDeTexte")
  @printinfo
  def DonneIdZoneDeTexte(self,idDoc,indice):
      res = list(self.dictDocs[idDoc].dictZones.keys())[indice-1]
      return res 
  
  pub.append("DonneIdZoneDeTexteCourante")
  @printinfo
  def DonneIdZoneDeTexteCourante(self,idDoc):
      res = self.dictDocs[idDoc].idZoneCourante
      return res
 
  pub.append("DonneLongueurZoneDeTexte")
  @printinfo
  def DonneLongueurZoneDeTexte(self,idDoc,idZone):
      res = len(self.dictDocs[idDoc].dictZones[idZone].texte)
      return res

  pub.append("DonneTitreDocCourant")  
  @printinfo
  def DonneTitreDocCourant(self):
      res = self.dictDocs[self.dictDocs['idDocCourant']].titre
      return res

  pub.append("DonneNbZonesDeTexte") 
  @printinfo     
  def DonneNbZonesDeTexte(self,idDoc):
      res = len(self.dictDocs[idDoc].dictZones)
      return res 

  pub.append("DonneIntervalle") 
  @printinfo
  def DonneIntervalle(self, idDoc,idZone,debut,fin):
      res = self.dictDocs[idDoc].dictZones[idZone].texte[debut:fin]
      return res
  
  pub.append("SelectionneIntervalle")
  @printinfo
  def SelectionneIntervalle(self, idDoc,idZone,debut,fin):
      pass

  pub.append("RemplaceIntervalle")  
  @printinfo
  def RemplaceIntervalle(self, idDoc,idZone,debut,fin,laChaine):
      orig = self.dictDocs[idDoc].dictZones[idZone].texte
      nouveauTexte = orig[:debut] + laChaine + orig[fin:]
      self.dictDocs[idDoc].dictZones[idZone].texte = nouveauTexte
      return nouveauTexte

if __name__=='__main__':

  if '--enregistrer' in sys.argv[1:]:
      print("Enregistrement du serveur COM")
      win32com.server.register.UseCommandLine(ServeurCorrecteur)
  else:
      localserver.serve(['{D390AE78-D6A2-47CF-B462-E4F2DC9C70F5}'])

def run(fn,*args):
    try:
        fn(*args)
    except com_error as error:
        tmp = error
        hr,msg,exc,arg = tmp.args
        print(exc[2])
