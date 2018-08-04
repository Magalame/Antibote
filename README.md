# Antibote
Interface python pour l'API d'Antidote 
Une version valide d'Antidote est requise

## API d'Antidote

Antidote fonctionne sur la base d'interfaces COM (voir la documentation dans le zip). Il y a trois composants principaux:

- Code python (client COM), c'est celui qui fait X tâche définie par le développeur

- serveur COM écrit par le développeur (ici InterfaceCOM.py, « Correcteur.Antidote »), il fait la jonction entre le code de départ, et l'API en soi. 
L'API nécessite que le serveur intermédiaire ait une certaine liste de fonctions implémentée (spécifiée dans les docs d'Antidote), 
l'API s'en servira ensuite pour communiquer au serveur ses requêtes.

- serveur COM (API écrite par Antidote, « Antidote.ApiOle » ). 

## COM et python
La documentation est relativement éparse, toutefois heureuseusement quelque chose de très simple suffit. Quelques choses à retenir toutefois:
- tout objet COM doit être enregistré par un identifiant unique, pour en obtenir un et l'afficher, utilisez:
```
import pythoncom
print(pythoncom.CreateGuid())
```
il doit être ensuite remplacer l'ID de ces lignes dans InterfaceCOM.py:

`_reg_clsid_ = "{D172DF78-D3A9-47CF-B462-E4F2DC9C70F5}"`

`localserver.serve(['{D172DF78-D3A9-47CF-B462-E4F2DC9C70F5}'])`

- tout object COM doit d'abord être engistré, et peut ensuite être utilisé, d'où les lignes dans InterfaceCOM.py:

```
  if '--register' in sys.argv[1:]  or '--unregister' in sys.argv[1:]:
      print("Enregistrement du serveur COM")
      win32com.server.register.UseCommandLine(ServeurCorrecteur)
  else:
      localserver.serve(['{D390AE78-D6A2-47CF-B462-E4F2DC9C70F5}'])

```

- en termes de gestion de la mémoire, la ligne `_reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER` se charge de séparer la mémoire
du serveur intermédiaire du contexte dans lequelle il est appelé. De mon expérience, si la mémoire n'est pas séparée, 
Antidote plante. 

[docu1](http://www.devshed.com/c/a/python/windows-programming-in-python-creating-com-servers)

[docu2](https://stackoverflow.com/questions/1054849/consuming-python-com-server-from-net)

## Interface python

### InterfaceCOM.py

- L'interaction entre le serveur intermédiaire et l'API est basée autour d'objets "Document", définis dans 
"Outils.py". Ils sont basé sur la structure implicite détaillée dans les documents officiels:

  1. Un document a plusieurs zones de texte
  2. Chaque zone de texte contient...du texte
  3. Chaque document a un ID unique
  4. Le serveur doit garder trace du document actif (via son ID), ça détermine le document en cours d'édition 

- Le serveur contient deux/trois fonctions qui rendent facile d'y téléverser des Documents, 
ce qui les rend accessibles à l'API d'Antidote, puis de les télécharger une fois corrigés.

- Pour lancer le serveur intermédiaire:
```
python3 ./InterfaceCOM.py --enregistrer
python3 ./InterfaceCOM.py
```

### Outils.py

- Contient l'implémentation de Document
- Et celle de Correcteur, qui se veut être une manière relativement simple d'interagir avec le serveur intermédiaire:

```
doc = Document()
doc.AjouteTexte(["bonjou","ça va?"]) #chaque item de la liste sera converti en une zone de texte
cor = Correcteur()
resultat = cor.CorrigeDeMeme(doc)  
resultat.AfficheDoc() #liste toutes les zones de texte dans le document
```




