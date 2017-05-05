# Auteur : Rioche Patrick   le 15/04/2017
#
# Mode d'emploi :
#
#   acc2clip.py doc_rptObjects.txt >ExportClips.vba
#
#   Ce programme genere du code vba access pour exporter
#   le contenu des tables access au format CLIPS
#
#   Entree :
#       Fichier.txt     fichier texte windows via access menu
#                       Outils de base de donnees
#                           Documentation de base de donnees
#                               selection toutes les tables
#                               export fichier texte
#
#   Sortie :
#       NA
#
#   Version :
#       V1.0    20/04/2017
#
__version__ = 'V1.0'

import os, sys

if len(sys.argv) == 1:
    print ( "Mode de d'emploi : " + __version__ )
    print ( "                                 " )
    print ( "    acc2clp.py doc_rptObjects.txt" )
    sys.exit(1)

#
#   Recuperation des arguments de la ligne de commande
#
sFic = sys.argv[1]


#
#   Initialisation global
#
dDicoTable = {}
global nRub
nRub = 1

#
#   Definition des fonctions
#
def ClearString( sTheString ):
    sS1 = sTheString.rstrip().lower()
    sS2 = sS1.replace(' ','_')
    sS3 = sS2.replace('(','')
    sS4 = sS3.replace(')','')    
    sS5 = sS4.replace(',','')
    sS6 = sS5.replace('\'','_')
    sS7 = sS6.replace('û','u')
    sS8 = sS7.replace('é','e')
    sS9 = sS8.replace('è','e')
    sSA = sS9.replace('à','a')
    sSB = sSA.replace('ç','c')
    return( sSB )
    

def AddDicoNomTable( sTable, nTable ):
    global nRub
    dDicoTable["table:" + str(nTable) + ":tbl" ] = sTable.rstrip()
    dDicoTable["table:" + str(nTable) + ":def" ] = ClearString( sTable )
    dDicoTable["table:" + str(nTable-1) + ":rub:nbe" ] = str(nRub)
    nRub = 1

    
def AddDicoNbTable( nNbTable ):
    dDicoTable["nombredetable:"] = str(nNbTable)

def AddDicoRubTable( nTable, sRub, sTyp ):
    global nRub
    dDicoTable["table:" + str(nTable) + ":rub:" + str(nRub)] = ClearString( sRub )
    dDicoTable["table:" + str(nTable) + ":typ:" + str(nRub)] = ClearString( sTyp )
    nRub = nRub + 1

def MefTypDefTemplate( sTheTyp ):
    if sTheTyp == 'texte':
        return( 'STRING' )
    elif sTheTyp == 'octet':
        return( 'INTEGER')
    elif sTheTyp == 'entie':
        return( 'INTEGER')
    elif sTheTyp == 'oui/n':
        return( 'INTEGER')
    elif sTheTyp == 'date/':
        return( 'INTEGER')
    
#
#   Ouverture du fichier doc_rptObjects.txt
#
nLigne = 0
nTable = 0


fO = open(sFic, "r")

for sLigne in  fO.readlines():
    #
    #   Selection des noms des tables
    #
    if sLigne[1:7] == "Table:":
        AddDicoNomTable( sLigne[8:50], nTable )
        nTable = nTable + 1

    #
    #   Selection des rubriques et type de tables
    #
    #print( ">" + sLigne[65:76] + "<" )
    if sLigne[65:70] == "Texte":
        #print ( sLigne[65:70] )
        AddDicoRubTable( nTable - 1, sLigne[9:64], sLigne[65:70] )
    if sLigne[65:70] == "Octet":
        #print ( sLigne[65:70] )
        AddDicoRubTable( nTable - 1, sLigne[9:64], sLigne[65:70] )
    elif sLigne[65:71] == "Entier":
        #print ( sLigne[65:71] )
        AddDicoRubTable( nTable - 1, sLigne[9:64], sLigne[65:70] )
    elif sLigne[65:72] == "Oui/Non":
        #print ( sLigne[65:72] )
        AddDicoRubTable( nTable - 1, sLigne[9:64], sLigne[65:70] )
    elif sLigne[65:75] == "Date/Heure":
        #print ( sLigne[65:75] )
        AddDicoRubTable( nTable - 1, sLigne[9:64], sLigne[65:70] )
        
    #print('=>' + sLigne +'<=')
    nLigne = nLigne + 1

#print( "Nombre de ligne : " + str(nLigne) )
AddDicoNomTable( sLigne[8:50], nTable )
AddDicoNbTable( nTable - 1 )

fO.close()

#print( dDicoTable )

#   =========================
#   Mise en forme restitution
#   =========================
sNbTableDico = dDicoTable["nombredetable:"]
#print( "Nb Table MEF : " + sNbTableDico )

#
#   Debut programme VBA pour Access
#
print ( 'Function ExportCLIPS()' )
print ( '   \'' )
print ( '   \'  Generer par acc2clp.py ' + __version__ )
print ( '   \'' )
print ( '   Open "Export Base de Fait.clp" For Output as #1' )
print ( '' )
print ( '   Set DbAccess = CurrentDb' )
print ( '' )

#
#   Pour chaque table
#
t = 0
while ( t < int(sNbTableDico)+1 ):
    sKeyTbl = "table:" + str(t) + ":tbl"
    sKeyDef = "table:" + str(t) + ":def"               
    sTheTableTbl = dDicoTable[sKeyTbl]
    sTheTableDef = dDicoTable[sKeyDef]               
    #print( '>' + str(t) +"<>" + sTheTableTbl + '<' )
    #print( '>' + str(t) +"<>" + sTheTableDef + '<' )

    #
    #       deftemplate table  
    #
    print ( '   Print #1, "(deftemplate ' + sTheTableDef + '"' ) 

    #
    #   Pour chaque rubrique
    #
    sKey = "table:" + str(t)+":rub:nbe"
    nTheRubNbe = dDicoTable[sKey]
    r = 1
    while ( r < int(nTheRubNbe)):
        sKeyRub = "table:" + str(t) + ":rub:" + str(r)
        sKeyTyp = "table:" + str(t) + ":typ:" + str(r)
        sTheTableRub = dDicoTable[sKeyRub]  
        sTheTableTyp = dDicoTable[sKeyTyp]   		
        #print( '>' + str(t)+ " : " + str(r) +"<>" + sTheTableRub + '>=<' + sTheTableTyp + '<' )

        #
        #   slot rubrique
        #
        print ( '   Print #1, "     (slot ' + sTheTableRub + ' (type ' + MefTypDefTemplate(sTheTableTyp) + ') (default ?NONE))"' ) 
        
        r = r + 1
    
    t = t + 1

    #
    #   fin deftemplate
    #
    print ( '   Print #1, ")"' )
    print ( '   Print #1, ""'  )

#   =========================
#   Constitution Base de Fait
#   =========================
print( '   Print #1, "(deffacts faits-initiaux"' )

#
#   Pour chaque table
#
t = 0
while ( t < int(sNbTableDico)+1 ):
    sKeyTbl = "table:" + str(t) + ":tbl"
    sKeyDef = "table:" + str(t) + ":def"               
    sTheTableTbl = dDicoTable[sKeyTbl]
    sTheTableDef = dDicoTable[sKeyDef]               
    #print( '>' + str(t) +"<>" + sTheTableTbl + '<' )
    #print( '>' + str(t) +"<>" + sTheTableDef + '<' )

    #
    #       Parcouris la table  
    #
    print ( '   Print #1, ""' )
    print ( '   Set RsTable = DbAccess.OpenRecordSet("' + sTheTableTbl + '", dbopentable )' )
    print ( '   While Not RsTable.EOF' )
    print ( '       Print #1, "     ( ' + sTheTableDef ) 

    #
    #   Pour chaque rubrique
    #
    sKey = "table:" + str(t)+":rub:nbe"
    nTheRubNbe = dDicoTable[sKey]
    r = 1
    while ( r < int(nTheRubNbe)):
        sKeyRub = "table:" + str(t) + ":rub:" + str(r)
        sKeyTyp = "table:" + str(t) + ":typ:" + str(r)
        sTheTableRub = dDicoTable[sKeyRub]  
        sTheTableTyp = dDicoTable[sKeyTyp]   		
        #print( '>' + str(t)+ " : " + str(r) +"<>" + sTheTableRub + '>=<' + sTheTableTyp + '<' )

        #
        #   slot rubrique
        #
        if sTheTableTyp == "texte":
            print ( '       Print #1, "         ( ' + sTheTableRub + ' " + Chr(34) + RsTable.Fields(' + str(r-1) + ').Value + Chr(34) + ")"' ) 
        else:
            print ( '       Print #1, "         ( ' + sTheTableRub + ' " + Format(RsTable.Fields(' + str(r-1) + ').Value, "0") + ")"' ) 
        
        r = r + 1
    
    t = t + 1

    #
    #   fin deftemplate
    #
    print ( '       Print #1, "     )"' )
    print ( '       RsTable.MoveNext' )
    print ( '   Wend' )
    print ( '   RsTable.Close' )
    print ( '' )
    
#
#   Fin Programme VBA
#
print ( '   Print #1, ")"' )
print ( '' )
print ( '   Close #1' )
print ( '' )
print ( 'End Function' )

#print( dDicoTable )

#
#   Fin acc2clp.py
#
