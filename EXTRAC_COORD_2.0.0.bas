Attribute VB_Name = "EXTRAC_COORD"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' source file for Function extractionDeCoordonees
'++----------------------------------------------------------------++
'## Auteur
'  +-- Laurion Nicolas
'  +-- 23.08.2018
'  +-- GNU licence
'  +-- Version 2.0.0
'++----------------------------------------------------------------++
'
'
'## [Description]
'
'-- Fonction permettant d'isoler deux coordonée (X Y) donnée en parametre avec une syntax spécifique.
'
'********************************************************************************************
'
'## [Syntax du texte traiter par la fonction]
'
'-- TEXT (xxxxxx.xxx yyyyyy.yyy, xxxxxx.xxx yyyyyy.yyy, xxxxxx.xxx yyyyyy.yyy)
'-- Les valeur de coordonée ne doivent pas impérativement avoir un nombre fixe de caractères
'-- dans l'exemple ci-dessus il y a 10 caractères incluant le DOT mais il peux il y en avoir beaucoup
'-- plus ou moins tant que la syntax est respectée
'-- Chaque paire de coordonée doit être séparée par une virgule
'-- Chaque valeur de coordonée au sein d'une paire doivent être séparée par un espace
'
'********************************************************************************************
'
'## [Paramètre]
'
'-- selectionnerLaCellule = Texte a traité au format string
'-- -- paramètre obligatoire
'
'-- extraireY_OU_X  = Au format numérique 0 OR 1
'-- --Ce paramètre est optionnel, si il est omis, le programme utlisera la valeur par défaut (getY) AKA 0
'
'-- paireDeMesuresAExtraire = dépend du nombre de mesure coordonnée disponible dans le texte passer en paramètre
'-- -- l'utilisateur doit savoir quelle paire de coordonée il veut isoler, si le parametre passé a la fonction
'-- -- dépasse le nombre de paire de coordonée trouvée reset paireDeMesuresAExtraire to default value AKA 0
'********************************************************************************************
'
'
'## [Fonctionnement & Utilisation]
'
'Ex 1 : LINESTRING (560885.873 131836.226, 560889.916 131841.012, 560897.095 131843.796)
'
'-- Le programme va supprimer les parenthèse ouverte / fermée plus le mot contenu dans la table
'----denominationKeyword qui est declarée dans la fonction extractionDeCoordonees
'
'-- Le programme va ensuite isoler toute les paire séparée par une virgule puis enlever les espaces inutile
'
'-- Le programme va ensuite choisir une paire de coordonée choisie par l'utilisateur via le paramètre <<paireDeMesuresAExtraire>>
'
'!--[résultat intermédiaire] : disons... L'utilisateur à choisi paireDeMesuresAExtraire = 0 (Cet-a dire la première paire de coordonées)
'!--[résultat intermédiaire] : 560885.873 131836.226
'
'-- Le programme va ensuite choisir entre la coordonée Y ou X en fonction du choix de l'utilisateur via le paramètre extraireY_OU_X
'--  getY = 0 (Valeur par défaut)
'--  getX = 1
'
'-- Le programme retourne la valeur isolée puis se termine.
'
'********************************************************************************************
'
'
'## [ChangeLog]
'
'
' V 1.0.0
' V 2.0.0
'********************************************************************************************
'
'
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Option Explicit
Option Base 1

Public Enum getChoice
    getY = 0
    getX = 1
End Enum

' # [Configurable Global scope variable declaration]

'[Syntax constant]
'You can modify these for changing the syntax used for text processing
'******************************************************************
Public Const openParenthesis As String = "("
Public Const closeParenthesis As String = ")"
Public Const peersDelimiter As String = ","
Public Const coordinateValueDelimiter As String = " "
'******************************************************************


' # [Not configurable Global scope variable declaration]
Public denominationKeyword
Public denominationKeywordSize
Public keywordArrayIsInitialized As Boolean


Sub initKeywordArray()
    'Ajouter des mot clef à supprimer de la chaine de caractere
    denominationKeyword = Array("LINESTRING", "COMPOUNDCURVE", "CIRCULARSTRING")
    
    denominationKeywordSize = UBound(denominationKeyword)
    keywordArrayIsInitialized = True
End Sub


Function extractionDeCoordonees(selectionnerLaCellule As String, _
                Optional extraireY_OU_X As getChoice = getY, _
                Optional paireDeMesuresAExtraire As Integer = 0) As Variant
Attribute extractionDeCoordonees.VB_Description = "Permet l'extraction de coordonées contenue dans un texte, suivant une syntax definie."
Attribute extractionDeCoordonees.VB_ProcData.VB_Invoke_Func = " \n19"
                
        On Error GoTo err
        Dim stringSize As Integer, splittedStringArray, splittedCoordinateArray
        stringSize = Len(selectionnerLaCellule)
        
        'initialize keyword table if flag = false
        If Not keywordArrayIsInitialized Then Call initKeywordArray
        'si le parametre depasse la selection possible EX: 0 ou 1
        If extraireY_OU_X > getX Then extractionDeCoordonees = CVErr(xlErrValue)

        ' if input is empty stop here and return error value
        If stringSize < 1 Then extractionDeCoordonees = CVErr(xlErrValue)
        
        'delete open parenthesis
        selectionnerLaCellule = Replace(selectionnerLaCellule, openParenthesis, "")
        'delete close parenthesis
        selectionnerLaCellule = Replace(selectionnerLaCellule, closeParenthesis, "")
        
        Dim pos As Integer
        'delete keyword from string if it is found
        For pos = 1 To denominationKeywordSize Step 1
            If InStr(selectionnerLaCellule, denominationKeyword(pos)) > 0 Then
                selectionnerLaCellule = Replace(selectionnerLaCellule, denominationKeyword(pos), "")
            End If
        Next
        
        'Split string in multiple subString
        splittedStringArray = Split(selectionnerLaCellule, peersDelimiter)
        
        'define counter var
        Dim s, counter
        s = UBound(splittedStringArray)
        
        'If size of splittedStringArray is smaller than 1 return error value
        'this mean that split doesn't return a value
        If s < 1 Then extractionDeCoordonees = CVErr(xlErrValue)
        
        ' Loop for removing unexpected space caracter without removing space between the two value
        ' this is do with Trim for each value of the array

        'enlever cette loop for pour faire un trim seulement sur la paire selectionnée par paireDeMesureAExtraire'
        For counter = 0 To s Step 1
            splittedStringArray(counter) = Trim(splittedStringArray(counter))
        Next
        
        ' check if the size of paireDeMesuresAExtraire is not bigger than the size of splittedStringArray
        'then return error value OR if uncommented set paireDeMesuresAExtraire to 0 AKA default value
        
        'If paireDeMesuresAExtraire > s Or paireDeMesuresAExtraire < 0 Then extractionDeCoordonees = CVErr(xlErrValue)
        If paireDeMesuresAExtraire > s Or paireDeMesuresAExtraire < 0 Then paireDeMesuresAExtraire = 0
        
        'split string with space for delimitation
        'faire le trim de la loop for ici sur splittedStringArray(paireDeMesureAExtraire)'
        splittedCoordinateArray = Split(splittedStringArray(paireDeMesuresAExtraire), coordinateValueDelimiter)
        
        ' if the size of splittedCoordinateArray is smaller than 1 this mean that split doesn't return a value
        If UBound(splittedCoordinateArray) < 1 Then
            extractionDeCoordonees = CVErr(xlErrValue)
        Else
            extractionDeCoordonees = Format(splittedCoordinateArray(extraireY_OU_X), "0.000")
            'extractionDeCoordonees = splittedCoordinateArray(extraireY_OU_X)
        End If
EndOfFunction:
    Exit Function
err:
    MsgBox Error$, vbCritical + vbOKOnly
    extractionDeCoordonees = CVErr(xlErrValue)
    Resume EndOfFunction
        
End Function

Public Sub DefineFunction_extractionDeCoordonees()
    Dim sFunctionName As String
    Dim sFunctionCategory As String
    Dim sFunctionDescription As String
    Dim aFunctionArguments(1 To 3) As String

    sFunctionName = "extractionDeCoordonees"
    sFunctionDescription = "Permet l'extraction de coordonées contenue dans un texte, suivant une syntax definie."
    sFunctionCategory = "Extraction de coordonnées"
    
    aFunctionArguments(1) = "Sélectionner la cellule contenant le text à analyser"
    
    aFunctionArguments(2) = "[Paramètre Optionnel !]" + vbCrLf + _
                            "Par défaut la coordonnée Y est extraite." + vbCrLf + _
                            "Pour extraire la coordonnée X veuillez saisir 1"

    aFunctionArguments(3) = "[Paramètre Optionnel !]" + vbCrLf + _
                            "Quelle paire de coordonnées voulez-vous extraire ? " + vbCrLf + _
                            "par défaut la première  sinon 1 = 2 ème paires etc.."

    Application.MacroOptions Macro:=sFunctionName, _
         Description:=sFunctionDescription, _
         Category:=sFunctionCategory, _
         ArgumentDescriptions:=aFunctionArguments
End Sub

