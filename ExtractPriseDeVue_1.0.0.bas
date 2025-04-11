Attribute VB_Name = "Module1"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' source file for Macro Excel extraction prise de vue image
'++----------------------------------------------------------------++
'## Auteur
'  +-- Laurion Nicolas
'  +-- Source : https://forum.excel-pratique.com/viewtopic.php?t=97464
'  +-- 23.02.2020
'  +-- GNU licence
'  +-- Version 1.0.0
'++----------------------------------------------------------------++
'
'
'## [Description]
'
'********************************************************************************************
'Macro pour extraire la date de prise de vue d'une image
'
'********************************************************************************************
'
'## [ChangeLog]
'
'
' V 1.0.0
'********************************************************************************************
'
'
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'Configuration:'

'Si sur <<true>> retourne une erreur excel dans la cellule, sur <<false>> retourne 00.01.1900
Public Const returnExcelErrorOnEmptyValue = False
Public Const returnExcelErrorOnFileNotFound = False


Function Extraction_DatePriseDeVue(ByVal path) As Date
    On Error GoTo Err

    
    'Déclaration des variables
    '-----
    Dim Fso As Object, oFichier As Object
    Dim objShell As Shell32.Shell
    Dim objFolder As Shell32.Folder
    Dim strFileName As Shell32.FolderItem
    Dim parentFolderPath As String, filename As String, output As String, extractedDate As Date
    '-----
    
    'Initialisation des objets
    '-----
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    If Fso.FileExists(path) Then
        Set oFichier = Fso.GetFile(path)
        parentFolderPath = Fso.GetParentFolderName(oFichier)
        filename = Fso.GetFileName(oFichier)
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.Namespace(parentFolderPath)
        Set FileDetailsObject = objFolder.Items.Item(filename)
        
        'Récupère la valeur de la prise de vue
        output = objFolder.GetDetailsOf(FileDetailsObject, 12)
        
        'Si il y'a une date
        If Not output = "" Then
            output = Replace(output, ".", "/")
            output = Replace(Replace(output, ChrW(8206), ""), ChrW(8207), "")  'enleve << ? >>
            splited = Split(output, " ")  'separe le temps de la date
            output = splited(0)           'recuperation de la date isolee
            extractedDate = CDate(output) 'conversion de du ype String a Date
        Else
            If returnExcelErrorOnEmptyValue Then
                Extraction_DatePriseDeVue = CVErr(xlErrValue)
                Exit Function
            Else
                Extraction_DatePriseDeVue = vbEmpty  'return vbEmpty AKA 00.01.1900
                Exit Function
            End If
        End If
    Else
        If returnExcelErrorOnFileNotFound Then
            Extraction_DatePriseDeVue = CVErr(xlErrValue)
            Exit Function
        Else
            Extraction_DatePriseDeVue = vbEmpty 'return vbEmpty AKA 00.01.1900
            Exit Function
        End If
    End If
    '-----
    
    Extraction_DatePriseDeVue = extractedDate
    Exit Function
    
Err:
    MsgBox Error$, vbCritical + vbOKOnly
    Extraction_DatePriseDeVue = vbEmpty
    Exit Function
End Function
