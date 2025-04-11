Attribute VB_Name = "Module1"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' source file for Macro Excel RE Transfert des mesures DP cellules et transfos
'++----------------------------------------------------------------++
'## Auteur
'  +-- Laurion Nicolas
'  +-- 13.08.2018
'  +-- GNU licence
'  +-- Version 1.0.0
'++----------------------------------------------------------------++
'
'
'## [Description]
'
'********************************************************************************************
'Ensemble de macro, pour saisir les donnée Transfert des mesures DP cellules et transfos
'
'
'L'utilisateur via des raccourci clavier ouvre des boîte de dialogue lui
'permettant de saisir les données, demandée par le programme
'********************************************************************************************
'
'
'
'## [Usage]
'
'********************************************************************************************
'Le programme demande automatiquement la date de mesure pour les saisie qui
'vont-être faites, si celle-ci n'est pas definie au préalable.
'L'utilisateur peut executer la macro <<updateDate>> pour mettre à jour cette information.
'
'
'
'Quand l'utilisateur passe à un autre groupe de station avec une date de saisie différente
'il doit executer la macro <<updateDate>> pour mettre à jour cette information.
'
'Le programme saisi automatique la date lors de la modification d'une ligne.
'
'Le programme valide automatiquement la case <<Traité?>> avec un "X"
'
'
'lors de la saisie des valeur :
'- l'utilisateur peut rentrer le caractère "/" lorsqu'il n'y a pas de valeur à saisir
'- l'utilisateur peut utiliser le raccourci ".60" à la place de "0.60"
'
'
'Les variantes des macro avec le nom [stricte] signifie que le programme
'cherchera exactement la chaîne de caracère donne en paramètre.
'
'EX: L'utilisateur entre "mont" alors qu'il cherche montreux
'   le programme cherchera uniquement les resultat étant égual à la chaine "mont"
'
'
'Pour faire une recherche avec seulement une partie du mot veuillez utilisez
'les Macro NON stricte
'
'
'
'********************************************************************************************
'
'## [ChangeLog]
'
'
' V 1.0.0
' V 1.0.1 Test nicolas parfaitement fonctionnel 15.08.2018
'********************************************************************************************
'
'
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'---------------------------------
'Configuration
'---------------------------------
Public Const max_EmplacementDeLaMesure As Integer = 45 'Maximum de champ à remplir par feuille, normalement il n'y en a que ~20 mais peut y'en avoir plus
'---------------------------------

'---------------------------------
'Global scope variable declaration
'---------------------------------
Public userDate
Public updateDateFlag As Boolean
'---------------------------------


'*********************************
'* updateDate    *
'*********************************
'*
Sub updateDate()
Attribute updateDate.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim loopCount As Integer, oldUserDate As String
    loopCount = 0
    oldUserDate = userDate 'Backup old value if user cancel
    userDate = ""
    'Faire jusqu'a ce que la saisie de l'utilisateur soit une date valide
    '
    Do Until IsDate(userDate) And Len(userDate) = 10
        userDate = InputBox("Saisir la date de mesure JJ.MM.AAAA s.v.p" + vbCrLf + "Tapper qqq  pour quitter la saisie.")
        'if user want to quit
        If userDate = "qqq" Or userDate = "QQQ" Then
            userDate = oldUserDate 'set to old value
            Exit Sub
        End If
        If loopCount > 1 Then 'if user enter wrong value 2 time
            Select Case MsgBox("Erreur, la date a été saisie " + "de façon erronée," + vbCrLf + _
                      "Voulez-vous retenter de saisir la valeur demandée ?", _
                       vbExclamation + vbRetryCancel, "Erreur de saisie")
                   Case vbRetry
                        loopCount = 0   'reset var
                        userDate = ""   'reset var
                   Case vbCancel
                        userDate = oldUserDate 'Reasign old value beacause user as canceld
                        Exit Sub
                   Case Else
                        MsgBox "Erreur non gerée", vbCritical + vbOKOnly
                        Exit Sub
            End Select
        End If
        'Incremente invalid attempt counter
        loopCount = loopCount + 1
    Loop
    updateDateFlag = False
End Sub


'*********************************
'* doFilter_LieuDit              *
'*********************************
'*
Sub doFilter_LieuDit()
Attribute doFilter_LieuDit.VB_ProcData.VB_Invoke_Func = "t\n14"
    'Call ResetFilters
    strInput = InputBox("Lieu-dit : ", "Lieu-dit / emplacement")
        If Len(strInput) > 0 Then
            Cells(1, 2).AutoFilter Field:=2, Criteria1:="=*" & strInput & "*"
        End If
    ActiveWindow.ScrollRow = 1
End Sub


'*********************************
'* doFilter_LieuDit_Strict       *
'*********************************
'*
Sub doFilter_LieuDit_Strict()
Attribute doFilter_LieuDit_Strict.VB_ProcData.VB_Invoke_Func = "z\n14"
    'Call ResetFilters
    strInput = InputBox("Lieu-dit [STRICT] : ", "Lieu-dit / emplacement [STRICT]")
        If Len(strInput) > 0 Then
            Cells(1, 2).AutoFilter Field:=2, Criteria1:=strInput
        End If
    ActiveWindow.ScrollRow = 1
End Sub


'*********************************
'* doFilter_CommuneName           *
'*********************************
'*
Sub doFilter_CommuneName()
Attribute doFilter_CommuneName.VB_ProcData.VB_Invoke_Func = "u\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la Commune : ", "Commune name")
        If Len(strInput) > 0 Then
            Cells(1, 1).AutoFilter Field:=1, Criteria1:="=*" & strInput & "*"
        End If
    ActiveWindow.ScrollRow = 1
End Sub


'*********************************
'* doFilter_CommuneNameStrict     *
'*********************************
'*
Sub doFilter_CommuneNameStrict()
Attribute doFilter_CommuneNameStrict.VB_ProcData.VB_Invoke_Func = "i\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la Commune [STRICT] : ", "Commune name [STRICT]")
        If Len(strInput) > 0 Then
            Cells(1, 1).AutoFilter Field:=1, Criteria1:=strInput
        End If
    ActiveWindow.ScrollRow = 1
End Sub


'*********************************
'* doFilter_PosteTechnique       *
'*********************************
'*
Sub doFilter_PosteTechnique()
Attribute doFilter_PosteTechnique.VB_ProcData.VB_Invoke_Func = "a\n14"
    'Call ResetFilters
    strInput = InputBox("ID PosteTechnique : ", "ID PosteTechnique")
        If Len(strInput) > 0 Then
            Cells(1, 11).AutoFilter Field:=11, Criteria1:="=*" & strInput & "*"
        End If
    ActiveWindow.ScrollRow = 1
End Sub


'*********************************
'* doFilter_PosteTechniqueStrict *
'*********************************
'*
Sub doFilter_PosteTechniqueStrict()
Attribute doFilter_PosteTechniqueStrict.VB_ProcData.VB_Invoke_Func = "s\n14"
    'Call ResetFilters
    strInput = InputBox("ID PosteTechnique [STRICT] : ", "ID PosteTechnique [STRICT]")
        If Len(strInput) > 0 Then
            Cells(1, 11).AutoFilter Field:=11, Criteria1:=strInput
        End If
    ActiveWindow.ScrollRow = 1
End Sub


'*********************************
'* doFilterDualCriteria          *
'*********************************
'*
Sub doFilterDualCriteria()
Attribute doFilterDualCriteria.VB_ProcData.VB_Invoke_Func = "q\n14"
    Call ResetFilters
    Call doFilter_CommuneName
    Call doFilter_LieuDit
End Sub


'*********************************
'* updateValue                   *
'*********************************
'*
Sub updateValue()
Attribute updateValue.VB_ProcData.VB_Invoke_Func = "r\n14"

    'Déclaration des variable
    Dim lastRow As Long, currentRow As Long
    Dim askEverytimeFor_TEV As Boolean, flag As Boolean
    Dim dateIn As Variant, ultraTEV_Value As String, rowsIdList() As Integer
    Dim arraySize As Integer
    
    'Condition pour mettre a jour les donnée dois etre false (false = flag unlock, true = lock)
    flag = False 'Par default doit etre sur false
    'Condition pour demander le status du TEV doit etre True
    askEverytimeFor_TEV = True
    
    'Si la date n'a pas étée definie...
    If updateDateFlag Then Call updateDate
    
'***************************
'First check for counting hidden rows
'***************************
    ' set var to last row's Index
    lastRow = Cells(Rows.count, 1).End(xlUp).Row
    'currentRow = 2 parce que on ne prend pas le header en position 1
    For currentRow = 2 To lastRow
        'Si la ligne n'est pas cachée
        If Not Rows(currentRow).Hidden Then
            'Incremente le compteur de ligne pour definire plus tard la taille du tableau de ligne filtrée
            arraySize = arraySize + 1
        End If
    Next
'Maintenant on connait le nombre exact de ligne visible
'on peut donc definir la taille du tableau contenant le numero de lligne a traité

'***************************
'Second check for listing visible rows ID and put in array
'***************************
    'Si aucune ligne visible n'a été trouvée dans le FIRST CHECK, Stop and display error message
    If arraySize > 0 Then
        'Size array par rapport au ligne visible filtrée
        ReDim rowsIdList(arraySize - 1)
        'Compteur pour la position du tableau rowsIdList
        Dim count As Integer
        count = 0
        'currentRow = 2 parce que on ne prend pas le header en position 1
        For currentRow = 2 To lastRow
            'Si la ligne n'est pas cachée
            If Not Rows(currentRow).Hidden Then
                'Recupère l'ID de la ligne stocke dans le tableau rowsIdList
                rowsIdList(count) = Rows(currentRow).Row
                'tant que le compteur est plus petit que la taille du tableau
                If count < arraySize Then
                'incremente le compteur
                    count = count + 1
                Else
                   'Quit for loop and continue execution
                   Exit For
                End If
            End If
        Next
    Else
        'Error arraySize < ou = 0
        MsgBox "Pas d'enregistrement à traité !!", vbCritical + vbOKOnly
    End If
    
    'Si le nombre de ligne visible trouvée est plus grand que le maximum de lignes pouvant concerner un rapport de station
    'Display error message and stop execution
    '(Valeur definie dans la partie [config] au début du fichier)
    If arraySize > max_EmplacementDeLaMesure Then
        MsgBox "Erreur la macro updateValue a été utiliser " + _
                "sans effectuer de recherche sur la station/commune " + _
                "il y a plus de " + CStr(max_EmplacementDeLaMesure) + " lignes", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    'Start value update loop
    For s = 0 To arraySize - 1 Step 1
        'Reset variable
        ultraTEV_Value = ""
        flag = False
        
        'Request user input
        inData = InputBox("Saisir " + Cells(rowsIdList(s), 6).Value + " : " + vbCrLf + "Tapper qqq  pour quitter la saisie.")
        
        'Si l'utilisateur saisi qqq OU QQQ alors stop execution
        If inData = "qqq" Or inData = "QQQ" Then
            Exit Sub
        'si inData n'est pas egal à : "/" ou "*" ou "*/off" ou "*/on" ... Alors continue la saisie de donée
        ElseIf Not inData = "/" And Not inData = "*" And Not inData = "*/off" And Not inData = "*/on" Then
            'Est ce que c'est une valeur numerique ?
            If Not IsNumeric(inData) Then
                'Non, display error message
                MsgBox "La valeur saisie doit être un nombre", vbCritical + vbOKOnly, "Erreur de saisie"
                'decremente s pour repasse sur cette question
                s = s - 1
                'reset inData
                inData = ""
                'Set flag true pour de pas rentre dans le statement de mise a jour des valeur.
                flag = True
            Else 'C'est une valeur numérique
            
                ' si sur true demande les info concernant le TEV
                If askEverytimeFor_TEV Then
                    userReturn = MsgBox("Est-ce que ultraTEV à été utilisé pour " + Cells(rowsIdList(s), 6).Value + " ? ", vbInformation + vbYesNo + vbDefaultButton2, "ultraTEV ?")
                    Select Case userReturn
                        Case vbYes 'Oui
                            ultraTEV_Value = "X"
                        Case vbNo 'Non
                            ultraTEV_Value = ""
                        Case Else 'Unset = Non
                            ultraTEV_Value = ""
                    End Select
                End If 'END OF If askEverytimeFor_TEV
            End If 'END OF If Not IsNumeric(inData)
        Else 'ELSE OF if inData = "qqq" Or inData = "QQQ" Then
        
            'Si l'utilisateur saisi un asterisque "*"
            'le programme n'enregistrera aucune valeur pour la ligne en court de traitement
            'sans mettre a jour la date et la case <<Traité?>>
            If inData = "*" Then
                flag = True
                
            'Si l'utilisateur saisi un combo "*/off"
            'le programme desactivera la demande de TEV
            ElseIf inData = "*/off" Then
                'decremente pour repasser sur cette ligne
                s = s - 1
                askEverytimeFor_TEV = False
                'lock update statement
                flag = True
                
            'Si l'utilisateur saisi un combo "*/on"
            'le programme activera la demande de TEV
            ElseIf inData = "*/on" Then
                'decremente pour repasser sur cette ligne
                s = s - 1
                askEverytimeFor_TEV = True
                'lock update statement
                flag = True

            Else 'Autre valeur : normalement impossible de passer dans cette partie
                inData = ""
            End If
        End If 'END OF if inData = "qqq" Or inData = "QQQ" Then
        
        'Si la mise a jour de valeur n'est pas bloquée par flag = true
        If Not flag Then
            'Update colonne <<Valorisation>>
            Cells(rowsIdList(s), 8).Value = inData
            'teste si la date est definie, le cas contraire execute la demande de date
            If IsNumeric(userDate) Then
                If Not userDate > 0 Then Call updateDate
            ElseIf userDate = vbEmpty Then
                Call updateDate
            End If
            'Update colonne <<Date de la mesure>>
            Cells(rowsIdList(s), 7).Value = userDate
            'Update colonne <<Ultra TEV>>
            Cells(rowsIdList(s), 10).Value = ultraTEV_Value
            'Update colonne <<Traité?>>
            Cells(rowsIdList(s), 13).Value = "X"
        End If
    Next
End Sub

'*********************************
'* ResetFilters                  *
'*********************************
'*
Sub ResetFilters()
Attribute ResetFilters.VB_ProcData.VB_Invoke_Func = "w\n14"
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim listObj As ListObject
    Set wb = ThisWorkbook
    'Set wb = ActiveWorkbook
    'This is if you place the macro in your personal wb to be able to reset the filters on any wb you're currently working on.
    'Remove the set wb = thisworkbook if that's what you need
    For Each ws In wb.Worksheets
        If ws.FilterMode Then
            ws.ShowAllData
        Else
        End If
        'This removes "normal" filters in the workbook, it doesn't remove table filters
        For Each listObj In ws.ListObjects
            If listObj.ShowHeaders Then
                listObj.AutoFilter.ShowAllData
                listObj.Sort.SortFields.Clear
            End If
        Next listObj
    Next
End Sub
