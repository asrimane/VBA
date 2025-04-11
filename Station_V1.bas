Attribute VB_Name = "Module1"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
 ' source file for Macro Excel RE Saisie de terre Station
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
'Ensemble de macro, pour saisir efficacement les donnée d'impedence de Terre des Station
'Transformatrice de la Romande Energie.
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
'lors de la saisie des valeur d'impédence l'utilisateur peut rentrer le caractère "/" lorsqu'il n'y a pas de valeur à saisir
'l'utilisateur peut aussi utiliser le raccourci ".60" à la place de "0.60"
'
'Lors de la saisie des question <<Conformité de l'impédance de contact>>
'l'utilisateur peut saisir les caractere ci-dessous pour valider l'entrée,
'le programme saisi automatiquement <<Oui>> ; <<Non>> ; <<Pas mesuré>>
'
'[Oui]
'-- "Oui", "oui", "OUI", "o", "O", "1"
'
'[Non]
'-- "Non", "non", "NON", "n", "N", "0"
'
'[Pas mesuré]
'-- Appuyer simplement sur <<ENTER>>
'-- Ou tout autre caractères qui n'est pas dans la liste au dessus
'
'
'Lors de la saisie des question <<La mise à terre est-elle mesurable ? >>
'l'utilisateur peut saisir les caractere ci-dessous pour valider l'entrée,
'le programme saisi automatiquement <<Oui>> ; <<Non>> ;
'
'[Oui]
'-- "Oui", "oui", "OUI", "o", "O", "1"
'
'[Non]
'-- "Non", "non", "NON", "n", "N", "0"
'-- Appuyer simplement sur <<ENTER>>
'-- Ou tout autre caractères qui n'est pas dans la liste positive au dessus.
'
'
'Les variantes des macro avec le nom [stricte] signifie que le programme
'cherchera exactement la chaîne de caracère donne en paramètre,
'EX: L'utilisateur entre "mont" alors qu'il cherche montreux
'le programme cherchera uniquement les resultat étant égual à la chaine "mont"
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
' V 1.0.0 Nicolas test 13.08.2018 saisie fonctionnel
'********************************************************************************************
'
'
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$





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
    Do Until IsDate(userDate) And Len(userDate) = 10
        userDate = InputBox("Saisir la date de mesure JJ.MM.AAAA s.v.p")
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
        loopCount = loopCount + 1
    Loop
    updateDateFlag = False
End Sub


'*********************************
'* doFilterStationName           *
'*********************************
'*
Sub doFilterStationName()
Attribute doFilterStationName.VB_ProcData.VB_Invoke_Func = "t\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la station : ", "Station name")         'ask input
    If Not IsEmpty(strInput) Then
        If Len(strInput) > 0 Then
            Cells(1, 2).AutoFilter Field:=2, Criteria1:="=*" & strInput & "*"   'set filter
        End If
    End If
End Sub


'*********************************
'* doFilterStationNameStrict     *
'*********************************
'*
Sub doFilterStationNameStrict()
Attribute doFilterStationNameStrict.VB_ProcData.VB_Invoke_Func = "z\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la station [STRICT] : ", "Station name [STRICT]")
    If Not IsEmpty(strInput) Then
        If Len(strInput) > 0 Then
            Cells(1, 2).AutoFilter Field:=2, Criteria1:=strInput
        End If
    End If
End Sub


'*********************************
'* doFilterCommuneName           *
'*********************************
'*
Sub doFilterCommuneName()
Attribute doFilterCommuneName.VB_ProcData.VB_Invoke_Func = "o\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la Commune : ", "Commune name")
    If Not IsEmpty(strInput) Then
        If Len(strInput) > 0 Then
            Cells(1, 6).AutoFilter Field:=6, Criteria1:="=*" & strInput & "*"
        End If
    End If
End Sub


'*********************************
'* doFilterCommuneNameStrict     *
'*********************************
'*
Sub doFilterCommuneNameStrict()
Attribute doFilterCommuneNameStrict.VB_ProcData.VB_Invoke_Func = "p\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la Commune [STRICT] : ", "Commune name [STRICT]")
    If Not IsEmpty(strInput) Then
        If Len(strInput) > 0 Then
            Cells(1, 6).AutoFilter Field:=6, Criteria1:=strInput
        End If
    End If
End Sub


'*********************************
'* doFilterLocaliteName          *
'*********************************
'*
Sub doFilterLocaliteName()
Attribute doFilterLocaliteName.VB_ProcData.VB_Invoke_Func = "u\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la localité : ", "localité name")
    If Not IsEmpty(strInput) Then
        If Len(strInput) > 0 Then
            Cells(1, 5).AutoFilter Field:=5, Criteria1:="=*" & strInput & "*"
        End If
    End If
End Sub


'*********************************
'* doFilterLocaliteNameStrict    *
'*********************************
'*
Sub doFilterLocaliteNameStrict()
Attribute doFilterLocaliteNameStrict.VB_ProcData.VB_Invoke_Func = "i\n14"
    'Call ResetFilters
    strInput = InputBox("Nom de la localité [STRICT] : ", "localité name [STRICT]")
    If Not IsEmpty(strInput) Then
        If Len(strInput) > 0 Then
            Cells(1, 5).AutoFilter Field:=5, Criteria1:=strInput
        End If
    End If
End Sub


'*********************************
'* doFilterDualCriteria          *
'*********************************
'*
Sub doFilterDualCriteria()
Attribute doFilterDualCriteria.VB_ProcData.VB_Invoke_Func = "q\n14"
    Call ResetFilters
    Call doFilterStationName
    Call doFilterCommuneName
End Sub


'*********************************
'* updateValue                   *
'*********************************
'*
Sub updateValue()
Attribute updateValue.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim DrLig As Long, Lig As Long, vStart As Long, vStop As Long
    Dim mesureName As Variant, counter As Long, flag As Boolean, dateIn As Variant
    
    mesureName = Array("Terre Générale", "Terre Séparée", "Terre Pontée", "Conformité de l'impédance de contact", "La mise à terre est-elle mesurable ?")
    
    flag = False
    counter = 0
    nameCount = 0
    
    If updateDateFlag Then Call updateDate

    DrLig = Cells(Rows.Count, 1).End(xlUp).Row
    For Lig = 2 To DrLig
        If Not Rows(Lig).Hidden Then
            If Not flag Then
                vStart = Lig
                flag = True
            Else
                counter = counter + 1
            End If
        End If
    Next
    vStop = vStart + counter
    res = vStop - vStart
    flag = False
    If res > 4 Then
        MsgBox "Erreur la macro updateValue a été utiliser " + _
                "sans effectuer de recherche sur la station/commune " + _
                "il y a plus de cing lignes", vbCritical + vbOKOnly
        Exit Sub
    End If
    For s = vStart To vStop Step 1
        flag = False
        inData = InputBox("Saisir " + mesureName(nameCount) + " : " + vbCrLf + "Tapper qqq  pour quitter la saisie.")
        Select Case nameCount
            Case 0, 1, 2
                If inData = "qqq" Or inData = "QQQ" Then
                    Exit Sub
                ElseIf Not inData = "/" Then
                    If Not IsNumeric(inData) Then
                        MsgBox "La valeur saisie doit être un nombre", vbCritical + vbOKOnly, "Erreur de saisie"
                        s = s - 1
                        nameCount = nameCount - 1
                        inData = ""
                        flag = True
                    End If
                Else
                    inData = ""
                End If
            Case 3
                Select Case inData
                
                Case "Oui", "oui", "OUI", "o", "O", "1"
                    inData = "Oui"
                Case "Non", "non", "NON", "n", "N", "0"
                    inData = "Non"
                Case Else
                    inData = "Pas mesuré"
                End Select
            Case 4
                Select Case inData
                    Case "Oui", "oui", "OUI", "o", "O", "1"
                        inData = "Oui"
                    Case "Non", "non", "NON", "n", "N", "0"
                        inData = "Non"
                    Case Else
                        inData = "Non"
                End Select
        End Select
        nameCount = nameCount + 1
        If nameCount > 4 Then nameCount = 0
        If Not flag Then
            Cells(s, 11).Value = inData
            If IsNumeric(userDate) Then
                If Not userDate > 0 Then Call updateDate
            ElseIf userDate = vbEmpty Then
                Call updateDate
            End If
            Cells(s, 13).Value = userDate
            Cells(s, 14).Value = "X"
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
       'This is if you place the macro in your personal wb to be able to reset the filters on any wb you're currently working on. Remove the set wb = thisworkbook if that's what you need
            For Each ws In wb.Worksheets
                If ws.FilterMode Then
                    ws.ShowAllData
                Else
                End If
                'This removes "normal" filters in the workbook - however, it doesn't remove table filters
                For Each listObj In ws.ListObjects
                    If listObj.ShowHeaders Then
                        listObj.AutoFilter.ShowAllData
                        listObj.Sort.SortFields.Clear
                    End If
                Next listObj
            Next
'And this removes table filters. You need both aspects to make it work.
End Sub


Sub reset_X_flag_toNothing()
Attribute reset_X_flag_toNothing.VB_ProcData.VB_Invoke_Func = "s\n14"
    Dim DrLig As Long, Lig As Long, vStart As Long, vStop As Long
    Dim counter As Long, flag As Boolean
    DrLig = Cells(Rows.Count, 1).End(xlUp).Row
    For Lig = 2 To DrLig
        If Not Rows(Lig).Hidden Then
            If Not flag Then
                vStart = Lig
                flag = True
            Else
                counter = counter + 1
            End If
        End If
    Next
    vStop = vStart + counter
    res = vStop - vStart
    flag = False
    If res > 4 Then
        MsgBox "Erreur la macro reset_X_flag_toNothing a été utiliser " + _
                "sans effectuer de recherche sur la station/commune " + _
                "il y a plus de cing lignes", vbCritical + vbOKOnly
        Exit Sub
    End If
    For s = vStart To vStop Step 1
        Cells(s, 14).Value = ""
    Next
End Sub

Sub reset_dateCells_toNothing()
Attribute reset_dateCells_toNothing.VB_ProcData.VB_Invoke_Func = "a\n14"
    Dim DrLig As Long, Lig As Long, vStart As Long, vStop As Long
    Dim counter As Long, flag As Boolean
    DrLig = Cells(Rows.Count, 1).End(xlUp).Row
    For Lig = 2 To DrLig
        If Not Rows(Lig).Hidden Then
            If Not flag Then
                vStart = Lig
                flag = True
            Else
                counter = counter + 1
            End If
        End If
    Next
    vStop = vStart + counter
    res = vStop - vStart
    flag = False
    If res > 4 Then
        MsgBox "Erreur la macro reset_dateCells_toNothing a été utiliser " + _
                "sans effectuer de recherche sur la station/commune " + _
                "il y a plus de cing lignes", vbCritical + vbOKOnly
        Exit Sub
    End If
    For s = vStart To vStop Step 1
        Cells(s, 13).Value = ""
    Next
End Sub

Sub reset_All_with_given_date()
Attribute reset_All_with_given_date.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim DrLig As Long, Lig As Long, vStart As Long, vStop As Long
    Dim counter As Long, flag As Boolean
    DrLig = Cells(Rows.Count, 1).End(xlUp).Row
    For Lig = 2 To DrLig
        If Not Rows(Lig).Hidden Then
            If Not flag Then
                vStart = Lig
                flag = True
            Else
                counter = counter + 1
            End If
        End If
    Next
    vStop = vStart + counter
    res = vStop - vStart
    flag = False
    If res > 4 Then
        MsgBox "Erreur la macro reset_All_with_given_date a été utiliser " + _
                "sans effectuer de recherche sur la station/commune " + _
                "il y a plus de cing lignes", vbCritical + vbOKOnly
        Exit Sub
    End If
    For s = vStart To vStop Step 1
        Cells(s, 11).Value = ""
        Cells(s, 13).Value = ""
        Cells(s, 14).Value = ""
    Next
End Sub


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

