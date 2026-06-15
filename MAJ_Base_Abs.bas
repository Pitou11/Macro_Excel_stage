Attribute VB_Name = "Module1"
Sub MajSignificationAbs(nomFeuilleSource As String, nomTableauSource As String)
'
'   Programme Macro pour mettre à jour la signification de chaque
' code d'abscence en fonction des informations présentes pour chaque filiale (n'utilisent pas toutes tous les codes).
'   Cette fonction n'a pas à être utilisée directement, elle est appelé par un UserForm afin d'en avoir une utilisation
' la plus simple possible et d'éviter les erreurs de recopiage.
'   PARAM:
'       nomFeuilleSource (string) : nom de la feuille dans laquelle on veut récuperer un tableau
'       nomTableauSource (string) : nom du tableau contenant les colonnes "Code" et "Signification" (ATTENTION ces colonnes doivent existées pour que le programme fonctionne)
'
'   ACTION: met à jour le tableau Base_ABS (ajout de nouveaux codes ou de significations). Si la signification n'est pas encore trouvée "Code non trouvé" est écrit dans la case correspondante
'
' last modif : 02/06/26 (création du macro et optimisation en utilisant un dictionnaire pour effectuer la recherche et la mise à jour en dehors d'Excel même)
' Léonor Martin-Nourry

    ' ####Optimisation 1####
    Application.ScreenUpdating = False ' Pour éviter rafraichissement de la page
    Application.Calculation = xlCalculationManual ' Stoppe le calcul des formules
    Application.EnableEvents = False ' Bloque les autres événements Excel
    
    'Début code
    Dim f1 As Worksheet, f2 As Worksheet
    Dim tbl1 As ListObject, tbl2 As ListObject
    Dim colSignification As ListColumn
    Dim i As Long
    Dim codesF2 As Variant, signifsF2 As Variant
    Dim codesF1 As Variant, res() As Variant
    Dim signifsF1 As Variant
    Dim codeRecherche As String
    Dim cellule As Range
    Dim nouvelleLigne As ListRow
    Dim cle As Variant
    Dim dico As Object
    
    Set dico = CreateObject("Scripting.Dictionary")
    
    Set f1 = ThisWorkbook.Sheets("Feuil1")
    Set f2 = ThisWorkbook.Sheets(nomFeuilleSource) ' La feuille choisie
    Set tbl1 = f1.ListObjects("BASE_ABS")
    Set tbl2 = f2.ListObjects(nomTableauSource)    ' Le tableau choisi
    
    ' Permet de ne pas bloquer si la colonne Signification n'existe pas
    On Error Resume Next
    Set colSignification = tbl1.ListColumns("Signification")
    On Error GoTo 0
    If colSignification Is Nothing Then
        Set colSignification = tbl1.ListColumns.Add
        colSignification.Name = "Signification"
    End If
    
    '==============================
    'optimisation avec utilisation de la ram
    '==============================
    
    codesF2 = tbl2.ListColumns("Code").DataBodyRange.Value
    signifsF2 = tbl2.ListColumns("Signification").DataBodyRange.Value
    
    For i = 1 To UBound(codesF2, 1) 'concrètement dico=tableau Feuil2
        If Not IsEmpty(codesF2(i, 1)) Then
            dico(CStr(codesF2(i, 1))) = signifsF2(i, 1)
        End If
    Next i
    
    codesF1 = tbl1.ListColumns("Code ABS").DataBodyRange.Value
    signifsF1 = colSignification.DataBodyRange.Value
    ReDim res(1 To UBound(codesF1, 1), 1 To 1) 'res a la même taille que codesF1
    
    For i = 1 To UBound(codesF1, 1) 'boucle basique de recherche
        codeRecherche = CStr(codesF1(i, 1))
        If dico.Exists(codeRecherche) Then
            res(i, 1) = dico(codeRecherche)
            dico.Remove codeRecherche
        Else
            If Not IsEmpty(signifsF1(i, 1)) And signifsF1(i, 1) <> "" And signifsF1(i, 1) <> "Code non trouvé" Then
                res(i, 1) = signifsF1(i, 1) ' On garde l'ancienne valeur
            Else
                res(i, 1) = "Code non trouvé" ' Vraiment inconnu
            End If
        End If
    Next i
    
    colSignification.DataBodyRange.Value = res
    '===
    'Ajout des nouveaux codes
    '===
    For Each cle In dico.keys
        Set nouvelleLigne = tbl1.ListRows.Add
        tbl1.ListColumns("Code ABS").DataBodyRange(nouvelleLigne.Index).Value = cle
        colSignification.DataBodyRange(nouvelleLigne.Index).Value = dico(cle)
    Next cle
    
    'Esthétisme
    colSignification.DataBodyRange.Font.Color = vbBlack
    For Each cellule In colSignification.DataBodyRange
        If cellule.Value = "Code non trouvé" Then
            cellule.Font.Color = vbRed
        End If
    Next cellule
    colSignification.Range.EntireColumn.AutoFit
    
        ' ####Fin Optimisation####
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Les significations ont été mises à jour avec succès !", vbInformation, "Terminé"
    
End Sub

Sub OuvrirMenuMaj()
    UserForm1.Show
End Sub
