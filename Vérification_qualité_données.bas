Attribute VB_Name = "Module1"
Sub Verification_Qualifications_2()
    '
    '   Fonction qui permet de vérifier si les :
    '       - services; -qualifications; -sites; -agence;
    ' présents dans une table de donnée, contenant les informations de vacations des
    ' agents (ex : CDI_CDD), sont bien référencés.
    '   Pour le bon fonctionnement de cette macro :
    '       - renommer les TCD tels que : TCD_services/sites/qualifs/agence
    '       - renommer les tableaux de ref : Ref_services/sites/qualifs/agence
    '       - Ajouter aux TDC une colonne somme de valeur (colonne des noms pour compter les vides)
    '       - Séparer les TCD d'au moins 2 cases
    '       - Remplir les tableaux de ref
    ' Léonor Martin-Nourry
    ' last update : (15/06/26)(ajout vérification agence(intérimaires))
    
    Dim wb As Workbook
    Dim wsRef As Worksheet
    Dim wsPre As Worksheet
    Dim tab_services As ListObject, tab_qualifs As ListObject, tab_sites As ListObject, tab_agence As ListObject
    Dim pt_services As PivotTable, pt_sites As PivotTable, pt_qualifs As PivotTable, pt_agence As PivotTable
    
    Set wb = ActiveWorkbook
    Set wsRef = wb.Sheets("REF")
    Set wsPre = wb.Sheets("Qualité")

    ' ==========================
    ' Définition des Tableaux de référence
    ' ==========================
    'Source
    Set tab_services = wsRef.ListObjects("Ref_services")
    Set tab_sites = wsRef.ListObjects("Ref_sites")
    Set tab_qualifs = wsRef.ListObjects("Ref_qualifs")
    Set tab_agence = wsRef.ListObjects("Ref_agence")
    'Cible
    Set pt_services = wsPre.PivotTables("TCD_services")
    Set pt_sites = wsPre.PivotTables("TCD_sites")
    Set pt_qualifs = wsPre.PivotTables("TCD_qualifs")
    Set pt_agence = wsPre.PivotTables("TCD_agence")

    ' ==========================
    ' Variable FAUX
    ' ==========================
    Dim hasFalse As Boolean
    hasFalse = False

    ' ==========================
    ' Dictionnaires
    ' ==========================
    
    Dim dictService As Object, dictSite As Object, dictQualif As Object, dictAgence As Object
    Set dictService = CreateObject("Scripting.Dictionary")
    Set dictSite = CreateObject("Scripting.Dictionary")
    Set dictQualif = CreateObject("Scripting.Dictionary")
    Set dictAgence = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim valeur As String
    Dim cellule As Range
    
    ' ==========================
    ' REMPLISSAGE DES DICTIONNAIRES VIA LES TABLEAUX
    ' ==========================
    
    ' SERVICES
    
        For i = 1 To tab_services.ListRows.Count
            valeur = tab_services.DataBodyRange(i, 1).Value
            If valeur <> "" Then dictService(valeur) = True
        Next i

    ' SITES
        For i = 1 To tab_sites.ListRows.Count
            valeur = tab_sites.DataBodyRange(i, 1).Value
            If valeur <> "" Then dictSite(valeur) = True
        Next i

    ' QUALIFICATIONS
        For i = 1 To tab_qualifs.ListRows.Count
            valeur = tab_qualifs.DataBodyRange(i, 1).Value
            If valeur <> "" Then dictQualif(valeur) = True
        Next i
        
    ' AGENCE
        For i = 1 To tab_agence.ListRows.Count
            valeur = tab_agence.DataBodyRange(i, 1).Value
            If valeur <> "" Then dictAgence(valeur) = True
        Next i
    
    ' ==========================
    ' VÉRIFICATION DANS LA FEUILLE "Qualité"
    ' ==========================
    
    ' SERVICE
    For Each cellule In pt_services.RowRange
        valeur = cellule.Value
        If valeur <> "" And cellule.Row > pt_services.RowRange.Rows(1).Row Then
            If InStr(1, valeur, "Total") > 0 Then
                cellule.Offset(0, 2).Clear
                cellule.Offset(0, 2).Interior.ColorIndex = xlNone
            Else
                CheckValeur cellule, dictService, hasFalse
            End If
        End If
    Next cellule
        
    ' SITE
    For Each cellule In pt_sites.RowRange
        valeur = cellule.Value
        If valeur <> "" And cellule.Row > pt_sites.RowRange.Rows(1).Row Then
            If InStr(1, valeur, "Total") > 0 Then
                cellule.Offset(0, 2).Clear
                cellule.Offset(0, 2).Interior.ColorIndex = xlNone
            Else
                CheckValeur cellule, dictSite, hasFalse
            End If
        End If
    Next cellule
    
    ' QUALIFICATION
    For Each cellule In pt_qualifs.RowRange
        valeur = cellule.Value
        If valeur <> "" And cellule.Row > pt_qualifs.RowRange.Rows(1).Row Then
            If InStr(1, valeur, "Total") > 0 Then
                cellule.Offset(0, 2).Clear
                cellule.Offset(0, 2).Interior.ColorIndex = xlNone
            Else
                CheckValeur cellule, dictQualif, hasFalse
            End If
        End If
    Next cellule
    
    ' AGENCE
    For Each cellule In pt_agence.RowRange
        valeur = cellule.Value
        If valeur <> "" And cellule.Row > pt_agence.RowRange.Rows(1).Row Then
            If InStr(1, valeur, "Total") > 0 Then
                cellule.Offset(0, 2).Clear
                cellule.Offset(0, 2).Interior.ColorIndex = xlNone
            Else
                CheckValeur cellule, dictAgence, hasFalse
            End If
        End If
    Next cellule
    
    ' ==========================
    ' RESULTAT N2
    ' ==========================
    With wsPre.Range("C2:D2")
        .Merge
        
        If hasFalse Then
            .Value = ChrW(9888) & "ATTENTION" ' ?? Unicode
            .Interior.Color = RGB(255, 255, 0) ' jaune
            .Font.Color = RGB(255, 0, 0) ' rouge
        Else
            .Value = "CHECK !"
            .Interior.Color = RGB(239, 125, 0) 'RGB(0, 176, 80) ' vert
            .Font.Color = RGB(255, 255, 255) ' blanc
        End If
        
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    MsgBox "Vérification complète terminée", vbInformation

End Sub

' ==========================
' Fonction commune
' ==========================
Sub CheckValeur(cellule As Range, dict As Object, ByRef hasFalse As Boolean)

    If dict.exists(cellule.Value) Then
        cellule.Offset(0, 2).Value = "VRAI"
        cellule.Offset(0, 2).Interior.Color = RGB(239, 125, 0) 'RGB(217, 217, 217) 'RGB(0, 176, 80)
        cellule.Offset(0, 2).Font.Color = RGB(255, 255, 255) 'RGB(0, 0, 0) 'RGB(255, 255, 255)
    Else
        cellule.Offset(0, 2).Value = "FAUX"
        cellule.Offset(0, 2).Interior.Color = RGB(0, 0, 0) 'RGB(192, 0, 0)
        cellule.Offset(0, 2).Font.Color = RGB(255, 255, 255)
        
        hasFalse = True
    End If

End Sub



