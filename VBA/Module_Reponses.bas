Attribute VB_Name = "Module_Reponses"
Option Explicit

'===============================================================================================
'  GAILLON - Suivi des VISAS
'  Ajout d'une réponse (visa + date + commentaire) d'une entité sur une fiche
'
'  À coller dans un NOUVEAU MODULE standard (Insertion > Module).
'  Macro à affecter au bouton : AjouterReponseVisa
'===============================================================================================

'---------------------------- PARAMÈTRES (à adapter si le tableau évolue) ----------------------
Public Const TITRE As String = "Réponse au visa"
Public Const FEUILLE_SUIVI As String = "Tableau de suivi"
Public Const FEUILLE_DONNEES As String = "Données"

Public Const LIG_ENTITES As Long = 4          ' ligne des noms d'entités (AB4, AE4, ...)
Public Const LIG_SOUS_ENTETES As Long = 5     ' ligne des sous-entêtes "Visa / Date / Commentaire"
Public Const PREM_LIGNE As Long = 6           ' première ligne de données
Public Const COL_FICHE As Long = 4            ' colonne D : Fiche n°
Public Const COL_INDICE As Long = 5           ' colonne E : ind.
Public Const COL_SUJET As Long = 9            ' colonne I : Sujet
Public Const COL_PREM_BLOC As Long = 28       ' colonne AB : 1er bloc destinataire
Public Const COL_DERN_BLOC As Long = 70       ' recherche des blocs jusqu'à cette colonne

Public Const LIG_VISA_DEB As Long = 47        ' Données : début liste des visas
Public Const LIG_VISA_FIN As Long = 60        ' Données : fin de recherche
Public Const COL_VISA_SIGNIF As Long = 3      ' Données colonne C : signification
Public Const COL_VISA_ABREV As Long = 5       ' Données colonne E : abreviation

'===============================================================================================
'  MACRO À AFFECTER AU BOUTON
'===============================================================================================
Public Sub AjouterReponseVisa()

    Dim frm As frmReponseVisa

    On Error GoTo Gest_Erreur

    If Not FeuilleExiste(FEUILLE_SUIVI) Then
        MsgBox "L'onglet """ & FEUILLE_SUIVI & """ est introuvable.", vbCritical, TITRE
        Exit Sub
    End If

    If DerniereLigneFiches() < PREM_LIGNE Then
        MsgBox "Aucune fiche n'est enregistrée dans l'onglet """ & FEUILLE_SUIVI & """.", _
               vbExclamation, TITRE
        Exit Sub
    End If

    Set frm = New frmReponseVisa
    frm.Show
    Unload frm
    Set frm = Nothing
    Exit Sub

Gest_Erreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, TITRE

End Sub

'===============================================================================================
'  FONCTIONS UTILISÉES PAR LE FORMULAIRE
'===============================================================================================

Public Function FeuilleSuivi() As Worksheet
    Set FeuilleSuivi = ThisWorkbook.Worksheets(FEUILLE_SUIVI)
End Function

Public Function FeuilleExiste(ByVal nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not ws Is Nothing
End Function

' Dernière ligne contenant un numéro de fiche
Public Function DerniereLigneFiches() As Long
    Dim ws As Worksheet
    Dim lig As Long
    Set ws = FeuilleSuivi()
    lig = ws.Cells(ws.Rows.Count, COL_FICHE).End(xlUp).Row
    If lig < PREM_LIGNE Then lig = PREM_LIGNE - 1
    DerniereLigneFiches = lig
End Function

' Renvoie le contenu d'une cellule en texte (jj/mm/aaaa pour les dates), sans erreur
Public Function TexteCellule(ByVal cel As Range) As String
    Dim v As Variant
    v = cel.Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsDate(v) Then
        TexteCellule = Format$(v, "dd/mm/yyyy")
    Else
        TexteCellule = Trim$(CStr(v))
    End If
End Function

' Transforme "MON LOGEMENT 27" & saut de ligne & "M. BANCE" en "MON LOGEMENT 27 - M. BANCE"
Public Function NettoyerLibelle(ByVal valeur As Variant) As String
    Dim txt As String
    Dim morceaux() As String
    Dim i As Long
    Dim res As String

    If IsError(valeur) Then Exit Function
    If IsEmpty(valeur) Then Exit Function

    txt = Replace(CStr(valeur), vbCr, vbLf)
    morceaux = Split(txt, vbLf)
    For i = LBound(morceaux) To UBound(morceaux)
        If Trim$(morceaux(i)) <> "" Then
            If res <> "" Then res = res & " - "
            res = res & Trim$(morceaux(i))
        End If
    Next i
    NettoyerLibelle = res
End Function

' Liste des blocs "destinataires" : colonne de départ + nom de l'entité
Public Sub ListerEntites(ByRef cols() As Long, ByRef noms() As String, ByRef nb As Long)

    Dim ws As Worksheet
    Dim col As Long
    Dim nom As String

    Set ws = FeuilleSuivi()
    nb = 0
    ReDim cols(1 To 60)
    ReDim noms(1 To 60)

    For col = COL_PREM_BLOC To COL_DERN_BLOC Step 3
        If LCase$(TexteCellule(ws.Cells(LIG_SOUS_ENTETES, col))) = "visa" Then
            nom = NettoyerLibelle(ws.Cells(LIG_ENTITES, col).Value)
            If nom <> "" Then
                nb = nb + 1
                cols(nb) = col
                noms(nb) = nom
            End If
        End If
    Next col

End Sub

' Liste des visas possibles (onglet Données) : code + libellé affiché
Public Sub ListerVisas(ByRef codes() As String, ByRef libelles() As String, ByRef nb As Long)

    Dim ws As Worksheet
    Dim lig As Long
    Dim code As String
    Dim signif As String

    nb = 0
    ReDim codes(1 To 30)
    ReDim libelles(1 To 30)

    If FeuilleExiste(FEUILLE_DONNEES) Then
        Set ws = ThisWorkbook.Worksheets(FEUILLE_DONNEES)
        For lig = LIG_VISA_DEB To LIG_VISA_FIN
            code = TexteCellule(ws.Cells(lig, COL_VISA_ABREV))
            signif = TexteCellule(ws.Cells(lig, COL_VISA_SIGNIF))
            If code <> "" Then
                nb = nb + 1
                codes(nb) = code
                If signif <> "" Then
                    libelles(nb) = code & " - " & signif
                Else
                    libelles(nb) = code
                End If
            End If
        Next lig
    End If

    ' Sécurité : si la liste n'a pas été trouvée dans l'onglet Données
    If nb = 0 Then
        nb = 6
        codes(1) = "F":   libelles(1) = "F - Favorable"
        codes(2) = "FAO": libelles(2) = "FAO - Favorable avec observations"
        codes(3) = "S":   libelles(3) = "S - Suspendu"
        codes(4) = "R":   libelles(4) = "R - Refusé"
        codes(5) = "ATT": libelles(5) = "ATT - En attente"
        codes(6) = "EXP": libelles(6) = "EXP - Expiré"
    End If

End Sub

' Écriture de la réponse dans les 3 cellules du bloc de l'entité
Public Function EnregistrerReponse(ByVal ligne As Long, _
                                   ByVal colBloc As Long, _
                                   ByVal codeVisa As String, _
                                   ByVal dateVisa As Variant, _
                                   ByVal texte As String) As Boolean

    Dim ws As Worksheet

    On Error GoTo Gest_Erreur

    Set ws = FeuilleSuivi()

    If Len(texte) > 32000 Then
        texte = Left$(texte, 32000)
        MsgBox "La réponse dépasse la taille maximale d'une cellule Excel :" & vbLf & _
               "le texte a été tronqué à 32 000 caractères.", vbExclamation, TITRE
    End If

    Application.EnableEvents = False

    ' Visa
    ws.Cells(ligne, colBloc).Value = codeVisa

    ' Date
    If IsDate(dateVisa) Then
        ws.Cells(ligne, colBloc + 1).Value = CDate(dateVisa)
    Else
        ws.Cells(ligne, colBloc + 1).ClearContents
    End If

    ' Commentaire / réponse
    If Len(texte) > 0 Then
        ws.Cells(ligne, colBloc + 2).Value = texte
    Else
        ws.Cells(ligne, colBloc + 2).ClearContents
    End If

    Application.EnableEvents = True
    EnregistrerReponse = True
    Exit Function

Gest_Erreur:
    Application.EnableEvents = True
    MsgBox "Impossible d'écrire la réponse." & vbLf & _
           "Erreur " & Err.Number & " : " & Err.Description & vbLf & vbLf & _
           "Vérifiez que la feuille n'est pas protégée.", vbCritical, TITRE
    EnregistrerReponse = False

End Function
