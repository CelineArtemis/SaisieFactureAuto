Attribute VB_Name = "insertDetailEffet"
Sub InsererDetailsEffet(CleEffet As Long, Designation As String, Quantite As Long, PrixUnitaireHT As Currency, _
                         DateExp As Date, Marge As Currency, SHP As Currency, NLot As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim CleLot As Long
    Dim ristourne As Currency
    Dim TVAPourCent As Long
    Dim NbColie As Long
    Dim Colissage As String
    Dim TailleCollie As Long
    Dim produitValide As Boolean
    Dim CleEmplacement As Long

    
    On Error GoTo ErrHandler
    
    ' Initialiser les valeurs par d�faut
    ristourne = 0
    TVAPourCent = 0
    NbColie = Quantite
    Colissage = "Unit�"
    TailleCollie = 1
    CleEmplacement = 60
    
    ' Ouvrir la base de donn�es actuelle
    Set db = CurrentDb()
    
    ' V�rifier la d�signation du produit avant l'insertion
    produitValide = VerifierProduit(Designation, NLot, PrixUnitaireHT, DateExp, Marge, SHP, CleEmplacement, Quantite, CleLot)
    
    If produitValide Then
        ' Construire la requ�te SQL
        sql = "INSERT INTO DetailEffet (Designation, Quantite, PrixUnitaireHT, CleLot, CleEffet, Ristourne, TVAPourCent, NbColie, Colissage, TailleCollie) " & _
              "VALUES ('" & Replace(Designation, "'", "''") & "', " & Quantite & ", " & Round(PrixUnitaireHT, 2) & ", " & CleLot & ", " & CleEffet & ", " & Round(ristourne, 2) & ", " & TVAPourCent & ", " & NbColie & ", '" & Replace(Colissage, "'", "''") & "', " & TailleCollie & ");"
        
        ' Ex�cuter la requ�te d'insertion
        db.Execute sql, dbFailOnError
    End If

Cleanup:
    ' Lib�rer l'objet base de donn�es
    Set db = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'insertion du d�tail de l'effet : " & Err.Description, vbCritical
    Resume Cleanup
End Sub


