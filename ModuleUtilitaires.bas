Attribute VB_Name = "ModuleUtilitaires"
' Module Utilitaires

Function ObtenirCleProduit(Designation As String) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim CleProduit As Long

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT CleProduit FROM Produit WHERE Designation = '" & Designation & "'", dbOpenSnapshot)

    If Not rs Is Nothing And Not rs.EOF Then
        CleProduit = rs!CleProduit
    Else
        CleProduit = 0
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ObtenirCleProduit = CleProduit
End Function

Function GererLot(db As DAO.Database, CleProduit As Long, NLot As String, PrixAchat As Currency, DateExp As Date, Marge As Currency, SHP As Currency, CleEmplacement As Long, Quantite As Long) As Long
    Dim rsLot As DAO.Recordset
    Dim CleLot As Long
    
    ' Chercher dans la table Lot avec la clé produit, NLot, PrixAchat et DateExp
    Set rsLot = db.OpenRecordset("SELECT CleLot, Quantite FROM Lot WHERE CleProduit = " & CleProduit & " AND NLot = '" & NLot & "' AND PrixAchat = " & Round(PrixAchat, 2) & " AND DateExp = #" & Format(DateExp, "yyyy-mm-dd") & "#", dbOpenDynaset)
    
    If Not rsLot Is Nothing And Not rsLot.EOF Then
        ' Le lot existe déjà
        CleLot = rsLot!CleLot
        ' Mettre à jour la quantité
        rsLot.Edit
        rsLot!Quantite = rsLot!Quantite + Quantite
        rsLot.Update
        rsLot.Close
    Else
        ' Lot non trouvé, appeler la fonction AjouterLot
        rsLot.Close
        AjouterLot CleProduit, NLot, PrixAchat, Marge, SHP, DateExp, CleEmplacement, Quantite, CleLot
    End If

    Set rsLot = Nothing
    GererLot = CleLot
End Function

Sub AjouterLot(CleProduit As Long, NLot As String, PrixAchat As Currency, Marge As Currency, SHP As Currency, DateExp As Date, CleEmplacement As Long, Quantite As Long, ByRef CleLot As Long)
    Dim db As DAO.Database
    Dim sql As String
    Dim rsNewLot As DAO.Recordset
    Dim CleStatutLot As Long
    Dim PrixVente As Currency
    Dim PPA As Currency
    
    CleStatutLot = 1
    
    ' Calculer PPA et PrixVente avec arrondi à deux chiffres après la virgule
    PPA = Round(PrixAchat + (PrixAchat * Marge / 100), 2)
    PrixVente = Round(PPA + SHP, 2)
    
    ' Ouvrir la base de données actuelle
    Set db = CurrentDb()
    
    ' Créer la requête d'insertion
    sql = "INSERT INTO Lot (cleProduit, NLot, PrixAchat, PrixVente, SHP, PPA, Marge, DateExp, CleEmplacement, CleStatutLot, Quantite) " & _
          "VALUES (" & CleProduit & ", '" & NLot & "', " & Round(PrixAchat, 2) & ", " & Round(PrixVente, 2) & ", " & Round(SHP, 2) & ", " & Round(PPA, 2) & ", " & Round(Marge, 2) & ", #" & Format(DateExp, "yyyy-mm-dd") & "#, " & CleEmplacement & ", " & CleStatutLot & ", " & Quantite & ")"
    
    ' Exécuter la requête d'insertion
    db.Execute sql, dbFailOnError
    
    ' Ouvrir un recordset pour obtenir la clé du dernier enregistrement ajouté
    Set rsNewLot = db.OpenRecordset("SELECT MAX(CleLot) AS DernierLot FROM Lot WHERE cleProduit = " & CleProduit & " AND NLot = '" & NLot & "' AND PrixAchat = " & Round(PrixAchat, 2) & " AND DateExp = #" & Format(DateExp, "yyyy-mm-dd") & "#", dbOpenSnapshot)
    
    ' Récupérer la clé du nouveau lot
    If Not rsNewLot.EOF Then
        CleLot = rsNewLot!DernierLot
    End If
    
    ' Nettoyer
    rsNewLot.Close
    Set rsNewLot = Nothing
    Set db = Nothing
End Sub


