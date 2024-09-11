Attribute VB_Name = "RechercheProduit"
Function VerifierProduit(Designation As String, NLot As String, PrixAchat As Currency, DateExp As Date, Marge As Currency, SHP As Currency, CleEmplacement As Long, Quantite As Long, ByRef CleLot As Long) As Boolean
    On Error GoTo GestionErreur
    Dim db As DAO.Database
    Dim trouve As Boolean
    Dim DesignationEquivalente As String
    Dim rstProduit As DAO.Recordset
    Dim rstEquivalence As DAO.Recordset

    ' Nettoyer la désignation pour éviter les apostrophes non échappées
    Designation = Replace(Designation, "'", "''")

    ' Ouvrir la base de données actuelle
    Set db = CurrentDb()

    ' Chercher dans la table Produit avec la désignation
    Set rstProduit = db.OpenRecordset("SELECT * FROM Produit WHERE Designation = '" & Designation & "'")

    If Not rstProduit.EOF Then
        ' Le produit est trouvé
        trouve = True

        ' Gérer le lot
        CleLot = GererLot(db, rstProduit!CleProduit, NLot, PrixAchat, DateExp, Marge, SHP, CleEmplacement, Quantite)
    Else
        ' Chercher dans la table EquivalenceDesignation
        Set rstEquivalence = db.OpenRecordset("SELECT DesignationEquivalente FROM EquivalenceDesignation WHERE DesignationOriginale = '" & Designation & "'")

        If Not rstEquivalence.EOF Then
            ' Obtenir la désignation équivalente
            DesignationEquivalente = rstEquivalence!DesignationEquivalente

            ' Chercher le produit équivalent
            Set rstProduit = db.OpenRecordset("SELECT * FROM Produit WHERE Designation = '" & DesignationEquivalente & "'")

            If Not rstProduit.EOF Then
                ' Gérer le lot avec la nouvelle désignation
                CleLot = GererLot(db, rstProduit!CleProduit, NLot, PrixAchat, DateExp, Marge, SHP, CleEmplacement, Quantite)
                trouve = True
            Else
                ' Aucun produit équivalent n'est trouvé, ajout à la variable globale
                ProductsToAdd.Add Designation
                trouve = False
            End If
        Else
            ' Si aucun équivalent n'est trouvé, demander à l'utilisateur une désignation équivalente
            DesignationEquivalente = ObtenirDesignationEquivalente(Designation)
            
            If Len(DesignationEquivalente) > 0 Then
                ' Ajouter l'équivalence dans la table DesignationEquivalence
                db.Execute "INSERT INTO EquivalenceDesignation (DesignationOriginale, DesignationEquivalente) " & _
                           "VALUES ('" & Designation & "', '" & DesignationEquivalente & "')", dbFailOnError

                ' Chercher le produit équivalent
                Set rstProduit = db.OpenRecordset("SELECT * FROM Produit WHERE Designation = '" & DesignationEquivalente & "'")

                If Not rstProduit.EOF Then
                    ' Gérer le lot avec la nouvelle désignation
                    CleLot = GererLot(db, rstProduit!CleProduit, NLot, PrixAchat, DateExp, Marge, SHP, CleEmplacement, Quantite)
                    trouve = True
                Else
                    ' Aucun produit équivalent n'est trouvé, ajout à la variable globale
                    ProductsToAdd.Add Designation
                    trouve = False
                End If
            Else
                ' Aucun équivalent n'est trouvé, ajout à la variable globale
                ProductsToAdd.Add Designation
                trouve = False
            End If
        End If
        rstEquivalence.Close
    End If

    ' Assurer la fermeture des objets
    rstProduit.Close
    Set db = Nothing
    VerifierProduit = trouve
    Exit Function

GestionErreur:
    ' Gérer les erreurs ici
    MsgBox "Erreur : " & Err.Description, vbCritical
    VerifierProduit = False
End Function
Function ObtenirDesignationEquivalente(DesignationOriginale As String) As String
    ' Ouvrir le formulaire de recherche en passant la désignation originale en argument
    DoCmd.OpenForm "recherche", acNormal, , , , acDialog, DesignationOriginale
    
    ' Attendre que le formulaire soit fermé
    DoEvents
    
    ' Retourner la désignation équivalente
    ObtenirDesignationEquivalente = DesignationGlobale
    
End Function


