Attribute VB_Name = "RechercheProduit"
Function VerifierProduit(Designation As String, NLot As String, PrixAchat As Currency, DateExp As Date, Marge As Currency, SHP As Currency, CleEmplacement As Long, Quantite As Long, ByRef CleLot As Long) As Boolean
    On Error GoTo GestionErreur
    Dim db As DAO.Database
    Dim trouve As Boolean
    Dim DesignationEquivalente As String
    Dim rstProduit As DAO.Recordset
    Dim rstEquivalence As DAO.Recordset

    ' Nettoyer la d�signation pour �viter les apostrophes non �chapp�es
    Designation = Replace(Designation, "'", "''")

    ' Ouvrir la base de donn�es actuelle
    Set db = CurrentDb()

    ' Chercher dans la table Produit avec la d�signation
    Set rstProduit = db.OpenRecordset("SELECT * FROM Produit WHERE Designation = '" & Designation & "'")

    If Not rstProduit.EOF Then
        ' Le produit est trouv�
        trouve = True

        ' G�rer le lot
        CleLot = GererLot(db, rstProduit!CleProduit, NLot, PrixAchat, DateExp, Marge, SHP, CleEmplacement, Quantite)
    Else
        ' Chercher dans la table EquivalenceDesignation
        Set rstEquivalence = db.OpenRecordset("SELECT DesignationEquivalente FROM EquivalenceDesignation WHERE DesignationOriginale = '" & Designation & "'")

        If Not rstEquivalence.EOF Then
            ' Obtenir la d�signation �quivalente
            DesignationEquivalente = rstEquivalence!DesignationEquivalente

            ' Chercher le produit �quivalent
            Set rstProduit = db.OpenRecordset("SELECT * FROM Produit WHERE Designation = '" & DesignationEquivalente & "'")

            If Not rstProduit.EOF Then
                ' G�rer le lot avec la nouvelle d�signation
                CleLot = GererLot(db, rstProduit!CleProduit, NLot, PrixAchat, DateExp, Marge, SHP, CleEmplacement, Quantite)
                trouve = True
            Else
                ' Aucun produit �quivalent n'est trouv�, ajout � la variable globale
                ProductsToAdd.Add Designation
                trouve = False
            End If
        Else
            ' Si aucun �quivalent n'est trouv�, demander � l'utilisateur une d�signation �quivalente
            DesignationEquivalente = ObtenirDesignationEquivalente(Designation)
            
            If Len(DesignationEquivalente) > 0 Then
                ' Ajouter l'�quivalence dans la table DesignationEquivalence
                db.Execute "INSERT INTO EquivalenceDesignation (DesignationOriginale, DesignationEquivalente) " & _
                           "VALUES ('" & Designation & "', '" & DesignationEquivalente & "')", dbFailOnError

                ' Chercher le produit �quivalent
                Set rstProduit = db.OpenRecordset("SELECT * FROM Produit WHERE Designation = '" & DesignationEquivalente & "'")

                If Not rstProduit.EOF Then
                    ' G�rer le lot avec la nouvelle d�signation
                    CleLot = GererLot(db, rstProduit!CleProduit, NLot, PrixAchat, DateExp, Marge, SHP, CleEmplacement, Quantite)
                    trouve = True
                Else
                    ' Aucun produit �quivalent n'est trouv�, ajout � la variable globale
                    ProductsToAdd.Add Designation
                    trouve = False
                End If
            Else
                ' Aucun �quivalent n'est trouv�, ajout � la variable globale
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
    ' G�rer les erreurs ici
    MsgBox "Erreur : " & Err.Description, vbCritical
    VerifierProduit = False
End Function
Function ObtenirDesignationEquivalente(DesignationOriginale As String) As String
    ' Ouvrir le formulaire de recherche en passant la d�signation originale en argument
    DoCmd.OpenForm "recherche", acNormal, , , , acDialog, DesignationOriginale
    
    ' Attendre que le formulaire soit ferm�
    DoEvents
    
    ' Retourner la d�signation �quivalente
    ObtenirDesignationEquivalente = DesignationGlobale
    
End Function


