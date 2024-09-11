Attribute VB_Name = "LireCSVFacture"
Dim referenceGlobale As String

Sub ImporterDonneesCSV()
    Dim fs As Object
    Dim ts As Object
    Dim ligne As String
    Dim valeurs() As String
    
    ' Définir le chemin du fichier CSV
    Dim cheminFichier As String
    cheminFichier = "C:\Users\BABIBORIS_WORK\Downloads\factures\facture.csv"
    
    ' Vérification du fichier
    If Dir(cheminFichier) = "" Then
        MsgBox "Fichier facture.csv introuvable!", vbExclamation
        Exit Sub
    End If
    
    InitializeGlobalVariables
    
    ' Créer des variables pour stocker les données
    Dim reference As String
    Dim dateInsertion As Date
    Dim fournisseur As String
    Dim montantHT As Currency
    Dim montantTTC As Currency
    Dim ristourne As Currency
    Dim totalSHP As Currency
    Dim totalPPA As Currency
    Dim pied As String
    Dim CleTypeEffet As Long
    Dim CleUser As Long
    Dim CleTiers As Long
    Dim produits As Variant
    
    ' Créer les objets FileSystemObject et TextStream
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.OpenTextFile(cheminFichier, 1) ' 1 = ForReading
    
    ' Lire la première ligne (en-têtes) et ignorer
    ts.ReadLine
    
    ' Lire la ligne suivante (les données)
    ligne = ts.ReadLine
    
    ' Diviser la ligne en valeurs en utilisant la virgule comme séparateur
    valeurs = Split(ligne, ",")
    
    ' Assigner les valeurs aux variables en respectant les types
    If UBound(valeurs) >= 8 Then
        reference = valeurs(0)
        dateInsertion = CDate(valeurs(1)) ' Convertir en Date
        fournisseur = Trim(UCase(valeurs(2))) ' Nettoyer et convertir en majuscules
        montantHT = CCur(valeurs(3)) ' Convertir en Currency
        montantTTC = CCur(valeurs(4)) ' Convertir en Currency
        ristourne = CCur(valeurs(5)) ' Convertir en Currency
        totalSHP = CCur(valeurs(6)) ' Convertir en Currency
        totalPPA = CCur(valeurs(7)) ' Convertir en Currency
        pied = valeurs(8)
        
        ' Valeurs fixées pour cet exemple
        CleTypeEffet = 9 ' Toujours 9 (Bon de réception)
        CleUser = 7 ' Toujours 7 (Compte utilisateur IA)
        
        ' Définir CleTiers en fonction du fournisseur
        Select Case fournisseur
            Case "BIOPURE"
                CleTiers = 306
            Case "SOMEPHARM"
                CleTiers = 299
            Case "PHARMA INVEST"
                CleTiers = 109
            Case "BCD PHAMA"
                CleTiers = 317
            Case "AZ VITA PHARM"
                CleTiers = 347
            Case Else
                ' Demander au utilisateur le nom du fournisseur s'il n'est pas dans la liste
                Dim nomFournisseur As String
                nomFournisseur = Trim(UCase(InputBox("QUEL EST LE NOM DU FOURNISSEUR ?", "Fournisseur inconnu")))
                
                ' Définir CleTiers en fonction du nom du fournisseur saisi par l'utilisateur
                Select Case nomFournisseur
                    Case "BIOPURE"
                        CleTiers = 306
                    Case "SOMEPHARM"
                        CleTiers = 299
                    Case "PHARMA INVEST"
                        CleTiers = 109
                    Case "BCD PHAMA"
                        CleTiers = 317
                    Case "AZ VITA PHARM"
                        CleTiers = 347
                    Case Else
                        MsgBox "Le fournisseur renseigné n'est pas reconnu. L'importation est annulée.", vbExclamation
                        ts.Close
                        Exit Sub
                End Select
        End Select
        
    Else
        MsgBox "Le nombre de colonnes dans le fichier CSV est incorrect."
        ts.Close
        Exit Sub
    End If
    
    referenceGlobale = reference
    ' Appeler ImporterProduitsCSV pour obtenir la liste des produits
    produits = ImporterProduitsCSV()
    
    If IsArray(produits) Then
        ' Appeler InsererDetailsEffetBatch en passant les détails de la facture et les produits
        Call InsererDetailsEffetBatch(CleTypeEffet, reference, dateInsertion, pied, montantHT, montantTTC, CleTiers, ristourne, CleUser, totalSHP, totalPPA, produits)
        ' MsgBox "Insertion des détails terminée!", vbInformation
    Else
        ' MsgBox "L'importation des produits a échoué.", vbExclamation
    End If
    
    ' Fermer le fichier
    ts.Close
    
    ' Renommer le fichier facture.csv en utilisant la valeur de reference
    Dim nouveauCheminFichier As String
    nouveauCheminFichier = "C:\Users\BABIBORIS_WORK\Downloads\factures\" & reference & ".csv"
    fs.MoveFile cheminFichier, nouveauCheminFichier
    

    ' Libérer les objets
    Set ts = Nothing
    Set fs = Nothing
End Sub

Function ImporterProduitsCSV() As Variant
    Dim fs As Object
    Dim ts As Object
    Dim ligne As String
    Dim valeurs() As String
    Dim produits As Collection
    Dim produit As Variant
    Dim produitArray() As Variant
    Dim i As Integer
    
    ' Définir le chemin du fichier CSV
    Dim cheminFichier As String
    cheminFichier = "C:\Users\BABIBORIS_WORK\Downloads\factures\produits.csv"
    
    ' Vérification du fichier
    If Dir(cheminFichier) = "" Then
        MsgBox "Fichier produits.csv introuvable!", vbExclamation
        Exit Function
    End If

    ' Créer une collection pour stocker les informations des produits
    Set produits = New Collection
    
    ' Créer les objets FileSystemObject et TextStream
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.OpenTextFile(cheminFichier, 1) ' 1 = ForReading
    
    ' Lire la première ligne (en-têtes) et ignorer
    ts.ReadLine
    
    ' Lire les lignes suivantes (les données)
    Do While Not ts.AtEndOfStream
        ligne = ts.ReadLine
        
        ' Diviser la ligne en valeurs en utilisant la virgule comme séparateur
        valeurs = Split(ligne, ",")
        
        ' Assigner les valeurs aux variables en respectant les types
        If UBound(valeurs) >= 6 Then
            produit = Array(valeurs(0), CLng(valeurs(1)), CCur(valeurs(2)), CDate(valeurs(3)), CCur(valeurs(4)), CCur(valeurs(5)), valeurs(6))
            produits.Add produit
        End If
    Loop
    
    ' Convertir la collection en un tableau
    ReDim produitArray(1 To produits.Count)
    For i = 1 To produits.Count
        produitArray(i) = produits(i)
    Next i
    
    ' Retourner le tableau de produits
    ImporterProduitsCSV = produitArray
    ' Fermer le fichier
    ts.Close
        
    ' Renommer le fichier produits.csv en reference_produits.csv
    Dim nouveauCheminFichier As String
    nouveauCheminFichier = "C:\Users\BABIBORIS_WORK\Downloads\factures\" & referenceGlobale & "_produits.csv"
    fs.MoveFile cheminFichier, nouveauCheminFichier
    

    ' Libérer les objets
    Set ts = Nothing
    Set fs = Nothing
End Function

