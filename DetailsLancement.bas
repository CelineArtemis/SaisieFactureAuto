Attribute VB_Name = "DetailsLancement"
Sub InsererDetailsEffetBatch(CleTypeEffet As Long, reference As String, insertDate As Date, pied As String, montantHT As Currency, MonantTTC As Currency, CleTiers As Long, ristourne As Currency, CleUser As Long, totalSHP As Currency, totalPPA As Currency, Details As Variant)
    ' Déclaration des variables pour les détails de l'effet
    Dim i As Long
    Dim CleEffet As Long
    

    ' Créer l'effet et obtenir CleEffet
    CleEffet = InsererEffet(CleTypeEffet, reference, insertDate, pied, montantHT, MonantTTC, CleTiers, ristourne, CleUser, totalSHP, totalPPA)
    
    ' Vérifier si CleEffet est valide avant de continuer
    If CleEffet <= 0 Then
        MsgBox "Erreur lors de la création de l'effet. L'insertion des détails a été annulée.", vbExclamation
        Exit Sub
    End If
    
    ' Boucle pour insérer chaque détail
    For i = LBound(Details) To UBound(Details)
        InsererDetailsEffet CleEffet, CStr(Details(i)(0)), CLng(Details(i)(1)), CCur(Details(i)(2)), CDate(Details(i)(3)), CCur(Details(i)(4)), CCur(Details(i)(5)), CStr(Details(i)(6))
    Next i
    AfficherProduitsManuels reference
    MsgBox "Facture ajoutée avec succès. On dit Merci Amine :D ", vbExclamation
    
End Sub
Sub AfficherProduitsManuels(reference As String)
    Dim item As Variant
    Dim message As String
    Dim fso As Object
    Dim logFile As Object
    Dim filePath As String
    
    ' Vérifier si la collection est initialisée et non vide
    If IsObject(ProductsToAdd) Then
        If ProductsToAdd.Count > 0 Then
            ' Créer le message à afficher
            message = "Facture & reference & ajoutée avec succès, VOICI LA LISTE A AJOUTER MANUELLEMENT :" & vbCrLf & vbCrLf
            For Each item In ProductsToAdd
                message = message & item & vbCrLf
            Next item
            
            ' Afficher le message
            MsgBox message, vbInformation
            
            ' Définir le chemin du fichier log sur le bureau avec le nom basé sur la référence
            filePath = Environ("USERPROFILE") & "\Desktop\" & reference & ".txt"
            
            ' Créer le fichier texte et écrire dedans
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set logFile = fso.CreateTextFile(filePath, True)
            logFile.WriteLine "Facture " & reference & " ajoutée avec succès, VOICI LA LISTE A AJOUTER MANUELLEMENT :" & vbCrLf
            For Each item In ProductsToAdd
                logFile.WriteLine item
            Next item
            
            ' Fermer le fichier
            logFile.Close
        End If
    End If
End Sub
