Attribute VB_Name = "DetailsLancement"
Sub InsererDetailsEffetBatch(CleTypeEffet As Long, reference As String, insertDate As Date, pied As String, montantHT As Currency, MonantTTC As Currency, CleTiers As Long, ristourne As Currency, CleUser As Long, totalSHP As Currency, totalPPA As Currency, Details As Variant)
    ' D�claration des variables pour les d�tails de l'effet
    Dim i As Long
    Dim CleEffet As Long
    

    ' Cr�er l'effet et obtenir CleEffet
    CleEffet = InsererEffet(CleTypeEffet, reference, insertDate, pied, montantHT, MonantTTC, CleTiers, ristourne, CleUser, totalSHP, totalPPA)
    
    ' V�rifier si CleEffet est valide avant de continuer
    If CleEffet <= 0 Then
        MsgBox "Erreur lors de la cr�ation de l'effet. L'insertion des d�tails a �t� annul�e.", vbExclamation
        Exit Sub
    End If
    
    ' Boucle pour ins�rer chaque d�tail
    For i = LBound(Details) To UBound(Details)
        InsererDetailsEffet CleEffet, CStr(Details(i)(0)), CLng(Details(i)(1)), CCur(Details(i)(2)), CDate(Details(i)(3)), CCur(Details(i)(4)), CCur(Details(i)(5)), CStr(Details(i)(6))
    Next i
    AfficherProduitsManuels reference
    MsgBox "Facture ajout�e avec succ�s. On dit Merci Amine :D ", vbExclamation
    
End Sub
Sub AfficherProduitsManuels(reference As String)
    Dim item As Variant
    Dim message As String
    Dim fso As Object
    Dim logFile As Object
    Dim filePath As String
    
    ' V�rifier si la collection est initialis�e et non vide
    If IsObject(ProductsToAdd) Then
        If ProductsToAdd.Count > 0 Then
            ' Cr�er le message � afficher
            message = "Facture & reference & ajout�e avec succ�s, VOICI LA LISTE A AJOUTER MANUELLEMENT :" & vbCrLf & vbCrLf
            For Each item In ProductsToAdd
                message = message & item & vbCrLf
            Next item
            
            ' Afficher le message
            MsgBox message, vbInformation
            
            ' D�finir le chemin du fichier log sur le bureau avec le nom bas� sur la r�f�rence
            filePath = Environ("USERPROFILE") & "\Desktop\" & reference & ".txt"
            
            ' Cr�er le fichier texte et �crire dedans
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set logFile = fso.CreateTextFile(filePath, True)
            logFile.WriteLine "Facture " & reference & " ajout�e avec succ�s, VOICI LA LISTE A AJOUTER MANUELLEMENT :" & vbCrLf
            For Each item In ProductsToAdd
                logFile.WriteLine item
            Next item
            
            ' Fermer le fichier
            logFile.Close
        End If
    End If
End Sub
