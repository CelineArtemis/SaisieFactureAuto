Attribute VB_Name = "insertEffet"
Function InsererEffet(CleTypeEffet As Long, reference As String, insertDate As Date, pied As String, _
                     montantHT As Currency, MonantTTC As Currency, CleTiers As Long, _
                     ristourne As Currency, CleUser As Long, totalSHP As Currency, _
                     totalPPA As Currency) As Long

    Dim db As DAO.Database
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim CleEffet As Long

    On Error GoTo ErrHandler
    
    ' Ouvrir la base de donn�es actuelle
    Set db = CurrentDb
    
    ' Construire la requ�te SQL (format de date yyyy-mm-dd pour Access)
    sql = "INSERT INTO Effet (CleTypeEffet, Reference, [Date], Pied, MontantHT, MonantTTC, CleTiers, Ristourne, CleUser, TotalSHP, TotalPPA) " & _
          "VALUES (" & CleTypeEffet & ", '" & Replace(reference, "'", "''") & "', #" & Format(insertDate, "yyyy-mm-dd") & "#, '" & Replace(pied, "'", "''") & "', " & montantHT & ", " & MonantTTC & ", " & CleTiers & ", " & ristourne & ", " & CleUser & ", " & totalSHP & ", " & totalPPA & ");"
    
    ' Ex�cuter la requ�te d'insertion
    db.Execute sql, dbFailOnError
    
    ' R�cup�rer la derni�re cl� primaire ins�r�e
    Set rs = db.OpenRecordset("SELECT @@IDENTITY AS LastID", dbOpenSnapshot)
    If Not rs.EOF Then
        CleEffet = rs!LastID
    Else
        MsgBox "Erreur lors de la r�cup�ration de la cl� primaire de l'effet.", vbExclamation
        CleEffet = -1
    End If
    rs.Close
    Set rs = Nothing

Cleanup:
    ' Lib�rer l'objet base de donn�es
    Set db = Nothing
    
    ' Retourner la cl� primaire
    InsererEffet = CleEffet
    Exit Function

ErrHandler:
    MsgBox "Erreur lors de l'insertion de l'effet : " & Err.Description, vbCritical
    CleEffet = -1
    Resume Cleanup
End Function


