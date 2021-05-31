Private Sub cmbPopulaCombo(cmb As ComboBox, selects As String, froms As String, Wheres As String, orderbys As String, ItemData As Boolean)
  Dim SpRecordSet As Recordset
  SQL = "Select " & selects & Chr(10)
  SQL = SQL & "From " & froms & Chr(10)
  If Wheres <> "" Then
    SQL = SQL & "WHERE " & Wheres & Chr(10)
  End If
  If orderbys <> "" Then
    SQL = SQL & "ORDER BY " & orderbys & Chr(10)
  End If
  
  Set SpRecordSet = CorpDb.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassTrough)
    
  cmb.Clear
  If Not SpRecordSet.EOF Then
    While Not SpRecordSet.EOF
      If ItemData Then
        cmb.AddItem (Trim(SpRecordSet.Fields(1)))
        cmb.ItemData(cmb.NewIndex) = SpRecordSet(0)
      Else
        cmb.AddItem (Trim(SpRecordSet.Fields(0)))
      End If
      SpRecordSet.MoveNext
    Wend
    If cmb.ListCount > 0 Then
      cmb.ListIndex = 0
    End If
    SpRecordSet.Close
    Set SpRecordSet = Nothing
  End If
End Sub
