Private Sub lstReq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Verifica se SortKey é a mesma que a atual
    If lstReq.SortKey <> ColumnHeader.Index - 1 Then
        'Quando clicar em uma coluna define sortkey para indice -1
        lstReq.SortKey = ColumnHeader.Index - 1
        lstReq.SortOrder = lvwAscending
    Else
    'Se a coluna ja esta selecionada entao altera a
    'propr. SetOrder para ser o oposto da coluna em uso
        lstReq.SortOrder = IIf(lstReq.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If

    'Define a propriedade Sorted para utilizar a ordem atual
    lstReq.Sorted = True
End Sub
