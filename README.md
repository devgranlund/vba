# vba
~~~~
Private Sub Add_Click()

    Dim tbl As table
    Dim rowCount As Integer
    Dim nameControl As ContentControl
    Dim addressControl As ContentControl
    Dim functionsControl As ContentControl
    Dim personnelControl As ContentControl
    Dim newControl As ContentControl
    Dim lProt As Long: Const Pwd As String = "testi"

    Set tbl = ActiveDocument.Tables(3)
    rowCount = tbl.Rows.Count

    With ActiveDocument
        If .ProtectionType <> wdNoProtection Then
            lProt = .ProtectionType
            .Unprotect Password:=Pwd
        End If
        
        tbl.Rows(rowCount).Select
        Selection.InsertRowsBelow (1)
    
        FillCell tbl:=tbl
        
        If lProt <> wdNoProtection Then .Protect Type:=lProt, NoReset:=True, Password:=Pwd
    End With
End Sub

Sub FillCell(ByRef tbl As table)
    rowCount = tbl.Rows.Count
    tbl.Cell(Row:=rowCount, Column:=1).Range.Select
    Set nameControl = ActiveDocument.ContentControls.Add(wdContentControlRichText, Selection.Range)
    
    tbl.Cell(Row:=rowCount, Column:=2).Range.Select
    Set addressControl = ActiveDocument.ContentControls.Add(wdContentControlRichText, Selection.Range)
    
    tbl.Cell(Row:=rowCount, Column:=3).Range.Select
    Set functionsControl = ActiveDocument.ContentControls.Add(wdContentControlRichText, Selection.Range)
    
    tbl.Cell(Row:=rowCount, Column:=4).Range.Select
    Set personnelControl = ActiveDocument.ContentControls.Add(wdContentControlRichText, Selection.Range)
    
    tbl.Cell(Row:=rowCount, Column:=5).Range.Select
    Set newControl = ActiveDocument.ContentControls.Add(wdContentControlCheckBox, Selection.Range)
End Sub
~~~~
