Attribute VB_Name = "Module1"
Sub create_ExcelTable()
Attribute create_ExcelTable.VB_ProcData.VB_Invoke_Func = " \n14"

    shtName = ActiveSheet.Name
    ActiveCell.CurrentRegion.Select
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = shtName
    ActiveSheet.ListObjects(shtName).TableStyle = "TableStyleLight2"
    
    ActiveSheet.ListObjects(shtName).Name = shtName

End Sub
