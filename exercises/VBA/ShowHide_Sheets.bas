Attribute VB_Name = "ShowHide_Sheets"
Option Explicit

Sub show_HiddenSheets()

Dim F As Object, r, sheetCount As Integer, sheetNames() As String

Set F = ActiveWorkbook

F.Sheets.Add
sheetCount = F.Sheets.Count

ReDim sheetNames(1 To sheetCount)
For r = 1 To sheetCount
    sheetNames(r) = F.Sheets(r).Name
    Cells(r, 1) = sheetNames(r)
    If F.Sheets(r).Visible = False Then
        Cells(r, 2) = "Hidden"
        F.Sheets(r).Visible = True
    End If

Next r

ActiveWindow.DisplayWorkbookTabs = True

End Sub

Sub hide_HiddenSheets()

Dim F As Object, r, sheetCount As Integer, sheetName As String, Ans As VbMsgBoxResult

Set F = ActiveWorkbook
sheetCount = F.Sheets.Count

For r = 1 To sheetCount
sheetName = ActiveSheet.Cells(r, 1).Value

If ActiveSheet.Cells(r, 1) = ActiveSheet.Name Then
        ActiveSheet.Cells(r, 2) = "This sheet [" & ActiveSheet.Name & "] is not part of the documenet and should be deleted."
Else
    If ActiveSheet.Cells(r, 2) = "Hidden" Then
        F.Sheets(sheetName).Visible = False
    End If
End If

Next r

Ans = MsgBox("Delete this sheet?", vbYesNo, "Nothing else to hide")
If Ans = vbYes Then ActiveSheet.Delete

End Sub

