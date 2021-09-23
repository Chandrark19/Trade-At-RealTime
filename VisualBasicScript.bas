Attribute VB_Name = "Module1"
Sub SaveWb()
    ThisWorkbook.RefreshAll
    ThisWorkbook.Save
    Application.OnTime Now + TimeValue("00:00:05"), "SaveWb"
End Sub
