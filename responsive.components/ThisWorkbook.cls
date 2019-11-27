VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If ThisWorkbook.VBProject.Saved Then
        Exit Sub
    End If
    
    Dim expPath As String
    Dim expName As String
    Dim buf     As String
    
    expPath = Left(ThisWorkbook.FullName, InStr(ThisWorkbook.FullName, ".xlsm")) & "components"
    
    If Dir(expPath, vbDirectory) = "" Then
        MkDir expPath
    Else
        buf = Dir(expPath & "\")
        Do While buf <> ""
            Kill expPath & "\" & buf
            buf = Dir()
        Loop
    End If
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        expName = comp.Name
        
        If comp.Type = 1 Then
            expName = expName & ".bas"
        ElseIf comp.Type = 3 Then
            expName = expName & ".frm"
        ElseIf comp.Type = 2 Or comp.Type = 100 Then
            expName = expName & ".cls"
        End If
        
        comp.Export expPath & "\" & expName
    Next
End Sub

Private Sub Workbook_Open()
    ThisWorkbook.Activate
    Dim clsR As clsResponsive
    Set clsR = New clsResponsive
    Call clsR.responsive
    Set clsR = Nothing
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ThisWorkbook.Activate
    Dim clsR As clsResponsive
    Set clsR = New clsResponsive
    Call clsR.responsive
    Set clsR = Nothing
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    ThisWorkbook.Activate
    Dim clsR As clsResponsive
    Set clsR = New clsResponsive
    Call clsR.responsive
    Set clsR = Nothing
End Sub
