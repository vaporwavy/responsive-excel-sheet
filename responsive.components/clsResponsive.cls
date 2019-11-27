VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResponsive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub responsive()

    Const strPattern As String = "^main$|^main\d*$"
    Dim RE As Object
    Dim nm As Name
    
    Set RE = CreateObject("VBScript.RegExp")
    
    With RE
        .Pattern = strPattern
        .IgnoreCase = True
        .Global = True
    End With
    
    For Each nm In ThisWorkbook.Names
        If RE.test(nm.Name) _
        And nm.RefersToRange.Parent.Name = ActiveSheet.Name Then
            Call setMain(nm)
        End If
    Next
    
    Set RE = Nothing
    
End Sub

Private Sub setMain(nm As Name)

    Dim widthMain   As Double
    Dim widthCols   As Double
    Dim colN        As Long
    Dim colR        As Long
    Dim colF        As Long
    
    widthMain = nm.RefersToRange.Width
    widthCols = nm.RefersToRange.Parent.Columns(1).Width
    
    colN = nm.RefersToRange.Column - 1
    colR = (ActiveWindow.ActivePane.VisibleRange.Width - widthMain) / 2 / widthCols
    colF = colR - colN
    
    If ActiveWindow.ActivePane.VisibleRange.Width > widthMain Then
        For i = 1 To colF
            nm.RefersToRange.Parent.Columns(1).Insert
        Next
        For i = -1 To colF Step -1
            nm.RefersToRange.Parent.Columns(1).Delete
        Next
    Else
        For i = 1 To colN
            nm.RefersToRange.Parent.Columns(1).Delete
        Next
    End If
    
End Sub

