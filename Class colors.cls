VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.EnableEvents = False
    
    Cells.ClearFormats
    
    Target.Interior.Color = vbBlack
    
    Dim I As Long
    
    For I = 1 To Target.Row - 1
        Cells(I, Target.Column).Interior.Color = vbRed
    Next I
    
    For I = 1 To Target.Column - 1
        Cells(Target.Row, I).Interior.Color = vbYellow
    Next I
    
    Application.EnableEvents = True
End Sub
