VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub RandomText()
    Dim range As range
    Set range = Sheets(1).range("E8")
    
    Dim RandomNumber As Integer
    RandomNumber = Int((56 - 0 + 1) * Rnd + 0)
     
    range.Font.ColorIndex = RandomNumber
    
    Dim randsize As Integer
    randsize = RandomNumber
    If randsize < 6 Then
        randsize = 6
    ElseIf randsize > 20 Then
        randsize = 20
    
    End If
    
range.Font.Size = randsize
    
range.Orientation = RandomNumber
    
    If RandomNumber < 28 Then
        range.Font.Bold = True
        range.Font.Italic = False
        range.Font.Underline = xlUnderlineStyleSingle
        range.Font.Strikethrough = True
    
    ElseIf RandomNumber >= 28 Then
        range.Font.Bold = False
        range.Font.Italic = True
        range.Font.Underline = False
        range.Font.Strikethrough = False
    End If
    
        
End Sub
