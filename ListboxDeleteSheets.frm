VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} flmListBoxExpl 
   Caption         =   "UserForm1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "flmListBoxExpl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "flmListBoxExpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Application.DisplayAlerts = False
    
    Dim i As Long
    
    For i = lbxWorksheets.ListCount - 1 To 0 Step -1   'total number of options, listed sheet in the box - from last to first
        If lbxWorksheets.Selected(i) Then ' return a boolean value true or false
            Worksheets(i + 1).Delete  ' takes the sheet from the index position - cannot be worksheet 0
            lbxWorksheets.RemoveItem i 'remove from the listbox
        End If
    Next i
    

    Application.DisplayAlerts = True
End Sub

Private Sub UserForm_Initialize()
    lbxWorksheets.RowSource = ""
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        lbxWorksheets.AddItem ws.Name 'add item to the listbox
                
    Next ws
    
    
End Sub
