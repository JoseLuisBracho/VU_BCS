VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formulario 
   Caption         =   "BCS - Assignment 2"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390
   OleObjectBlob   =   "Formulario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
For sel = 0 To ListBox1.ListCount - 1
ListBox1.Selected(sel) = True
Next sel
Else
For sel = 0 To ListBox1.ListCount - 1
ListBox1.Selected(sel) = False
Next sel
End If
End Sub

Private Sub CommandButton1_Click()
For X = 1 To ListBox1.ListCount
If ListBox1.Selected(X - 1) = True Then
Worksheets(X).Activate
Call stocks_yearly
End If
Next X
End Sub

Private Sub CommandButton2_Click()
Unload Formulario
End Sub

Private Sub UserForm_Initialize()
Call addSheets
End Sub

Sub addSheets()
Dim sheet As Worksheet
For Each sheet In Worksheets
Formulario.ListBox1.AddItem sheet.Name
Next sheet
End Sub


