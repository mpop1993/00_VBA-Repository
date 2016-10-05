VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} genericFormAdmin 
   Caption         =   "Template"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3855
   OleObjectBlob   =   "genericFormAdmin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "genericFormAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdmin_Click()
    
On Error GoTo Handler
 
    If btnAdmin.Value Then
        adminMain.Height = 335
    Else
        adminMain.Height = 214
    End If
    
    Exit Sub
    
Handler:

    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Error"
    
End Sub

Private Sub btnExit_Click()
    ' Close application on Exit button push
    Application.ActiveWorkbook.Close
End Sub

Private Sub btnLogin_Click()

On Error GoTo Handler
 
    If TextBox1.Value = USER And TextBox2.Value = PASS Then
        Application.Visible = True
        Unload Me
    Else
        MsgBox "Please enter a valid user and password"
    End If
    
    Exit Sub
    
Handler:

    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Error"
    
End Sub

Private Sub UserForm_Initialize()
    adminMain.Height = 214
    Application.Visible = False
    
End Sub
