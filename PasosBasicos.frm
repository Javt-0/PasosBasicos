VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Control de datos"
   ClientHeight    =   6456
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10380
   OleObjectBlob   =   "PasosBasicos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    MsgBox "Vamos a introducir datos en la celda A8"
    Hoja1.Cells(8, 1) = "Buenos dias"
    CommandButton2.Visible = True
    CommandButton1.Enabled = False
End Sub

Private Sub CommandButton2_Click()
    MsgBox "Coger los datos de la celda A8"
    TextBox1.Text = Hoja1.Cells(8, 1)
    CommandButton2.Visible = False
    CommandButton1.Enabled = True
End Sub
