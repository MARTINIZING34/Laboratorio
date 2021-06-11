VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminicio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Inicio"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":2872A
   ScaleHeight     =   9375
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":56F02
      OLEDBString     =   $"Form1.frx":56F8A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Doctores"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtcontraseña 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   690
      IMEMode         =   3  'DISABLE
      Left            =   15240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6960
      Width           =   3135
   End
   Begin VB.TextBox txtusuario 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   690
      Left            =   15240
      TabIndex        =   0
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   15840
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   11520
      Top             =   8400
      Width           =   2895
   End
End
Attribute VB_Name = "frminicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
    txtusuario.ForeColor = RGB(69, 110, 174)
    txtcontraseña.ForeColor = RGB(69, 110, 174)
    Image1.Picture = LoadPicture(App.Path & "/ingresar.jpg")
    Image2.Picture = LoadPicture(App.Path & "/salir.jpg")
    Me.Icon = LoadPicture(App.Path & "/logo5.ico")
    'Image3.Picture = LoadPicture(App.Path & "/user.png")
End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "/ingresar2.jpg")
    
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "/ingresar.jpg")
    'Verificación de usuario y contraseña

Adodc1.RecordSource = "select * from Doctores where Nombre = '" + txtusuario.Text + "' and Contraseña = '" + txtcontraseña.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    MsgBox "Revise sus datos y vuelva a intentarlo", vbInformation, "Laboratorios el Puente "
Else
    Usuario = txtusuario.Text
    frmbuscar.Show
    Unload Me
End If
End Sub


Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "/salir2.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "/salir.jpg")
If MsgBox("¿Desea salir?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
    Unload Me
End If
End Sub
