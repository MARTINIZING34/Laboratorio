VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmbuscar 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Selección de reactivos y marcadores tumorales"
   ClientHeight    =   8835
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   16500
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Cambiar de usuarios"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   19
      Top             =   8160
      Width           =   7095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14400
      TabIndex        =   18
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdusuarios 
      Caption         =   "Administrar usuarios"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   15
      Top             =   6960
      Width           =   7095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16800
      MaskColor       =   &H00000000&
      TabIndex        =   14
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdmodificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   13
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdregistrar 
      BackColor       =   &H00FFFF80&
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14400
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtusuario 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   9360
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   960
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
      Connect         =   $"Form2.frx":B2C0
      OLEDBString     =   $"Form2.frx":B348
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Administrador"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   16320
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Form2.frx":B3D0
      OLEDBString     =   $"Form2.frx":B458
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Reactivos"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":B4E0
      Height          =   4335
      Left            =   360
      TabIndex        =   10
      Top             =   4440
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      _Version        =   393216
      ForeColor       =   -2147483646
      HeadLines       =   1
      RowHeight       =   35
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   16800
      TabIndex        =   9
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtbuscartexto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.TextBox txtmarca 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtfecha 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtcantidad 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Utilizar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblbuscar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   8400
      TabIndex        =   17
      Top             =   3480
      Width           =   120
   End
   Begin VB.Label lblusuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   14160
      TabIndex        =   16
      Top             =   3480
      Width           =   1410
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione un reactivo de la lista:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   12240
      TabIndex        =   1
      Top             =   3480
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hormonas y marcadores tumorales"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   16695
   End
End
Attribute VB_Name = "frmbuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdmodificar_Click()
If Len(Trim(txtbuscartexto.Text)) = 0 Then
    MsgBox "Seleccione el reactivo a modificar", vbInformation, "Laboratorios el Puente "
Else
    rsReactivos.Find "NombreReactivos = '" & txtbuscartexto.Text & "'", , , 1
    If rsReactivos.BOF = False And rsReactivos.EOF = False Then
        'cargar datos a las cajas de texto
        txtnombre.Text = rsReactivos.Fields("NombreReactivos")
        txtcantidad.Text = rsReactivos.Fields("NúmeroReactivos")
        txtfecha.Text = rsReactivos.Fields("FechaExpiración")
        txtmarca.Text = rsReactivos.Fields("Marca")
        Nombre = txtnombre.Text
        Cantidad = txtcantidad.Text
        Fecha = txtfecha.Text
        Marca = txtmarca.Text
        If MsgBox("¿Desea modificar el reactivo seleccionado?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
            Adodc2.RecordSource = "select * from Administrador where Nombre =  '" + txtusuario.Text + "'"
            Adodc2.Refresh
            If Adodc2.Recordset.EOF Then
                MsgBox "Solo administradores pueden modificar reactivos", vbInformation, "Laboratorios el Puente "
            Else
                Verificar = 0
                frmregistrar.Show
                Unload Me
            End If
        End If
    End If
End If
End Sub

Private Sub cmdregistrar_Click()
    If MsgBox("¿Desea registrar un nuevo reactivo?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Nombre = ""
        Cantidad = ""
        Fecha = ""
        Marca = ""
        Verificar = 1
        Adodc2.RecordSource = "select * from Administrador where Nombre =  '" + txtusuario.Text + "'"
        Adodc2.Refresh
        If Adodc2.Recordset.EOF Then
            MsgBox "Solo administradores pueden registrar reactivos", vbInformation, "Laboratorios el Puente "
        Else
            frmregistrar.Show
            
            Unload Me
        End If
    End If
End Sub

Private Sub cmdusuarios_Click()
    If MsgBox("¿Desea administrar usuarios?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Adodc2.RecordSource = "select * from Administrador where Nombre =  '" + txtusuario.Text + "'"
        Adodc2.Refresh
        If Adodc2.Recordset.EOF Then
            MsgBox "Solo administradores pueden añadir usuarios", vbInformation, "Laboratorios el Puente "
        Else
            Unload Me
            frmdatos.Show
            End If
    End If
End Sub

Private Sub Command1_Click()
    If Len(Trim(txtbuscartexto.Text)) = 0 Then
        MsgBox "Seleccione un reactivo", vbInformation, "Laboratorios el Puente"
    Else
       If txtusuario.Text = "" Then
            MsgBox "Seleccione un usuario", vbInformation, "Laboratorios el Puente"
        Else
            rsReactivos.Find "NombreReactivos = '" & txtbuscartexto.Text & "'", , , 1
            If rsReactivos.BOF = False And rsReactivos.EOF = False Then
            'cargar datos a las cajas de texto
                txtnombre.Text = rsReactivos.Fields("NombreReactivos")
                txtcantidad.Text = rsReactivos.Fields("NúmeroReactivos")
                txtfecha.Text = rsReactivos.Fields("FechaExpiración")
                txtmarca.Text = rsReactivos.Fields("Marca")
                Nombre = txtnombre.Text
                Cantidad = txtcantidad.Text
                Fecha = txtfecha.Text
                Marca = txtmarca.Text
                If txtcantidad = 0 Then
                        MsgBox "Cantidad insufuciente registre más reactivos", vbInformation, "Laboratorios el Puente"
                Else
                If txtcantidad.Text < 25 Then
                    MsgBox "La cantidad restante de reactivos es inferior a 25 por favor registre más reactivos", vbInformation, "Laboratorios el Puente"
                Else
                    If MsgBox("¿Desea utilizar el reactivo?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
                        frmuso.Show
                        Unload Me
                    End If
                End If
                frmuso.Show
                End If
            End If
    End If
    End If

End Sub


Private Sub Command2_Click()
    If MsgBox("¿Desea salir?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Command3_Click()
If Len(Trim(txtbuscartexto.Text)) = 0 Then
    MsgBox "Seleccione el reactivo a eliminar", vbInformation, "Laboratorios el Puente "
Else
    rsReactivos.Find "NombreReactivos = '" & txtbuscartexto.Text & "'", , , 1
    If rsReactivos.BOF = False And rsReactivos.EOF = False Then
        If MsgBox("¿Desea eliminar el reactivo seleccionado?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
            Adodc2.RecordSource = "select * from Administrador where Nombre =  '" + txtusuario.Text + "'"
            Adodc2.Refresh
            If Adodc2.Recordset.EOF Then
                MsgBox "Solo administradores pueden eliminar reactivos", vbInformation, "Laboratorios el Puente "
            Else
                rsReactivos.Delete
                MsgBox "Reactivo eliminado", vbInformation, "Laboratorios el Puente "
                Unload Me
                frmbuscar.Show
                DataGrid1.Refresh
            End If
        End If
    End If
End If

End Sub

Private Sub Command4_Click()
    If MsgBox("¿Desea visualizar los reportes de uso?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
                        DataReport1.Show
                    End If
    
End Sub

Private Sub Command5_Click()
    If MsgBox("¿Desea cambiar de usuario?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        frminicio.Show
        Unload Me
    End If
End Sub

Private Sub DataGrid1_Click()
    txtbuscartexto.Text = DataGrid1.Columns(1).Text
    lblbuscar.Caption = DataGrid1.Columns(1).Text
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & "/logo5.ico")
    Label3.ForeColor = RGB(69, 110, 174)
    Label2.ForeColor = RGB(69, 110, 174)
    'lblbuscar.ForeColor = RGB(69, 110, 174)
    'lblusuario.ForeColor = RGB(69, 110, 174)
    'lblusuario.ForeColor = RGB(69, 110, 174)
    txtusuario.Text = Usuario
    lblusuario.Caption = txtusuario.Text
    TablaReactivos
    formato
End Sub
Sub formato()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 2300
    DataGrid1.Columns(2).Width = 2300
    DataGrid1.Columns(3).Width = 2500
    DataGrid1.Columns(4).Width = 2500
    DataGrid1.Columns(5).Width = 0
    DataGrid1.ForeColor = RGB(69, 110, 174)
End Sub

