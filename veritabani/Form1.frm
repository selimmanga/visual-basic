VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   18705
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   18705
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6495
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton listele 
      Caption         =   "Listele"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label yazi 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   2
      Top             =   2040
      Width           =   10815
   End
   Begin VB.Menu alis 
      Caption         =   "ALIÞ"
      Begin VB.Menu firta 
         Caption         =   "Firma Tanýmla"
      End
      Begin VB.Menu sil 
         Caption         =   "Firma Sil"
      End
   End
   Begin VB.Menu satis 
      Caption         =   "SATIÞ"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub firta_Click()
Form2.Show
End Sub


Private Sub listele_Click()
Dim liste As New ADODB.Recordset ' Recordset: veritabaný içinde dolaþan imleç
Dim say As Integer
Dim metin As String

say = 0

'Çoklu kullanýcýlar için adOpenStatic kullanýyoruz
liste.Open "select * from tablo1", "dsn=masa", adOpenStatic

With grid
.Rows = liste.RecordCount + 1
.Cols = 3

For t = 1 To liste.RecordCount
.TextMatrix(t, 0) = liste("ad")
.TextMatrix(t, 1) = liste("soyad")
.TextMatrix(t, 2) = liste("yas")
say = say + CInt(liste("yas"))
'Yeni kayýta geçmesi için bunu yazmamýz gerekir yoksa hepsine ayný veriyi yazdýrýr
liste.MoveNext
Next

For p = 1 To liste.RecordCount
If CInt(.TextMatrix(p, 2)) > CInt(say / liste.RecordCount) Then
.Row = p
.Col = 2
.CellBackColor = vbRed
End If
Next

End With

metin = "Listedeki isimlerin yaþ ortalamasý " & CInt(say / liste.RecordCount) & "'dir."
yazi.Caption = metin
End Sub
