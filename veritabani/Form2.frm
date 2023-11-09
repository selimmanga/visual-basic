VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   10110
   ClientLeft      =   1725
   ClientTop       =   825
   ClientWidth     =   12600
   LinkTopic       =   "Form2"
   ScaleHeight     =   10110
   ScaleWidth      =   12600
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   8415
      Begin VB.CommandButton kaydet 
         Caption         =   "Kaydet"
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
         Left            =   2280
         TabIndex        =   7
         Top             =   4440
         Width           =   3975
      End
      Begin VB.TextBox yas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox soyad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   2160
         Width           =   4695
      End
      Begin VB.TextBox ad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Yaþ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Soyad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Ad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub kaydet_Click()
Dim kaydet As New ADODB.Recordset

'Çoklu kullanýcýlar için adOpenStatic kullanýyoruz
kaydet.Open "insert into tablo1(ad, soyad, yas) values ('" & ad.Text & "', '" & soyad.Text & "','" & yas.Text & "')", "dsn=masa", adOpenStatic
aa = MsgBox("Kayýt baþarýlý", vbInformation, "Bilgi")
End Sub
