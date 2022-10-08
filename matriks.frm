VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form matriks 
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   10680
      TabIndex        =   7
      Top             =   6360
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid toplam 
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   7080
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   25
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   8160
      Width           =   13095
      Begin VB.Label bilgi 
         AutoSize        =   -1  'True
         Caption         =   "Bilgi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Problemi Farklý Kaydet"
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   6360
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Veri Listesini Aç"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   6360
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excel'e aktar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid matrikslist 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   10821
      _Version        =   393216
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "matriks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public say As Integer
Public newname As String
Public hata As Boolean

Private Sub Command1_Click()
On Error Resume Next
Dim a, i As Long
If matrikslist.Rows = 1 Then ' Eðer msflexgrid boþsa aktarma yapma
Title = "Hata !"
msg = "Microsoft Excel' e aktarýlacak herhangi bir kayýt bulunamadý."
Answer = MsgBox(msg, vbCritical, Title)
Exit Sub
End If
Screen.MousePointer = 11
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)

For a = 1 To matrikslist.Cols
For i = 1 To matrikslist.Rows
xlSheet.Cells(1, a).Font.Bold = True ' Excel' e aktarýlan kayýtlarýn ilk sütunu baþlýklar olduðu için karakterleri bold
xlSheet.Cells(i, a) = Replace(matrikslist.TextMatrix(i - 1, a - 1), ",", ".")
'xlSheet.Cells(i, 3).NumberFormat = "#,##0.00" 'burada excel tarafna aktarýlan kaydýn formatýný belirleyebilirsiniz
Next i
Next a

Screen.MousePointer = 0
xlBook.Application.Visible = True

End Sub

Private Sub Command2_Click()
Me.Hide
Listoftask.Show
End Sub






Private Sub Command4_Click()
Call hesapla
End Sub


Private Sub Form_Load()
Me.Caption = "Model: " & isim
hata = False
Call matriksyukle
Call hesapla
Call diyagonal
End Sub


Sub matriksyukle()
On Local Error Resume Next
Call sayibuldur
MsgBox isim
Dim x, hizzala As Integer

Call veri_ac(False, False)
Call tablo_ac("Select * from matriks where name='" & isim & "'")
MsgBox say, vbCritical
matrikslist.Cols = say + 1
matrikslist.Rows = 1

For hizzala = 0 To say
matrikslist.ColWidth(hizzala) = 200 * Len(say)
Next

x = 0
Do While Not tablo.EOF
If x = 346 Then MsgBox "Dur"
    x = x + 1
    matrikslist.AddItem ""
    matrikslist.TextMatrix(x, 0) = x
    matrikslist.TextMatrix(0, x) = x
    If x >= say Then Exit Do
    tablo.MoveNext
Loop
tablo.Close
veri.Close

Call veridoldur
End Sub

Sub sayibuldur()
Call veri_ac(False, False)
Call tablo_ac("select * from maintable where name='" & isim & "'")
    say = 0
    Do While Not tablo.EOF
    say = say + 1
    tablo.MoveNext
    Loop
tablo.Close
veri.Close
End Sub

Sub veridoldur()
Dim donx, dony As Integer
Call veri_ac(False, False)
Call tablo_ac("Select * from matriks  where name='" & isim & "' order by x, y")
For donx = 1 To say
For dony = 1 To say
matrikslist.TextMatrix(dony, donx) = tablo("deger")

matrikslist.Col = donx
matrikslist.Row = dony

If tablo("deger") = 1 Or tablo("deger") = "1" Then
    matrikslist.CellBackColor = vbRed
    matrikslist.CellForeColor = vbYellow
Else
    matrikslist.CellBackColor = vbWhite
    matrikslist.CellForeColor = vbBlack
End If

tablo.MoveNext
Next 'dony
Next 'donx
End Sub




'Private Sub matrikslist_Click() ' Seçtiðimiz matriksin saðdan soldan ok çýkartarak kordinantlarýný göstermesini saðlayan scriptler
'Dim bx, by As Integer
'For bx = 1 To matrikslist.Col
'matrikslist.Col = bx
'matrikslist.CellBackColor = 13563087
'Next 'bx
'For by = 1 To matrikslist.Row
'matrikslist.Row = by
'matrikslist.CellBackColor = 13563087
'Next 'by
'End Sub

Private Sub matrikslist_DblClick()

On Local Error Resume Next

If matrikslist.TextMatrix(matrikslist.Row, matrikslist.Col) = 0 Then
    matrikslist.TextMatrix(matrikslist.Row, matrikslist.Col) = 1
    matrikslist.CellBackColor = vbRed
    matrikslist.CellForeColor = vbYellow
    
Else
    matrikslist.TextMatrix(matrikslist.Row, matrikslist.Col) = 0
    matrikslist.CellBackColor = vbWhite
    matrikslist.CellForeColor = vbBlack
End If

End Sub


Sub hesapla()
On Local Error Resume Next
Dim bsatir, bsutun As Integer
toplam.Cols = sayisi + 1
toplam.Rows = 2

toplam.Clear
toplam.TextMatrix(0, 0) = "TASK NO"
toplam.TextMatrix(1, 0) = "TOPLAM"
For bsutun = 1 To sayisi
    toplam.TextMatrix(0, bsutun) = bsutun
    For bsatir = 1 To sayisi
        toplam.TextMatrix(1, bsutun) = Val(matrikslist.TextMatrix(bsatir, bsutun)) + Val(toplam.TextMatrix(1, bsutun))
        toplam.ColWidth(bsutun) = 150 * Len(sayisi)
    Next 'bsatir
Next 'bsutun
End Sub

Sub diyagonal()
On Local Error Resume Next
Dim diysay As Integer

For diysay = 1 To sayisi
    matrikslist.Col = diysay
    matrikslist.Row = diysay
    matrikslist.CellBackColor = 13420238
Next 'diysay

End Sub
