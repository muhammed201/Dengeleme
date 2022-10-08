VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MainMenu 
   Caption         =   "New Job"
   ClientHeight    =   3150
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Previous Jobs"
      Height          =   2415
      Left            =   4920
      TabIndex        =   6
      Top             =   360
      Width           =   5535
      Begin MSFlexGridLib.MSFlexGrid tasklist 
         Height          =   1455
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Job"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.TextBox numberoftasks 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox jobname 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Tasks"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Job Name"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Ýþlem Menüsü"
      Visible         =   0   'False
      Begin VB.Menu sil 
         Caption         =   "Modeli Sil"
         Index           =   1
      End
      Begin VB.Menu kopya 
         Caption         =   "Modeli Klonla"
         Index           =   0
      End
      Begin VB.Menu iptal 
         Caption         =   "Ýptal"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public hata As Boolean
Public newname As String
Private Sub Command1_Click()
On Local Error GoTo hata
Dim arttir, arttirx, arttiry As Integer

'*******KONTROL***********************************************************
If IsNumeric(numberoftasks.Text) = False Then
    MsgBox "Number of Task kýsmýna sayýsal bir deðer girmeniz gerekiyor", vbCritical
    Exit Sub
End If

If jobname.Text = Empty Or numberoftasks.Text = Empty Then
    MsgBox "Jobname veya Task of Number'ý boþ býrakamazsýnýz", vbCritical
    Exit Sub
End If

''*******KONTROL***********************************************************


Call veri_ac(False, False)
Call tablo_ac("Select * from maintable")


For arttir = 1 To numberoftasks
    tablo.AddNew
    tablo("name") = jobname.Text
    tablo("taskno") = arttir
    tablo.Update
Next ' numberoftasks
tablo.Close

Call tablo_ac("Select * from matriks")
For arttirx = 1 To numberoftasks
    For arttiry = 1 To numberoftasks
    tablo.AddNew
    tablo("name") = jobname.Text
    tablo("x") = arttirx
    tablo("y") = arttiry
    tablo.Update
    Next 'arttiry
Next 'arttirx




   MsgBox "Kayýt iþlemi gerçekleþtirilmiþtir", vbInformation
tablo.Close
veri.Close
isim = jobname.Text
sayisi = numberoftasks
Call taskyukle
hata:
If Err = 3022 Then
    MsgBox "Bu isimde önceden bir Model ismi kaydedilmiþtir lütfen yeniden deneyiniz", vbCritical, "Hata!"
End If
End Sub




Private Sub Form_Load()
Call taskyukle
End Sub

Sub taskyukle()
On Local Error Resume Next
Dim x As Integer
Call veri_ac(False, False)
Call tablo_ac("Select name from maintable group by name")
tasklist.Cols = 2
tasklist.Rows = 1

tasklist.TextMatrix(0, 0) = "TASK NUMBER"
tasklist.TextMatrix(0, 1) = "TASK NAME"


x = 0
Do While Not tablo.EOF
x = x + 1
tasklist.AddItem ""
tasklist.TextMatrix(x, 0) = x
tasklist.TextMatrix(x, 1) = tablo("name")

tablo.MoveNext
Loop
tablo.Close
veri.Close
End Sub




Private Sub kopya_Click(Index As Integer)
If hata = True Then
    hata = False
    Exit Sub
End If

Call matrikaydet

If hata = True Then
    Exit Sub
    hata = False
End If

Call mainkaydet
Call MainMenu.taskyukle
End Sub

Private Sub sil_Click(Index As Integer)
On Local Error Resume Next
Dim x As Integer
Dim sor As Integer
For x = 1 To (Val(tasklist.Rows) - 1)
    tasklist.Row = x
    If tasklist.CellBackColor = 4326608 Then
    sor = MsgBox(tasklist.TextMatrix(x, 1) & " isimli modeli silmek istediðinizden eminmisiniz?", vbYesNo)
            If sor = 6 Then
                Call veri_ac(False, False)
                Call tablo_ac("Select * from maintable where name='" & tasklist.TextMatrix(x, 1) & "'")
                Call tablo_ac2("Select * from matriks where name='" & tasklist.TextMatrix(x, 1) & "'")
                
                Do While Not tablo.EOF
                tablo.Delete
                tablo.MoveNext
                Loop
                
                
                Do While Not tablo2.EOF
                tablo2.Delete
                tablo2.MoveNext
                Loop
                
                MsgBox (tasklist.TextMatrix(x, 1) & " isimli model silinmiþtir")
                tablo.Close
                tablo2.Close
                Call taskyukle
            Else
                MsgBox "Ýptal Ettiniz"
        End If
    End If
Next
End Sub

Private Sub tasklist_Click()
On Local Error Resume Next
Dim x, Y As Integer
Y = tasklist.Row
isim = tasklist.TextMatrix(tasklist.Row, 1)
For x = 1 To tasklist.Rows - 1
    tasklist.Row = x
    tasklist.Col = 0
    tasklist.CellBackColor = tasklist.BackColor
    tasklist.Col = 1
    tasklist.CellBackColor = tasklist.BackColor
  
Next

If tasklist.Row = 0 Then
    tasklist.Col = 0
    tasklist.CellBackColor = 12632256
    tasklist.Col = 1
    tasklist.CellBackColor = 12632256
    Exit Sub
End If
    
    tasklist.Row = Y
    tasklist.Col = 0
    tasklist.CellBackColor = 4326608
    tasklist.Col = 1
    tasklist.CellBackColor = 4326608
    secilisatir = Y

End Sub

Private Sub tasklist_DblClick()
Call sayibul(tasklist.TextMatrix(tasklist.Row, 1))
isim = tasklist.TextMatrix(tasklist.Row, 1)
MsgBox tasklist.TextMatrix(tasklist.Row, 1) & " Modeli çaðýrýldý", vbInformation
matriks.Show
End Sub

Sub sayibul(modeladi As String)
Call veri_ac(False, False)
Call tablo_ac("Select * from maintable where name='" & modeladi & "'")
sayisi = 0
Do While Not tablo.EOF
  sayisi = sayisi + 1
tablo.MoveNext
Loop

End Sub

Private Sub tasklist_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If menu.Visible = True Then GoTo son
If Button = 2 Then
PopupMenu menu
End If
son:
End Sub

Sub mainkaydet()
Call veri_ac(False, False)
Call tablo_ac("Select * from maintable where name='" & isim & "'")
Call tablo_ac2("Select * from maintable")
tablo.MoveFirst
Do While Not tablo.EOF
    tablo2.AddNew
    tablo2("name") = newname
    tablo2("taskno") = tablo("taskno")
    tablo2("tasktime") = tablo("tasktime")
    tablo2("taskside") = tablo("taskside")
    tablo2("taskip") = tablo("taskip")
    tablo2("tasksyn") = tablo("tasksyn")
    tablo2("xi") = tablo("xi")
    tablo2("nupi") = tablo("nupi")
    tablo2("isi") = tablo("isi")
    tablo2("ti") = tablo("ti")
    tablo2("fi") = tablo("fi")
    tablo2.Update
    tablo.MoveNext
Loop

   MsgBox "Seçilen Modelin Verileri Baþarýlý Bir Þekilde Kolonlanmýþtýr..", vbInformation
tablo.Close
veri.Close

End Sub

Sub matrikaydet()
Call veri_ac(False, False)
Call tablo_ac("Select * from matriks where name='" & isim & "'")
Call tablo_ac2("Select * from matriks")
newname = InputBox("Lütfen yeni bir Model ismi giriniz..", "New Model", tablo("name") & "_" & Now)
If newname = Empty Or newname = " " Then
 MsgBox "Ýptal edildi", vbExclamation
 hata = True
 Exit Sub
End If

tablo.MoveFirst
Do While Not tablo.EOF
    
    tablo2.AddNew
    tablo2("name") = newname
    tablo2("x") = tablo("x")
    tablo2("y") = tablo("y")
    tablo2("deger") = tablo("deger")
    tablo2.Update
tablo.MoveNext
Loop

MsgBox "Seçilen Modelin Matriksi Baþarýyla Klonlanmýþtýr...", vbInformation
tablo.Close
tablo2.Close
veri.Close
End Sub
