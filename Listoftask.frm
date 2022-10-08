VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Listoftask 
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   2
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DENGELE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   7920
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid tasklist 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   12303
      _Version        =   393216
   End
   Begin VB.Label bilgi 
      Caption         =   "Ýþlem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Listoftask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ctr, ctl, nr, nl, wr, wl As Integer

Private Sub Command1_Click()

'n iþ sayýsý
'c çevrim süresi
'CTR, CTL istasyonda kullanýlan mevcut süre(sað ve sol)
'NR, NL þu ana kadar açýlan istasyon sayýsý
'WR, WL sol veya sað istasyon müsait mi deðil mi?
'k istenilen þartlar gerçekleþirse 1 arttýrýlacak
'Dim n, c, ctr, ctl, nr, nl, wr, wl, k, gx, gy, giden, gdon, mini, enkart, ddon, dmini, secilen, secilen2, bl As Integer
Dim n, c, k, gx, gy, giden, gdon, mini, enkart, ddon, dmini, secilen, secilen2, bl As Integer
Dim i As Integer
Dim btaskip(3000), isi(3000), ti(3000), pst(3000)
' xi(), fi(), satir(), nupi(), tsksd()

Call veri_ac(False, False)
Call tablo_ac("select * from maintable where name='" & isim & "'")

ctr = 0
ctl = 0
nr = 1
nl = 1
wr = 0
wl = 0
k = 0 ' Bulunan deðer oldukça  1 arttýracak
gx = 0
gy = 0
c = InputBox("Lütfen bir çevrim süresi giriniz", "Çevrim Süresi")
n = sayisi


Adim1:
'MsgBox n, vbExclamation
If n <= 0 Then
    bilgi.Caption = "Ýþlem: TAMAM"
Else
    bilgi.Caption = "Ýþlem: " & n & ". model hesaplanýyor..."
End If
If n = 0 Then
GoTo son
ElseIf (wr = 0 And ctr <= ctl) Or (wr = 0 And wl = 1) Then GoTo Adim2
ElseIf (wl = 0 And ctl < ctr) Or (wr = 1 And wl = 0) Then GoTo Adim3
Else
GoTo Adim6
End If

Adim2:
'Adim2***********************************************************
For i = 1 To sayisi
pst(i) = Empty
Next

i = 1
tablo.MoveFirst
k = 0
dmini = 0
Do While Not tablo.EOF
If tasklist.TextMatrix(i, 5) = 0 And tasklist.TextMatrix(i, 6) = 0 And (tasklist.TextMatrix(i, 2) = "R" Or tasklist.TextMatrix(i, 2) = "E") Then
    k = k + 1
    bulunanlar(k) = i
    Call bulhesap(i)
    If ctr >= Val(enb(i)) Then
        pst(i) = ctr
    Else
        pst(i) = Val(enb(i))
    End If
    
    End If
    tablo.MoveNext
    
    
     i = i + 1
    Loop
    If k >= 1 Then '***
    mini = 1453
    
    For gdon = 1 To k
        If Val(pst(bulunanlar(gdon))) <= mini Then
            mini = Val(pst(bulunanlar(gdon)))
         End If
    Next 'gdon
    
    enkart = 0
    For gdon = 1 To k
            If Val(pst(bulunanlar(gdon))) = Val(mini) Then
            enkart = enkart + 1
            enk(enkart) = Val(bulunanlar(gdon))
         End If
    Next 'gdon
     
     If enkart >= 1 Then ' 2-1
       
        dmini = 0
        For ddon = 1 To enkart
                If tasklist.TextMatrix(enk(ddon), 1) >= dmini Then
                    dmini = tasklist.TextMatrix(enk(ddon), 1)
                End If
        Next 'enkart
        
        For ddon = 1 To enkart
                If tasklist.TextMatrix(enk(ddon), 1) = dmini Then
                    secilen = Val(enk(ddon))
                    ddon = enkart + 1
                End If
        Next 'enkart
                      
        End If
        
        If IsNumeric(tasklist.TextMatrix(secilen, 4)) = True Then
        secilen2 = tasklist.TextMatrix(secilen, 4)
        GoTo Adim4
        End If
        
        
        If (Val(tasklist.TextMatrix(secilen, 1)) + Val(pst(secilen))) <= Val(c) Then
                tasklist.TextMatrix(secilen, 5) = 1
                tasklist.TextMatrix(secilen, 7) = nr
                tasklist.TextMatrix(secilen, 8) = 1
                ctr = Val(pst(secilen)) + tasklist.TextMatrix(secilen, 1)
                tasklist.TextMatrix(secilen, 9) = ctr
                n = n - 1
                
            For bl = 1 To sayisi
        
                If Val(matriks.matrikslist.TextMatrix(secilen, bl)) = 1 Then
                    tasklist.TextMatrix(bl, 6) = Val(tasklist.TextMatrix(bl, 6)) - 1
                End If
            Next 'bl
                
                GoTo Adim1
        Else
                wr = 1
                GoTo Adim1
        End If
                



Else ' ***
    If wl = 0 Then
      GoTo Adim3
    Else
      GoTo Adim6
    End If
    
End If ' ***




'Adim2***********************************************************



Adim3:

'Adim3***********************************************************
For i = 1 To sayisi
pst(i) = Empty
Next

i = 1
tablo.MoveFirst
k = 0
dmini = 0

Do While Not tablo.EOF
If tasklist.TextMatrix(i, 5) = 0 And tasklist.TextMatrix(i, 6) = 0 And (tasklist.TextMatrix(i, 2) = "L" Or tasklist.TextMatrix(i, 2) = "E") Then
    k = k + 1
    bulunanlar(k) = i
    Call bulhesap3(i)
    If ctl >= Val(enb(i)) Then
        pst(i) = ctl
    Else
        pst(i) = Val(enb(i))
    End If
    
    
    
    
    End If
    tablo.MoveNext
        
    i = i + 1
    Loop
    
    If k >= 1 Then '***
    mini = 1453
    
    For gdon = 1 To k
        If Val(pst(bulunanlar(gdon))) <= mini Then
             mini = pst(bulunanlar(gdon))
         End If
    Next 'gdon
    
    enkart = 0
    
    For gdon = 1 To k
            If Val(pst(bulunanlar(gdon))) = Val(mini) Then
            enkart = enkart + 1
                enk(enkart) = bulunanlar(gdon)
            End If
    Next 'gdon
     
     If enkart >= 1 Then ' 2-1
       
        dmini = 0
        For ddon = 1 To enkart
                If tasklist.TextMatrix(enk(ddon), 1) >= dmini Then
                    dmini = tasklist.TextMatrix(enk(ddon), 1)
                End If
        Next 'enkart
        
        For ddon = 1 To enkart
                If tasklist.TextMatrix(enk(ddon), 1) = dmini Then
                    secilen = Val(enk(ddon))
                    ddon = enkart + 1
                End If
        Next 'enkart
                      
        End If
        
        If IsNumeric(tasklist.TextMatrix(secilen, 4)) = True Then
            secilen2 = tasklist.TextMatrix(secilen, 4)
            GoTo Adim5
        End If
        
        If (Val(tasklist.TextMatrix(secilen, 1)) + Val(pst(secilen))) <= Val(c) Then
                tasklist.TextMatrix(secilen, 5) = 1
                'MsgBox Val(tasklist.TextMatrix(secilen, 1)) + ctl
                tasklist.TextMatrix(secilen, 7) = nl
                tasklist.TextMatrix(secilen, 8) = 2
                ctl = Val(pst(secilen)) + tasklist.TextMatrix(secilen, 1)
                tasklist.TextMatrix(secilen, 9) = ctl
                n = n - 1
                
                                
            For bl = 1 To sayisi
                If Val(matriks.matrikslist.TextMatrix(secilen, bl)) = 1 Then
                    tasklist.TextMatrix(bl, 6) = Val(tasklist.TextMatrix(bl, 6)) - 1
                End If
            Next 'n
                
                GoTo Adim1
                
        Else
                wl = 1
                GoTo Adim1
        End If
                
    


Else ' ***
    If wr = 0 Then
      GoTo Adim2
    Else
      GoTo Adim6
    End If
    
End If ' ***

'Adim3***********************************************************


Adim4:
'Adim4***********************************************************
If ctl > ctr Then
    pst(secilen) = ctl
    pst(secilen2) = ctl
Else
    pst(secilen) = ctr
    pst(secilen2) = ctr
End If

If (Val(pst(secilen)) + Val(tasklist.TextMatrix(secilen, 1)) <= c) And (Val(pst(secilen2)) + Val(tasklist.TextMatrix(secilen2, 1)) <= c) Then
    tasklist.TextMatrix(secilen, 5) = 1
    tasklist.TextMatrix(secilen, 7) = nr
    tasklist.TextMatrix(secilen, 8) = 1
    ctr = Val(pst(secilen)) + tasklist.TextMatrix(secilen, 1)
    tasklist.TextMatrix(secilen, 9) = ctr
    
    tasklist.TextMatrix(secilen2, 5) = 1
    tasklist.TextMatrix(secilen2, 7) = nl
    tasklist.TextMatrix(secilen2, 8) = 2
    ctl = Val(pst(secilen2)) + tasklist.TextMatrix(secilen2, 1)
    tasklist.TextMatrix(secilen2, 9) = ctl
Else
    nr = nr + 1
    nl = nl + 1
    ctr = 0
    ctl = 0
    wr = 0
    wl = 0
    
    tasklist.TextMatrix(secilen, 5) = 1
    tasklist.TextMatrix(secilen, 7) = nr
    tasklist.TextMatrix(secilen, 8) = 1
    ctr = tasklist.TextMatrix(secilen, 1)
    tasklist.TextMatrix(secilen, 9) = ctr
    tasklist.TextMatrix(secilen2, 5) = 1
    tasklist.TextMatrix(secilen2, 7) = nl
    tasklist.TextMatrix(secilen2, 8) = 2
    ctl = tasklist.TextMatrix(secilen2, 1)
    tasklist.TextMatrix(secilen2, 9) = ctl
End If
n = n - 2
For bl = 1 To sayisi
    If Val(matriks.matrikslist.TextMatrix(secilen, bl)) = 1 Then
       tasklist.TextMatrix(bl, 6) = Val(tasklist.TextMatrix(bl, 6)) - 1
    End If
    
    If Val(matriks.matrikslist.TextMatrix(secilen2, bl)) = 1 Then
       tasklist.TextMatrix(bl, 6) = Val(tasklist.TextMatrix(bl, 6)) - 1
    End If
    
Next 'bl

GoTo Adim1

'Adim4***********************************************************


Adim5:
'Adim5***********************************************************
If ctl > ctr Then
    pst(secilen) = ctl
    pst(secilen2) = ctl
Else
    pst(secilen) = ctr
    pst(secilen2) = ctr
End If

If (Val(pst(secilen)) + Val(tasklist.TextMatrix(secilen, 1)) <= c) And (Val(pst(secilen2)) + Val(tasklist.TextMatrix(secilen2, 1)) <= c) Then
    tasklist.TextMatrix(secilen, 5) = 1
    tasklist.TextMatrix(secilen, 7) = nl
    tasklist.TextMatrix(secilen, 8) = 2
    ctl = Val(pst(secilen)) + tasklist.TextMatrix(secilen, 1)
    tasklist.TextMatrix(secilen, 9) = ctl
    
    tasklist.TextMatrix(secilen2, 5) = 1
    tasklist.TextMatrix(secilen2, 7) = nr
    tasklist.TextMatrix(secilen2, 8) = 1
    ctr = Val(pst(secilen2)) + tasklist.TextMatrix(secilen2, 1)
    tasklist.TextMatrix(secilen2, 9) = ctr
Else
    nr = nr + 1
    nl = nl + 1
    ctr = 0
    ctl = 0
    wr = 0
    wl = 0
    
    tasklist.TextMatrix(secilen, 5) = 1
    tasklist.TextMatrix(secilen, 7) = nl
    tasklist.TextMatrix(secilen, 8) = 2
    ctl = tasklist.TextMatrix(secilen, 1)
    tasklist.TextMatrix(secilen, 9) = ctl
    
    tasklist.TextMatrix(secilen2, 5) = 1
    tasklist.TextMatrix(secilen2, 7) = nr
    tasklist.TextMatrix(secilen2, 8) = 1
    ctr = tasklist.TextMatrix(secilen2, 1)
    tasklist.TextMatrix(secilen2, 9) = ctr
End If
n = n - 2
For bl = 1 To sayisi
    If Val(matriks.matrikslist.TextMatrix(secilen, bl)) = 1 Then
       tasklist.TextMatrix(bl, 6) = Val(tasklist.TextMatrix(bl, 6)) - 1
    End If
    
    If Val(matriks.matrikslist.TextMatrix(secilen2, bl)) = 1 Then
       tasklist.TextMatrix(bl, 6) = Val(tasklist.TextMatrix(bl, 6)) - 1
    End If
    
Next 'bl

GoTo Adim1
'Adim5***********************************************************



Adim6:
'Adim6***********************************************************
nr = nr + 1
nl = nl + 1
ctr = 0
ctl = 0
wr = 0
wl = 0
GoTo Adim1
'Adim6***********************************************************


son:
MsgBox "ÝÞLEM BÝTMÝÞTÝR", vbInformation
Exit Sub
tablo.Close
veri.Close
'
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim a, i As Long
If tasklist.Rows = 1 Then ' Eðer msflexgrid boþsa aktarma yapma
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

For a = 1 To tasklist.Cols
For i = 1 To tasklist.Rows
xlSheet.Cells(1, a).Font.Bold = True ' Excel' e aktarýlan kayýtlarýn ilk sütunu baþlýklar olduðu için karakterleri bold
xlSheet.Cells(i, a) = Replace(tasklist.TextMatrix(i - 1, a - 1), ",", ".")
xlSheet.Cells(i, 4).NumberFormat = "@" 'burada excel tarafna aktarýlan kaydýn formatýný belirleyebilirsiniz
Next i
Next a

Screen.MousePointer = 0
xlBook.Application.Visible = True

End Sub

Private Sub Form_Load()
Me.Caption = "Model :" & isim
Call yukle
End Sub

Sub yukle()
On Local Error Resume Next
Dim x As Integer
Call veri_ac(False, False)
Call tablo_ac("Select * from maintable where name='" & isim & "'")
tasklist.Cols = 10
tasklist.Rows = 1
tasklist.ColWidth(0) = 500
tasklist.ColWidth(1) = 500
tasklist.ColWidth(2) = 500
tasklist.ColWidth(3) = 500
tasklist.ColWidth(4) = 500
tasklist.ColWidth(5) = 500
tasklist.ColWidth(6) = 500
tasklist.ColWidth(7) = 500
tasklist.ColWidth(8) = 500
tasklist.ColWidth(9) = 500

tasklist.TextMatrix(0, 0) = "TASK"
tasklist.TextMatrix(0, 1) = "TASK TIME"
tasklist.TextMatrix(0, 2) = "TASK SIDE"
tasklist.TextMatrix(0, 3) = "TASK IP"
tasklist.TextMatrix(0, 4) = " SYN "
tasklist.TextMatrix(0, 5) = " X(i) "
tasklist.TextMatrix(0, 6) = " NUP(i) "
tasklist.TextMatrix(0, 7) = " IS(i) "
tasklist.TextMatrix(0, 8) = " T(i) "
tasklist.TextMatrix(0, 9) = " F(i) "


x = 0
Do While Not tablo.EOF
x = x + 1
tasklist.AddItem ""
tasklist.TextMatrix(x, 0) = tablo("taskno")
tasklist.TextMatrix(x, 1) = tablo("tasktime")
tasklist.TextMatrix(x, 2) = tablo("taskside")
tasklist.TextMatrix(x, 3) = tablo("taskip")
tasklist.TextMatrix(x, 4) = tablo("tasksyn")
tasklist.TextMatrix(x, 5) = tablo("xi")
tasklist.TextMatrix(x, 6) = tablo("nupi")
tasklist.TextMatrix(x, 7) = tablo("isi")
tasklist.TextMatrix(x, 8) = tablo("ti")
tasklist.TextMatrix(x, 9) = tablo("fi")
tablo.MoveNext
Loop
tablo.Close
veri.Close
End Sub


Private Sub tasklist_DblClick()
'MsgBox tasklist.TextMatrix(tasklist.Row, tasklist.Col)
If tasklist.Col = 5 Or tasklist.Col = 6 Or tasklist.Col = 7 Or tasklist.Col = 8 Or tasklist.Col = 9 Then
    MsgBox "Bu deðerler dýþarýdan deðer almaz", vbCritical, "Uyarý"
ElseIf tasklist.Col = 2 Then
        If tasklist.TextMatrix(tasklist.Row, tasklist.Col) = Empty Or tasklist.TextMatrix(tasklist.Row, tasklist.Col) = " " Then
        tasklist.TextMatrix(tasklist.Row, tasklist.Col) = "L"
        ElseIf tasklist.TextMatrix(tasklist.Row, tasklist.Col) = "L" Then
        tasklist.TextMatrix(tasklist.Row, tasklist.Col) = "R"
        ElseIf tasklist.TextMatrix(tasklist.Row, tasklist.Col) = "R" Then
        tasklist.TextMatrix(tasklist.Row, tasklist.Col) = "E"
        Else
        tasklist.TextMatrix(tasklist.Row, tasklist.Col) = " "
        End If
Else
    tasklist.TextMatrix(tasklist.Row, tasklist.Col) = Val(tasklist.TextMatrix(tasklist.Row, tasklist.Col)) + 1
End If
End Sub



Sub bulhesap(gln As Integer)

Dim sydir As Integer
enb(gln) = 0
For sydir = 1 To sayisi

    If matriks.matrikslist.TextMatrix(sydir, gln) = 1 Then
                If tasklist.TextMatrix(sydir, 7) = nl And tasklist.TextMatrix(sydir, 8) = 2 Then
                   gecici(gln, sydir) = tasklist.TextMatrix(sydir, 9)
                    If enb(gln) <= gecici(gln, sydir) Then
                        enb(gln) = gecici(gln, sydir)
                    End If
                   'MsgBox gecici(gln, sydir)
                   End If
        'MsgBox gln & " " & sydir
    End If
Next 'sydir
End Sub



Sub bulhesap3(gln As Integer)

Dim sydir As Integer
enb(gln) = 0
For sydir = 1 To sayisi
    
    If matriks.matrikslist.TextMatrix(sydir, gln) = 1 Then
                If tasklist.TextMatrix(sydir, 7) = nr And tasklist.TextMatrix(sydir, 8) = 1 Then
                   gecici(gln, sydir) = tasklist.TextMatrix(sydir, 9)
                   If enb(gln) <= gecici(gln, sydir) Then
                    enb(gln) = gecici(gln, sydir)
                   End If
                   'MsgBox gecici(gln, sydir)
                End If
        'MsgBox gln & " " & sydir
        End If
Next 'sydir
End Sub

