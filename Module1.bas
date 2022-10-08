Attribute VB_Name = "Module1"
Global veri As Database
Global veri2 As Database
Global tablo As Recordset
Global tablo2 As Recordset
Global isim As String
Global sayisi As Integer
Global gecici(3000, 3000)
Global enb(3000)
Global enk(3000)
Global bulunanlar(3000)

Sub veri_ac(X1 As Boolean, X2 As Boolean)
Set veri = Workspaces(0).OpenDatabase(App.Path & "\veri.mdb", X1, X2)
End Sub


Sub tablo_ac(sql As String)
Set tablo = veri.OpenRecordset(sql)
End Sub

Sub tablo_ac2(sql As String)
Set tablo2 = veri.OpenRecordset(sql)
End Sub



