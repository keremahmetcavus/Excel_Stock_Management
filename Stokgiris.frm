VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Stokgiris 
   Caption         =   "Stok Giri�"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "Stokgiris.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Stokgiris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BUTTON_STOKGIRIS_Click()
    Dim SEARCH_KISAKOD As Range
    Dim SEARCH_SPEC As Range
    
    ''STOCKCONTROL.Activate
    Sheets("STOKLAR").Activate
    
    
    K�sakod_Text = TB_KISAKOD.Text '' k�sakod C55
    Spec_Text = TB_SPEC.Text ''spec numaras� 7I0087
    
    Set SEARCH_KISAKOD = Sheets("STOKLAR").Range("A1", Range("A1").End(xlDown))
    SEARCH_KISAKOD.Select
    
    Set SEARCH_SPEC = Sheets("STOKLAR").Range("B1", Range("B1").End(xlDown))
    SEARCH_SPEC.Select
    Set BUL_KISAKOD = SEARCH_KISAKOD.Find(What:=K�sakod_Text, LookIn:=xlValues, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                
    Set BUL_SPEC = SEARCH_SPEC.Find(What:=Spec_Text, LookIn:=xlValues, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                
    EN_ALT = Sheets("STOKLAR").Range("a65536").End(3).Row + 1
    EN_KAYITALT = Sheets("KAYITLAR").Range("a65536").End(3).Row + 1
    
    If TB_ADET.Text = "" Or TB_KISAKOD.Text = "" Or TB_SPEC.Text = "" _
    Or TB_4.Text = "" Or CB_BOLINO.Text = "" Or CB_SORUMLU.Text = "" _
    Or CB_GIRISYAPAN.Text = "" Then
    MsgBox "L�tfen Zorunlu alanlar� doldurunuz", , UYARI!
    Exit Sub
    
    ElseIf Len(TB_4.Text) <> 10 Or Mid(TB_4.Text, 3, 1) <> "." Or Mid(TB_4.Text, 6, 1) <> "." Then
        MsgBox "L�tfen tarihi GG/AA/YYYY format�na uygun bir �ekilde giriniz.", , HATA!
        Cancel = True
        Exit Sub
        
    ElseIf Not IsDate(TB_4) Then
        MsgBox "L�tfen ge�erli bir tarih giriniz.", , HATA!
        Cancel = True
        Exit Sub
        
    Else
                
            If BUL_KISAKOD Is Nothing Or BUL_SPEC Is Nothing Then
                i = 1
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = K�sakod_Text
                i = i + 1 'i=2
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = Spec_Text
                i = i + 1 'i=3
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = CB_BOLINO.Text
                i = i + 1 'i=4
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = TB_ADET.Text
                i = i + 1 'i=5
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = TB_4.Text
                i = i + 1 'i=6
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = CB_SORUMLU.Text
                i = i + 1 'i=7
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = CB_GIRISYAPAN.Text
                i = i + 1 'i=8
                Sheets("STOKLAR").Cells(EN_ALT, i).Value = TB_NOT.Text
                i = 0
                
                i = 1
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = K�sakod_Text
                i = i + 1 'i=2
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Spec_Text
                i = i + 1 'i=3
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_BOLINO.Text
                i = i + 1 'i=4
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_ADET.Text
                i = i + 1 'i=5
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_4.Text
                i = i + 1 'i=6
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_SORUMLU.Text
                i = i + 1 'i=7
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_GIRISYAPAN.Text
                i = i + 1 'i=8
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_NOT.Text
                i = i + 1
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Now
                i = i + 1
                Sheets("KAYITLAR").Cells(EN_ALT, i).Value = "YEN� STOK OLU�TURMA"
                i = 0
                MsgBox Spec_Text + " Spec Numaral� " + K�sakod_Text + " �l��s�nden " & vbNewLine _
                & TB_ADET.Text + " Adet Stok Giri�i Yap�lm��t�r.", , "Stok Onayland�."
            Else
            
                Dim mesaj As String
                Dim cevap As Integer
                mesaj = Spec_Text + " Spec Numaral� " + K�sakod_Text + " �l��s�nden " & vbNewLine _
                & " Stok Mevcut. Stok giri�i yap�ls�n m�?"
                cevap = MsgBox(mesaj, vbYesNo)
                
                If cevap = vbYes Then
                
                    Do
                        Set BUL_KISAKOD = SEARCH_KISAKOD.Find(What:=K�sakod_Text, After:=BUL_KISAKOD)
                    Loop While BUL_KISAKOD.Offset(0, 1).Value <> BUL_SPEC.Text
                    
                    BUL_KISAKOD.Offset(0, 1).Select

                    Dim GIRIS_ADET As Integer
                    Dim BUL_ADET As Integer
                    Dim TOPLAM As Integer
                    
                    BUL_ADET = BUL_KISAKOD.Offset(0, 3).Value
                    GIRIS_ADET = TB_ADET.Text
                    
                    Dim MESAJ_2 As String
                    Dim CEVAP_2 As Integer
                    
                    MESAJ_2 = Spec_Text + " Spec Numaral� " + K�sakod_Text + " �l��s�nden " _
                    & BUL_ADET & " Adet Stok Mevcut." & vbNewLine & GIRIS_ADET & "  Adet Daha Stok Giri�i Yap�ls�n M�?"
                    
                    CEVAP_2 = MsgBox(MESAJ_2, vbYesNo)
                    
                    If CEVAP_2 = vbYes Then
                    
                    TOPLAM = BUL_ADET + GIRIS_ADET
                    BUL_KISAKOD.Offset(0, 3).Value = TOPLAM
                    BUL_KISAKODR = BUL_KISAKOD.Row
                    i = 3
                    Sheets("STOKLAR").Cells(BUL_KISAKODR, i).Value = CB_BOLINO.Text
                    i = 5
                    Sheets("STOKLAR").Cells(BUL_KISAKODR, i).Value = TB_4.Text
                    i = i + 1 'i=6
                    Sheets("STOKLAR").Cells(BUL_KISAKODR, i).Value = CB_SORUMLU.Text
                    i = i + 1 'i=7
                    Sheets("STOKLAR").Cells(BUL_KISAKODR, i).Value = CB_GIRISYAPAN.Text
                    i = i + 1 'i=8
                    Sheets("STOKLAR").Cells(BUL_KISAKODR, i).Value = TB_NOT.Text
                    i = 0
                        
                    
                    
                    
                    i = 1
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = K�sakod_Text
                    i = i + 1 'i=2
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Spec_Text
                    i = i + 1 'i=3
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_BOLINO.Text
                    i = i + 1 'i=4
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_ADET.Text
                    i = i + 1 'i=5
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_4.Text
                    i = i + 1 'i=6
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_SORUMLU.Text
                    i = i + 1 'i=7
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_GIRISYAPAN.Text
                    i = i + 1 'i=8
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_NOT.Text
                    i = i + 1
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Now
                    i = i + 1
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = "MEVCUT STO�A �LAVE"
                    i = 0
                
                
                    Else
                    i = 1
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = K�sakod_Text
                    i = i + 1 'i=2
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Spec_Text
                    i = i + 1 'i=3
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_BOLINO.Text
                    i = i + 1 'i=4
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_ADET.Text
                    i = i + 1 'i=5
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_4.Text
                    i = i + 1 'i=6
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_SORUMLU.Text
                    i = i + 1 'i=7
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_GIRISYAPAN.Text
                    i = i + 1 'i=8
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_NOT.Text
                    i = i + 1
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Now
                    i = i + 1
                    Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = "STOK G�R�� �PTAL�"
                    i = 0
                    MsgBox "Stok Giri�i �ptal Edildi."
                    Exit Sub
                    End If
                Else
                i = 1
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = K�sakod_Text
                i = i + 1 'i=2
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Spec_Text
                i = i + 1 'i=3
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_BOLINO.Text
                i = i + 1 'i=4
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_ADET.Text
                i = i + 1 'i=5
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_4.Text
                i = i + 1 'i=6
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_SORUMLU.Text
                i = i + 1 'i=7
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = CB_GIRISYAPAN.Text
                i = i + 1 'i=8
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = TB_NOT.Text
                i = i + 1
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = Now
                i = i + 1
                Sheets("KAYITLAR").Cells(EN_KAYITALT, i).Value = "STOK G�R�� �PTAL�"
                i = 0
                MsgBox "Stok Giri�i �ptal Edildi."
                Exit Sub
                End If
            End If
    End If
End Sub

Private Sub CommandButton1_Click()
Unload Me
UserForm1.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me
Application.Visible = True
ActiveWorkbook.Save
ActiveWorkbook.Close
End Sub

Private Sub TB_4_Change() ' Tarih k�s�tlay�c�
            Dim TB_4_Text As String
            TB_4_Text = TB_4.Text
            If Len(TB_4_Text) > 0 Then
                TB_4_TR = Right(TB_4_Text, 1)
                If Not (IsNumeric(TB_4_TR)) Then
                Beep
                    TB_4_Text = Left(TB_4_Text, ((Len(TB_4_Text)) - 1))
                ElseIf (Len(TB_4_Text) = 3 And Mid(TB_4_Text, 3, 1) <> "/") Then
                    TB_4_Text = Left(TB_4_Text, 2) & "/" & Right(TB_4_Text, 1)
                ElseIf (Len(TB_4_Text) = 6 And Mid(TB_4_Text, 6, 1) <> "/") Then
                    TB_4_Text = Left(TB_4_Text, 5) & "/" & Right(TB_4_Text, 1)
                ElseIf (Len(TB_4_Text)) > 10 Then
                Beep
                    TB_4_Text = Left(TB_4_Text, 10)
                End If
            End If
            TB_4.Text = TB_4_Text
End Sub
Private Sub TB_ADET_Change()

            If (Len(TB_ADET.Text)) > 0 Then
            If Not (IsNumeric(Right(TB_ADET.Text, 1))) Then
            Beep
            TB_ADET.Text = Left(TB_ADET.Text, ((Len(TB_ADET.Text)) - 1))
            
            ElseIf Len(TB_ADET.Text) > 5 Then
            Beep
            TB_ADET.Text = Left(TB_ADET.Text, ((Len(TB_ADET.Text)) - 1))
            End If
        End If
End Sub
Private Sub TB_KISAKOD_Change()
            If Len(TB_KISAKOD.Text) > 3 Then
            Beep
            TB_KISAKOD.Text = Left(TB_KISAKOD.Text, ((Len(TB_KISAKOD.Text)) - 1))
            End If
End Sub
Private Sub TB_SPEC_Change()
            If Len(TB_SPEC.Text) > 6 Then
            Beep
            TB_SPEC.Text = Left(TB_SPEC.Text, ((Len(TB_SPEC.Text)) - 1))
            End If
End Sub




Private Sub UserForm_Initialize()
   TB_4.Text = Format(Now(), "DD/MM/YYYY")
    Call finder 'Comboboxlar� doluran kodu �a��r
End Sub
''GER� ALMAYI UNUTMA!!!
