VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerBookSystem 
   Caption         =   "�w���t�Τ���"
   ClientHeight    =   7060
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   8820.001
   OleObjectBlob   =   "CustomerBookSystem.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "CustomerBookSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub CommandButton1_Click()

Dim rowCnt As Integer
Dim dtrange As Range
Set dtrange = Sheets("�w���t�Τ���").UsedRange '�]�w�w�ϥνd��'
rowCnt = dtrange.Rows.Count                    '�]�w�w�ϥΦC���'
Sheets("�w���t�Τ���").Select
'�ťտ�J
  If cusname.Text = " " Or cusphone.Text = " " Or bookmonth.Text = " " Or bookdate.Text = " " Or booktime.Text = " " Then
  MsgBox ("�Х��T��g���"), vbInformation
  Exit Sub
  End If

Cells(rowCnt + 1, "B").Value = cusname.Text                                             '��J�U�ȩm�W'
Cells(rowCnt + 1, "C").Value = cusphone.Text                                            '��J�U�ȹq��'
Cells(rowCnt + 1, "D").Value = bookmonth.Text                                           '��J�w�����'
Cells(rowCnt + 1, "E").Value = bookdate.Text                                            '��J�w�����'
Cells(rowCnt + 1, "F").Value = booktime.Text                                            '��J�w���ɶ�'
Cells(rowCnt + 1, "L").Value = memo.Text                                                '��J�Ƶ�'
datasets

'2021/7/17�A�s�W�s���ǦC'
Dim rIdxx As Integer
Dim i As Integer
Dim rowCntxxx As Integer
Dim dtRangexxx As Range
Set dtRangexxx = Sheets("�w���t�Τ���").UsedRange
rowCntxxx = dtRangexxx.Rows.Count
Sheets("�w���t�Τ���").Select
If Cells(rowCntxxx, "B").Value <> "" Then                                                   '�P�_�A�p�GB��쪺�Ȥ�����Ŧr��'

   For i = 2 To rowCntxxx                                                                   '�]�wi���ȵ���2��̫�@�C'

    Cells(i, "A").Value = i - 1                                                             'A��쪺�ȵ���{���x�s���m-1'
    
   Next
 
End If

'�P�w�O�_���|��
Dim e As Integer
Dim f As Integer
Dim rowCntt As Integer
Dim dtrangee As Range
Set dtrangee = Sheets("�w���t�Τ���").UsedRange '�]�w�w�ϥνd��'
rowCntt = dtrangee.Rows.Count                    '�]�w�w�ϥΦC���'
Sheets("�w���t�Τ���").Select
For e = rowCntt To rowCntt                                           ' �T�w�O�_���|��
    For f = 2 To 100
        If Sheets("�w���t�Τ���").Cells(e, "C").Value = Sheets("�|���򥻸��").Cells(f, "C").Value Then
        Cells(e, "M").Value = "Y"
        MsgBox ("�|��")
        Exit Sub
        Else
        e = e
        Cells(e, "M").Value = "N"
        End If
    Next
Next



cusname.Text = " "     '�e�X��ƫ�A�t�Τ������^��l��
cusphone.Text = " "
bookmonth.Text = " "
bookdate.Text = " "
booktime.Text = " "
memo.Text = " "

End Sub


Private Sub UserForm_Initialize()

'2021/7/18�A�s�W�U�Կﶵ'

bookmonth.AddItem "1��"  '��ܹw�����'
bookmonth.AddItem "2��"
bookmonth.AddItem "3��"
bookmonth.AddItem "4��"
bookmonth.AddItem "5��"
bookmonth.AddItem "6��"
bookmonth.AddItem "7��"
bookmonth.AddItem "8��"
bookmonth.AddItem "9��"
bookmonth.AddItem "10��"
bookmonth.AddItem "11��"
bookmonth.AddItem "12��"

bookdate.AddItem "1"  '��ܹw�����'
bookdate.AddItem "2"
bookdate.AddItem "3"
bookdate.AddItem "4"
bookdate.AddItem "5"
bookdate.AddItem "6"
bookdate.AddItem "7"
bookdate.AddItem "8"
bookdate.AddItem "9"
bookdate.AddItem "10"
bookdate.AddItem "11"
bookdate.AddItem "12"
bookdate.AddItem "13"
bookdate.AddItem "14"
bookdate.AddItem "15"
bookdate.AddItem "16"
bookdate.AddItem "17"
bookdate.AddItem "18"
bookdate.AddItem "19"
bookdate.AddItem "20"
bookdate.AddItem "21"
bookdate.AddItem "22"
bookdate.AddItem "23"
bookdate.AddItem "24"
bookdate.AddItem "25"
bookdate.AddItem "26"
bookdate.AddItem "27"
bookdate.AddItem "28"
bookdate.AddItem "29"
bookdate.AddItem "30"
bookdate.AddItem "31"

booktime.AddItem "11:00"  '��ܹw���ɶ�'
booktime.AddItem "12:00"
booktime.AddItem "13:00"
booktime.AddItem "14:00"
booktime.AddItem "15:00"
booktime.AddItem "16:00"
booktime.AddItem "17:00"
booktime.AddItem "18:00"
booktime.AddItem "19:00"



End Sub

'2021/7/17�A���s�W�����ȡA�̭����Ȯ���'

Sub CommandButton2_Click()
DeleteAllforbook.Show
End Sub

Sub CommandButton3_Click()
'2021/7/18�A���X����'
Me.Hide
Unload Me
CustomerUse.Show
End Sub
