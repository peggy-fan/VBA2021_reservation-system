VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProductBookSystem 
   Caption         =   "���~�q�ʨt��"
   ClientHeight    =   6980
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8760.001
   OleObjectBlob   =   "ProductBookSystem.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "ProductBookSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1s_Click()

Dim rowCntxx As Integer                                                                   '�ŧi�ܼ�'

Dim dtRangexx As Range                                                                    '�ŧi�d��'

'�ťտ�J
  If cusnamee.Text = " " Or cusphonee.Text = " " Or productbookk.Text = " " Or bookknum.Text = " " Then
  MsgBox ("�Х��T��g���"), vbInformation
  Exit Sub
  End If

Set dtRangexx = Sheets("���~�q�ʨt�Τ���").UsedRange                                      '�]�w�w�ϥνd��'
Sheets("���~�q�ʨt�Τ���").Select
rowCntxx = dtRangexx.Rows.Count                                                           '�]�w�w�ϥΦC���'

Cells(rowCntxx + 1, "B").Value = cusnamee.Text                                             '��J�U�ȩm�W'
Cells(rowCntxx + 1, "C").Value = cusphonee.Text                                            '��J�q�ܸ��X'
Cells(rowCntxx + 1, "D").Value = productbookk.Text                                         '��J�q�ʲ��~'
Cells(rowCntxx + 1, "E").Value = bookknum.Text                                             '��J�q�ʼƶq'



                           
'2021/7/17�A�s�W�s���ǦC'
Dim rIdxx As Integer
Dim i As Integer
Dim rowCntxxx As Integer
Dim dtRangexxx As Range
Set dtRangexxx = Sheets("���~�q�ʨt�Τ���").UsedRange
rowCntxxx = dtRangexxx.Rows.Count
Sheets("���~�q�ʨt�Τ���").Select
If Cells(rowCntxxx, "B").Value <> "" Then                                                   '�P�_�A�p�GB��쪺�Ȥ�����Ŧr��'

   For i = 2 To rowCntxxx                                                                   '�]�wi���ȵ���2��̫�@�C'

    Cells(i, "A").Value = i - 1                                                             'A��쪺�ȵ���{���x�s���m-1'
    
   Next
 
End If

cusnamee.Text = " "        '�e�X��ƫ�A�t�Τ������^��l��
cusphonee.Text = " "
productbookk.Text = " "
bookknum.Text = " "

End Sub


Private Sub UserForm_Initialize()

'2021/7/18�A�s�W�U�Կﶵ'

productbookk.AddItem "�饻FIOLE�~�v��"  '��ܭq�ʲ��~'
productbookk.AddItem "�饻FIOLE��v��"
productbookk.AddItem "�饻FIOLE�V��"

bookknum.AddItem "1"  '��ܭq�ʼƶq'
bookknum.AddItem "2"
bookknum.AddItem "3"
bookknum.AddItem "4"
bookknum.AddItem "5"
bookknum.AddItem "6"
bookknum.AddItem "7"
bookknum.AddItem "8"
bookknum.AddItem "9"
bookknum.AddItem "10"

End Sub

'2021/7/17�A���s�W�����ȡA�̭����Ȯ���'

Sub CommandButton2s_Click()
DeleteAllforProduct.Show
End Sub

Sub CommandButton3s_Click()
'2021/7/18�A���X����'
Me.Hide
Unload Me
CustomerUse.Show
End Sub

