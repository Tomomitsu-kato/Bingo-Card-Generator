VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenu 
   Caption         =   "���j���["
   ClientHeight    =   5040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3780
   OleObjectBlob   =   "frmMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���j���[�̍ő�ŏ���������
Private Sub UserForm_Initialize()
    txtMin.Value = MIN_DEFAULT
    txtMax.Value = MAX_DEFAULT
End Sub

'�V���b�t��
Private Sub CmdShuffle_Click()
    Shuffle
End Sub

'���Z�b�g
Private Sub CmdReset_Click()
    Reset
End Sub

'���
Private Sub CmdpPrint_Click()
    Me.Hide
    PrintIt
    Me.Show
End Sub
