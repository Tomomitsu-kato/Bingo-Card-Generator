VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenu 
   Caption         =   "メニュー"
   ClientHeight    =   5040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3780
   OleObjectBlob   =   "frmMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'メニューの最大最小を初期化
Private Sub UserForm_Initialize()
    txtMin.Value = MIN_DEFAULT
    txtMax.Value = MAX_DEFAULT
End Sub

'シャッフル
Private Sub CmdShuffle_Click()
    Shuffle
End Sub

'リセット
Private Sub CmdReset_Click()
    Reset
End Sub

'印刷
Private Sub CmdpPrint_Click()
    Me.Hide
    PrintIt
    Me.Show
End Sub
