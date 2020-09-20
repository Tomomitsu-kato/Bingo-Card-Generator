Attribute VB_Name = "ModMain"
Option Explicit

Dim nList As New Collection
Dim num As Integer
Dim max As Integer
Dim min As Integer

Public Const MAX_DEFAULT As Integer = 75
Public Const MIN_DEFAULT As Integer = 1

'���j���[��\��
Public Sub Menu()
    frmMenu.Show
End Sub

'�V���b�t��
Public Sub Shuffle()
    Dim flg As Boolean
    
    Set nList = New Collection
    flg = checkMinMax()
    
    If flg Then
        targetTable
    End If
End Sub

'���Z�b�g
Public Sub Reset()
    Set nList = New Collection
    deleteTable
End Sub

'���
Public Sub PrintIt()
    Worksheets("�r���S�J�[�h").PrintPreview
End Sub

'�N�������j���[���J��
Private Sub auto_open()
    Menu
End Sub

'�ő�ŏ��̓��̓`�F�b�N
Private Function checkMinMax() As Boolean
    Dim n0 As Integer
    Dim n1 As Integer
    
    If IsNumeric(frmMenu.txtMin.Value) Then
        n0 = frmMenu.txtMin.Value
    Else
        MsgBox "�ŏ��ɂ͐�������͂��Ă��������B", vbExclamation
        checkMinMax = False
        Exit Function
    End If
    
    If IsNumeric(frmMenu.txtMax.Value) Then
        n1 = frmMenu.txtMax.Value
    Else
        MsgBox "�ő�ɂ͐�������͂��Ă��������B", vbExclamation
        checkMinMax = False
        Exit Function
    End If
    
    If Abs(n1 - n0) < 25 Then
        MsgBox "�ŏ��ƍő�̊Ԃ�25�ȏ�ɂȂ�悤�ɂ��Ă��������B", vbExclamation
        checkMinMax = False
        Exit Function
    End If
    
    If n0 < n1 Then
        min = n0
        max = n1
    Else
        min = n1
        max = n0
    End If
    
    checkMinMax = True
    
End Function

'�e�[�u�����w��
Private Sub targetTable()
    Dim arr_rng() As Variant
    Dim i As Integer
    
    Set nList = New Collection
    arr_rng = Array("C4", "J4", "C12", "J12", "C20", "J20")

    For i = 0 To 5
        writeTable (arr_rng(i))
        Set nList = New Collection
    Next i
End Sub

'�e�[�u���ɏ���
Private Sub writeTable(str As String)
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 4
        For j = 0 To 4
            If i = 2 And j = 2 Then
                'FREE�̂��ߏ����Ȃ�
            Else
                getRnd
                nList.Add (num)
                Range(str).Offset(i, j).Value = num
            End If
        Next j
    Next i
End Sub

'�e�[�u��������
Private Sub deleteTable()
    Dim arr_rng() As Variant
    Dim i As Integer
    
    arr_rng = Array("C4", "J4", "C12", "J12", "C20", "J20")
    
    For i = 0 To 5
        delTable (arr_rng(i))
    Next i
End Sub

'�e�[�u��������
Private Sub delTable(str As String)
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 4
        For j = 0 To 4
            If i = 2 And j = 2 Then
                'FREE�̂��ߏ����Ȃ�
            Else
                Range(str).Offset(i, j).Value = ""
            End If
        Next j
    Next i
End Sub
'nList���Əd�����Ȃ������_�����l
Private Sub getRnd()
    Randomize
    Dim n As Integer
    Dim l As Variant
    Dim b As Boolean
    
    n = Int((max - min + 1) * Rnd + min)
    b = False
    For Each l In nList
        If (n = l) Then
            b = True
        End If
    Next l
        
    If (b) Then
        getRnd
    Else
        num = n
    End If
End Sub
