VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPeriod 
   Caption         =   "�N������"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3570
   OleObjectBlob   =   "formPeriod.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// �\��ۑ�����ۂ̃t�@�C�����Ɏg�p����N�����w��
Option Explicit

'// �u���s�v�{�^���N���b�N��
Private Sub cmdEnter_Click()

    Call saveChart(cmbYear.Value, cmbMonth.Value)
    Unload Me
    
End Sub

'// �t�H�[���N����
Private Sub UserForm_Initialize()
    
    Dim i As Long

    '// �R���{�{�b�N�X�̒l�ݒ�(���݂���5�N�܂��܂őI�����Ƃ��Ēǉ�)
    With cmbYear
        For i = Year(Now) - 5 To Year(Now)
            .AddItem i & "�N"
        Next
        .Value = Year(Now) & "�N"
    End With
    
    With cmbMonth
        For i = 1 To 12
            .AddItem i & "��"
        Next
        .Value = Month(Now) & "��"
    End With
End Sub

'// �u����v�{�^���N���b�N��
Private Sub cmdClose_Click()

    Unload Me

End Sub
