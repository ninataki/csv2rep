'Attribute VB_Name = "SCO�񍐎���"
'�z��t�H�[�}�b�g
'�薼   ��(�ŐV�R�����g)    �D��x(�m�x)  �X�e�[�^�X  �ۑ肠��    �S����    �S������    KFA g�@��

Sub ALL_SCO�񍐏��쐬()
    Call �ݒ�ύX
    Call �S�̐��`
    Call �����ւ�
    Call �薼�Ə󋵂̐��`
    Call �C�R�[���̏���
    Call �I�[�g�t�B���^
    Call �^�C�g���̐F
End Sub



'�����̂���͈͂�I���A�܂�Ԃ��A���E�������낦�ɂ���A�ڐ�����������A�g��������B�㉺���F3��قǂɂ���B
Sub �ݒ�ύX()
    ActiveWindow.DisplayGridlines = False '�ڐ������\��
End Sub

Sub �S�̐��`()
    With Range("A1").CurrentRegion '�A�N�e�B�u�̈�̒�`
    .Borders.LineStyle = xlContinuous '�r��������
    .HorizontalAlignment = xlLeft '������
    .VerticalAlignment = xlCenter '�㉺��������
    .WrapText = True '�܂�Ԃ�
    .ColumnWidth = 10 '����
    .RowHeight = 30 '�㉺��
    .Interior.ColorIndex = xlNone
    End With
End Sub

'�u�薼�v���O�ɂ���A�ŐV�̃R�����g����
Sub �����ւ�()
    If Range("B1").Value = "�薼" Then '���֍ς݂Ȃ珈�����Ȃ�
        Exit Sub
    End If
    Columns("B").Cut '�ŐV�̃R�����g=��B ��C��
    Columns("C").Insert Shift:=xlToRight
    Columns("C").Cut '�薼=��C ��B��
    Columns("B").Insert Shift:=xlToRight
    Range("C1").Value = "��" 'C1�Z�������l�[��
End Sub

'�ŐV�̃R�����g�Ƒ薼�����ɐL�΂��B�ŐV�̃R�����g�͏㑵��
Sub �薼�Ə󋵂̐��`()
    Range("B1").EntireColumn.ColumnWidth = 40  '��������
    Range("C1").EntireColumn.ColumnWidth = 130 '�󋵂̉�������
    Range("C1").Value = "��" 'C1�Z����ύX
    Range("C1").EntireColumn.VerticalAlignment = xlTop '�󋵂͏㑵���ɂ���
    Range("C1").VerticalAlignment = xlCenter '�^�C�g���������������ɒ���
End Sub

Sub �C�R�[���̏���()
    Cells.Replace What:="=", Replacement:="'"
End Sub

Sub �I�[�g�t�B���^()
  Range("A1").AutoFilter
End Sub

'�^�C�g����Z���΂ɂ���
Sub �^�C�g���̐F()
    Range("A1").AutoFilter Field:=1, Criteria1:="#" '�^�C�g���s��I��
    With Range("A1").CurrentRegion
    .Interior.ThemeColor = xlThemeColorAccent6 'offce�J���[�̈�ԉE(��)
    .Interior.TintAndShade = 0.5 '2�Ԗڂɔ����F
    End With
    ActiveSheet.ShowAllData '�I�[�g�t�B���^�S�\��
End Sub

