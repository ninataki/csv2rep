'Attribute VB_Name = "�񍐎���"
'�z��t�H�[�}�b�g
'�薼    ��    �X�e�[�^�X  �ۑ肠��    �S����  ����    �J�n��  �X�V��  �D��x  �e�`�P�b�g

Sub ALL_�s���敔�񍐏��쐬()
    Call �ݒ�ύX
    Call �S�̐��`
    Call �����ւ�
    Call �薼�Ə󋵂̐��`
    Call �e�`�P�b�g����
    Call �^�C�g���̐F
    Call �X�e�[�^�X����
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
    .RowHeight = 80 '�㉺��
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


'�t�B���^������A�e�`�P�b�g(�u�e�`�P�b�g�v�񂪋�)��I������B
'�^�C�g���Ɛe�`�P�b�g���㉺��1�s�ɂ���A�܂�Ԃ��Ȃ��A���΂ɂ���A�e�`�P�b�g�̃^�C�g���ȊO���폜�A�u�X�e�[�^�X�v���󔒂ɂ���
'�F�� <https://kosapi.com/post-3405/>
Sub �e�`�P�b�g����()

    '�e�`�P�b�g���t�B���^���Đ��`
    Range("A1").AutoFilter Field:=11, Criteria1:="" '�e�`�P�b�g���󔒍s
    With Range("A1").CurrentRegion
    .Interior.ThemeColor = xlThemeColorAccent6 'offce�J���[�̈�ԉE(��)
    .Interior.TintAndShade = 0.8 '1�Ԗڂɔ����F
    .RowHeight = 30 '�㉺��������������
    .WrapText = False '�܂�Ԃ��Ȃ�
    End With
    
    '�󋵂̕s�v�ȉӏ����폜
    Call �^�C�g�������I��(3)  '�R��ځ���
    Selection.ClearContents '�폜

    '�X�e�[�^�X���󔒂ɂ���
    Call �^�C�g�������I��(4) '�S��ځ��X�e�[�^�X
    Selection.Value = "-"

    ActiveSheet.ShowAllData '�I�[�g�t�B���^�S�\��

End Sub

'�^�C�g����Z���΂ɂ���
Sub �^�C�g���̐F()
    
    Range("A1").AutoFilter Field:=11, Criteria1:="�e*" '�^�C�g���s��I��
    With Range("A1").CurrentRegion
    .Interior.ThemeColor = xlThemeColorAccent6 'offce�J���[�̈�ԉE(��)
    .Interior.TintAndShade = 0.5 '2�Ԗڂɔ����F
    End With

    ActiveSheet.ShowAllData '�I�[�g�t�B���^�S�\��
    
End Sub


'�X�e�[�^�X�ɐF��t����uDoing���΁A�ۑ肠��͉��F�A�v�uReview�����F�v�uToD�����΁v�uBacklog���D�F�v
'�X�e�[�^�X�́uBacklog�v���\���ɂ���
' ColorIndex  <https://tripbowl.com/excel-vba/label-color-change/#color>
Sub �X�e�[�^�X����()
    
    '�X�e�[�^�X�̐F�t��
    Range("A1").AutoFilter Field:=4, Criteria1:="Review"
    Call �^�C�g�������I��(4)
    Selection.Interior.ColorIndex = 33 '��

    Range("A1").AutoFilter Field:=4, Criteria1:="Backlog"
    Call �^�C�g�������I��(4)
    Selection.Interior.ColorIndex = 15 '�O���[

    Range("A1").AutoFilter Field:=4, Criteria1:="ToDo"
    Call �^�C�g�������I��(4)
    Selection.Interior.ColorIndex = 43 '������

    Range("A1").AutoFilter Field:=4, Criteria1:="Doing"
    Call �^�C�g�������I��(4)
    Selection.Interior.ColorIndex = 10 '��

    Range("A1").AutoFilter Field:=5, Criteria1:="�͂�" '�X�e�[�^�X Doing & �ۑ肠��
    Call �^�C�g�������I��(4)
    Selection.Interior.ColorIndex = 6 '���F
    
    '�I�[�g�t�B���^�ݒ肵Backlog�����O���ĕ\��
    Range("A1").AutoFilter '�t�B���^����
    Range("A1").AutoFilter Field:=4, Criteria1:="<>Backlog"
    
End Sub

Function �^�C�g�������I��(ByVal num As Integer)
    Range("A1").CurrentRegion.Columns(num).Select '�X�e�[�^�X��D��I��
    Selection.Offset(1, 0).Select '�I��̈������ɂ��炷
    Selection.Resize(Selection.Rows.Count - 1).Select '�I��̈������炷
End Function

