Attribute VB_Name = "Module1"
Public Const SETTING_START_COL = 4       ' �ݒ�V�[�g�̊J�n�s�ԍ�
Public Const SETTING_HEADER_ROW = 2      ' �ݒ�V�[�g�̃^�X�N�ꗗ�̌��o���s�̗�ԍ�
Public Const SETTING_TASK_START_ROW = 3  ' �ݒ�V�[�g�̃^�X�N�ꗗ�̊J�n�s�̗�ԍ�
Public Const SETTING_TASK_START_COL = 4  ' �ݒ�V�[�g�̃^�X�N�ꗗ�̊J�n��̗�ԍ�
Public Const SETTING_TASK_END_COL = 5    ' �ݒ�V�[�g�̃^�X�N�ꗗ�̏I����̗�ԍ�
Public Const SETTING_TASK_PRIOR_COL = 6  ' �ݒ�V�[�g�̃^�X�N�ꗗ�̕��ёւ��D�捀�ڗ�̗�ԍ�


' �^�X�N�ꗗ���\�[�g����
Sub �^�X�N�̃\�[�g()
    Dim headerRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As String
    Dim endCol As String
    Dim initSelection As Range
    Dim priorColNames() As String
    Dim targetCol As String
    Dim i As Long
    
    ' �\�[�g���s�O�̏���
    Application.ScreenUpdating = False                                                    ' �X�N���[���̍X�V���~�߂�
    Set initSelection = Selection                                                         ' ���݂̃Z���I���ʒu��ޔ�����
    headerRow = Worksheets("�ݒ�").Cells(SETTING_START_COL, SETTING_HEADER_ROW).value     ' �ݒ�V�[�g�̌��o���s�ԍ����擾
    startRow = Worksheets("�ݒ�").Cells(SETTING_START_COL, SETTING_TASK_START_ROW).value  ' �ݒ�V�[�g�̊J�n�s�ԍ����擾
    startCol = Worksheets("�ݒ�").Cells(SETTING_START_COL, SETTING_TASK_START_COL).value  ' �ݒ�V�[�g�̊J�n��̎擾���擾
    endCol = Worksheets("�ݒ�").Cells(SETTING_START_COL, SETTING_TASK_END_COL).value      ' �ݒ�V�[�g�̏I����̎擾���擾
    Range(startCol & startRow).Select                                                     ' �J�n�Z���̑I��
    endRow = Selection.End(xlDown).Row                                                    ' �ŏI�s�̎擾
    
    ' �f�[�^�̕��ёւ��D��x��ݒ肷��
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear  ' �\�[�g�̐ݒ��������
    priorColNames = Split(GetPriorColNames(), ",")
    For i = LBound(priorColNames) To UBound(priorColNames)
        targetCol = FindColumn(headerRow, priorColNames(i))
        ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add _
            Key:=Range(targetCol & startRow & ":" & targetCol & endRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
    Next i
    
    ' �f�[�^�̕��ёւ�
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(GetNextColumn(startCol) & startRow & ":" & endCol & endRow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' No.�̔ԍ����I�[�g�t�B���Őݒ肵�Ȃ���
    Range(startCol & startRow & ":" & startCol & startRow + 1).Select
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown)), Type:=xlFillValues
    
    ' �\�[�g���s��̌㏈��
    initSelection.Select               ' �����̃Z���I���ʒu�ɖ߂�
    Application.ScreenUpdating = True  ' �X�N���[���̍X�V���ĊJ����
End Sub


' �\�[�g��D�悷�鍀�ږ��̎擾
Function GetPriorColNames() As String
    Dim startRow As Long
    Dim i As Integer
    Dim value As String
    
    i = SETTING_START_COL
    Do While Worksheets("�ݒ�").Cells(i, SETTING_TASK_PRIOR_COL).value <> ""   ' �ݒ�V�[�g�̕��ёւ��D�捀�ڂ̊J�n�s���擾
        GetPriorColNames = GetPriorColNames & "," & Worksheets("�ݒ�").Cells(i, SETTING_TASK_PRIOR_COL).value
        i = i + 1 ' ���̍s�Ɉړ�
    Loop
    
    ' �擪��","������
    GetPriorColNames = Mid(GetPriorColNames, 2)
End Function


' �w�蕶����������̍s�ԍ����n�_�Ɍ������A���������s�ԍ���Ԃ�
Function FindColumn(ByVal rowNum As Long, ByVal searchStr As String) As String
    Dim i As Long
    Dim lastColumn As Long
    
    ' �ŏI��̎擾
    lastColumn = Cells(rowNum, Columns.Count).End(xlToLeft).Column
    
    ' �w�蕶�����T��
    For i = 1 To lastColumn
        If Cells(rowNum, i).value = searchStr Then
            FindColumn = Split(Cells(rowNum, i).Address, "$")(1)  ' ��ԍ��̎擾
            Exit Function                                         ' �Ώۂ̗񂪌��������ꍇ�͊֐����I��
        End If
    Next i
    
    ' �Ώۂ̗񂪌�����Ȃ������ꍇ�̓G���[�I��
    Err.Raise 9999, , "�w��̕����񂪌�����܂���ł���"
End Function


' ���̗񖼂��擾����
Function GetNextColumn(ByVal columnStr As String) As String
    Dim currentColumun As Range
    Dim nextColumun As Range
    
    Set currentColumn = Range(columnStr & "1")         ' �w�肵�����1�s�ڂ̃Z�����擾
    Set nextColumn = currentColumn.Offset(0, 1)        ' �w�肵����̎��̗�̃Z�������擾
    GetNextColumn = Split(nextColumn.Address, "$")(1)  ' �w�肵�����̗�̗񖼂��擾����
End Function


' �V�����s����ԉ��ɒǉ�����
Sub �s�̒ǉ�()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As String
    Dim endCol As String
    Dim networkdaysCol As String
    Dim initSelection As Range
    
    ' �s�ǉ��O�̏���
    Application.ScreenUpdating = False                                                    ' �X�N���[���̍X�V���~�߂�
    Set initSelection = Selection                                                         ' ���݂̃Z���I���ʒu��ޔ�����
    headerRow = Worksheets("�ݒ�").Cells(SETTING_START_COL, SETTING_HEADER_ROW).value     ' �ݒ�V�[�g�̌��o���s�ԍ����擾
    startRow = Worksheets("�ݒ�").Cells(SETTING_START_COL, SETTING_TASK_START_ROW).value  ' �ݒ�V�[�g�̊J�n�s�ԍ����擾
    startCol = Worksheets("�ݒ�").Cells(SETTING_START_COL, SETTING_TASK_START_COL).value  ' �ݒ�V�[�g�̊J�n��̎擾���擾
    Range(startCol & startRow).Select                                                     ' �J�n�Z���̑I��
    endRow = Selection.End(xlDown).Row                                                    ' �ŏI�s�̎擾
    ActiveSheet.Outline.ShowLevels columnlevels:=2                                        ' �O���[�v�����ꂽ���K���\��������
    
    '�V�����s�̑}��
    Set ws = ActiveSheet                                    ' �A�N�e�B�u�ȃV�[�g���擾
    ws.Rows(endRow + 1).Insert Shift:=xlDown                ' �V�����s��}��
    ws.Rows(endRow).EntireRow.Copy                          ' �O�̍s���R�s�[
    ws.Rows(endRow + 1).PasteSpecial Paste:=xlPasteFormats  ' �O�̍s�̏�����V�����s�ɃR�s�[
    
    ' �uNo.�v�̗�͑O�̍s�Ɂ{1����
    Range(startCol & endRow + 1).value = Range(startCol & endRow).value + 1
    
    ' �u�����v�̗�͑O�̍s�̐������R�s�[����
    networkdaysCol = FindColumn(headerRow, "����")                          ' ������̎擾
    Range(networkdaysCol & endRow).Copy                                     ' �ǉ��O�̍ŏI�s���R�s�[
    Range(networkdaysCol & endRow + 1).PasteSpecial Paste:=xlPasteFormulas  ' �ǉ��s�ɑO�s�̐������R�s�[����
    
    ' �s�ǉ���̌㏈��
    Application.CutCopyMode = False    ' �N���b�v�{�[�h���N���A����
    initSelection.Select               ' �����̃Z���I���ʒu�ɖ߂�
    Application.ScreenUpdating = True  ' �X�N���[���̍X�V���ĊJ����
End Sub
