Attribute VB_Name = "�t�@�C�����X�g��"
Option Explicit

Public Const glb_sht_main = "Main"
Public Const glb_path = "C4"

'===================================
' �V�[�g�V�K�쐬 ���Ԃō쐬
'===================================
Function add_sht(ByRef �p�X As String) As String
    �p�X = Sheets(glb_sht_main).Range(glb_path).Value
    
    Dim �t�H���_�� As String
    �t�H���_�� = fol_name(�p�X)
    
    Dim sht As Worksheet
    Set sht = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    sht.Name = �t�H���_�� & Format(Now(), "h��mm��ss�b")
    
    add_sht = sht.Name
End Function

'===================================
' �t�H���_���擾
'===================================
Function fol_name(ByVal �p�X As String) As String
    Dim pos As Long
    pos = InStrRev(�p�X, "\")
    fol_name = Mid(�p�X, pos + 1, 30)
End Function

'===================================
' ���C������
'===================================
Sub Create_BookList_From_Folder()
    Dim �p�X As String
    Dim MyPath As String
    Dim cnt As Long
    Dim shtName As String
    
    Application.ScreenUpdating = False

    '�V�[�g�쐬��t�H���_�p�X�擾
    shtName = add_sht(�p�X)
    MyPath = �p�X
    
    '�擪�s
    With ActiveSheet
        .Range("A1").Value = "Path"
        .Range("B1").Value = "FolderName"
        .Range("C1").Value = "FileName"
        .Range("D1").Value = "�X�V����"
    End With
    
    cnt = 2
    
    'FileSystemObject�����
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�ꗗ�����o��
    Call Create_BookList_From_Folder2(fso, MyPath, cnt)
    
    Columns("B:D").AutoFit

    '�t�H���_�����̃V�[�g�쐬
    Dim newSht As Worksheet
    Dim lastRow As Long
    Dim baseSht As Worksheet
    Set baseSht = ActiveSheet
    
    baseSht.Columns(2).Copy
    Set newSht = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    newSht.Range("A1").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    
    lastRow = newSht.Cells(newSht.Rows.Count, 1).End(xlUp).Row
    newSht.Range("A1:A" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    newSht.Name = shtName & "_�t�H���_�ꗗ"

    Application.ScreenUpdating = True
    
    '���b�Z�[�W
    MsgBox "����"
    
End Sub

'===================================
' �ċA�I�Ƀt�@�C�������擾
'===================================
Sub Create_BookList_From_Folder2(fso As Object, MyPath As String, ByRef cnt As Long)
    Dim folder As Object
    Dim fileItem As Object
    Dim subFolder As Object
    
    Set folder = fso.GetFolder(MyPath)
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    '�t�@�C���ꗗ
    For Each fileItem In folder.Files
        ws.Cells(cnt, 1).Value = fileItem.Path
        ws.Cells(cnt, 2).Value = fileItem.ParentFolder.Name
        ws.Cells(cnt, 3).Value = fileItem.Name
        ws.Cells(cnt, 4).Value = fileItem.DateLastModified
        
        ws.Hyperlinks.Add Anchor:=ws.Cells(cnt, 2), Address:=fileItem.ParentFolder.Path
        ws.Hyperlinks.Add Anchor:=ws.Cells(cnt, 3), Address:=fileItem.Path
        
        cnt = cnt + 1
    Next fileItem

    '�T�u�t�H���_�̍ċA����
    For Each subFolder In folder.SubFolders
        Call Create_BookList_From_Folder2(fso, subFolder.Path, cnt)
    Next subFolder
End Sub

