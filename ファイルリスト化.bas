Attribute VB_Name = "ファイルリスト化"
Option Explicit

Public Const glb_sht_main = "Main"
Public Const glb_path = "C4"

'===================================
' シート新規作成 時間で作成
'===================================
Function add_sht(ByRef パス As String) As String
    パス = Sheets(glb_sht_main).Range(glb_path).Value
    
    Dim フォルダ名 As String
    フォルダ名 = fol_name(パス)
    
    Dim sht As Worksheet
    Set sht = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    sht.Name = フォルダ名 & Format(Now(), "h時mm分ss秒")
    
    add_sht = sht.Name
End Function

'===================================
' フォルダ名取得
'===================================
Function fol_name(ByVal パス As String) As String
    Dim pos As Long
    pos = InStrRev(パス, "\")
    fol_name = Mid(パス, pos + 1, 30)
End Function

'===================================
' メイン処理
'===================================
Sub Create_BookList_From_Folder()
    Dim パス As String
    Dim MyPath As String
    Dim cnt As Long
    Dim shtName As String
    
    Application.ScreenUpdating = False

    'シート作成､フォルダパス取得
    shtName = add_sht(パス)
    MyPath = パス
    
    '先頭行
    With ActiveSheet
        .Range("A1").Value = "Path"
        .Range("B1").Value = "FolderName"
        .Range("C1").Value = "FileName"
        .Range("D1").Value = "更新日時"
    End With
    
    cnt = 2
    
    'FileSystemObjectを作る
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '一覧書き出す
    Call Create_BookList_From_Folder2(fso, MyPath, cnt)
    
    Columns("B:D").AutoFit

    'フォルダだけのシート作成
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
    newSht.Name = shtName & "_フォルダ一覧"

    Application.ScreenUpdating = True
    
    'メッセージ
    MsgBox "完成"
    
End Sub

'===================================
' 再帰的にファイル情報を取得
'===================================
Sub Create_BookList_From_Folder2(fso As Object, MyPath As String, ByRef cnt As Long)
    Dim folder As Object
    Dim fileItem As Object
    Dim subFolder As Object
    
    Set folder = fso.GetFolder(MyPath)
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'ファイル一覧
    For Each fileItem In folder.Files
        ws.Cells(cnt, 1).Value = fileItem.Path
        ws.Cells(cnt, 2).Value = fileItem.ParentFolder.Name
        ws.Cells(cnt, 3).Value = fileItem.Name
        ws.Cells(cnt, 4).Value = fileItem.DateLastModified
        
        ws.Hyperlinks.Add Anchor:=ws.Cells(cnt, 2), Address:=fileItem.ParentFolder.Path
        ws.Hyperlinks.Add Anchor:=ws.Cells(cnt, 3), Address:=fileItem.Path
        
        cnt = cnt + 1
    Next fileItem

    'サブフォルダの再帰処理
    For Each subFolder In folder.SubFolders
        Call Create_BookList_From_Folder2(fso, subFolder.Path, cnt)
    Next subFolder
End Sub

