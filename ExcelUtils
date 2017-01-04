''' <sumarry>
''' プロジェクトをエクスポート
''' </sumarry>
''' <remarks>
''' ※ツール - 参照設定から
''' "Microsoft Visual Basic for Applications Extensibility"
''' を追加してください。
''' </remarks>
Private Sub export()
On Error GoTo FINALLY
    
    ' このブックのVBProjectをオブジェクト変数に格納する。
    Dim Obj As VBIDE.VBProject
    Set Obj = ThisWorkbook.VBProject

    ' VBProjectに存在するコンポーネント数を変数に格納する。
    Dim CompCnt As Long
    CompCnt = Obj.VBComponents.Count
    
    ' 出力先
    Dim ExportDir As String
    ExportDir = ThisWorkbook.path & "\src"
    
    ' 出力先フォルダを作成
    If Dir(ExportDir, vbDirectory) <> "" Then
        Call Kill(ExportDir & "\*")
        Call RmDir(ExportDir)
    End If

    ' VBProjectに存在するコンポーネントを一つずつ参照する。
    Dim lp As Long
    For lp = 1 To CompCnt
        ' エクスポートするファイルの拡張子を決定する。
        Dim str As String
        Select Case Obj.VBComponents(lp).Type
            Case 1
                str = ".bas"
            Case 2
                str = ".frm"
            Case Else
                str = ".cls"
        End Select
        
        ' エクスポートを実行する。
        With Obj.VBComponents(lp)
            .export (ExportDir & "\" & .Name & str)
        End With
    Next

    ' オブジェクトを破棄する。
    Set Obj = Nothing
    
    Exit Sub
FINALLY:
    Call MsgBox(Err.Description, vbExclamation, "プロジェクトのエクスポートに失敗しました。")

End Sub

Public Function getWorksheets(Optional ByRef book As Workbook = Nothing) As Dictionary
    Dim dict As New Dictionary
    Dim ws As Worksheet
    If book Is Nothing Then Set book = ThisWorkbook
    For Each ws In book.Worksheets
        Call dict.Add(ws.Name, ws)
    Next
    Set getWorksheets = dict
End Function
