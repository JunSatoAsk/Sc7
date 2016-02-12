Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core

Public Class FrmMain

    '印刷項目の構造体 文字列
    Private Structure stShift
        Public dtStime As DateTime  '開始時刻
        Public dtStimeDbl As Double '開始時刻
        Public dtEtime As DateTime  '終了時刻
        Public dtEtimeDbl As Double '終了時刻
        Public strName As String    '出力文字列
        Public bcol As Long         '背景色
        Public iNum As Integer      '出力行
        Public bSyain As Boolean    '社員 
    End Structure

    'フォームロードイベント
    Private Sub FrmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        ddlYear.SelectedIndex = 0
        Label1.Text = "aaa"
        Label1.Text = "bbb"
        Label1.Text = "cccadasdasdasa"

    End Sub

    '終了ボタン　
    Private Sub btnEnd_Click(sender As System.Object, e As System.EventArgs) Handles btnEnd.Click

        Me.Close()

    End Sub

    '作成ボタン
    Private Sub btn_Exec_Click(sender As System.Object, e As System.EventArgs) Handles btn_Exec.Click

        'SaveFileDialogクラスのインスタンスを作成
        Dim ofd As New OpenFileDialog()

        Try
            'はじめのファイル名を指定する
            ofd.FileName = "シフト（給与・実働)ベース.xlsx"
            ofd.Filter = _
                "EXCELファイル(*.xlsx;*.xls)|*.xlsx;*.xls|すべてのファイル(*.*)|*.*"
            ofd.FilterIndex = 1
            ofd.Title = "開くファイルを選択してください"
            ofd.RestoreDirectory = True
            ofd.CheckPathExists = True

            'ダイアログを表示する
            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            'Excelオブジェクト
            Dim oExcel As New Excel.Application
            'WorkBookオブジェクト
            Dim oBook As Excel.Workbook
            Dim oSheet As Excel.Worksheet
            Dim oSheet2 As Excel.Worksheet

            'Excelを表示にする
            oExcel.Application.Visible = True
            oExcel.DisplayAlerts = False

            oBook = oExcel.Workbooks.Open(ofd.FileName)

            If wssChk(oBook, "週間スケジュール") = False Then
                Me.TopMost = True
                MessageBox.Show("週間スケジュールのシートがありません。", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.TopMost = False
                oBook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                oBook = Nothing
                oExcel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                oExcel = Nothing
                Return
            End If

            If wssChk(oBook, ddlYear.SelectedItem.ToString + "_週間") = True Then
                Me.TopMost = True
                MessageBox.Show("作成するシートが存在します。削除してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.TopMost = False
                oBook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                oBook = Nothing
                oExcel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                oExcel = Nothing
                Return
            End If

            oSheet = DirectCast(oBook.Sheets(ddlYear.SelectedItem.ToString), Excel.Worksheet)

            oSheet2 = DirectCast(oBook.Sheets("週間スケジュール"), Excel.Worksheet)

            oSheet2.Select()
            '現在アクティブなシートを新規コピー
            oSheet2.Copy(After:=oSheet2)
            'Sheet1のシート名を"新規ワークシート"という名前に変更
            oBook.ActiveSheet.Name = ddlYear.SelectedItem.ToString + "_週間"

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet2)
            oSheet2 = Nothing

            oSheet2 = DirectCast(oBook.Sheets(ddlYear.SelectedItem.ToString + "_週間"), Excel.Worksheet)

            Dim iOfstC As Integer = 3
            Dim iOfstR As Integer = 5
            Dim iOfs As Integer = daysofweekofs(oSheet.Cells(iOfstR - 1, iOfstC).text)
            For iCol As Integer = 0 To 30
                'MessageBox.Show(oSheet.Cells(iOfstR, iOfstC + iCol).text)
                If oSheet.Cells(iOfstR - 2, iOfstC + iCol).text = Nothing Then
                    '最終日まで来たので終了する。
                    Exit For
                End If

                Dim aryShift(My.Settings.RowNum - 1) As stShift
                Dim arySort(My.Settings.RowNum - 1) As Double
                '一日分のデータを取得
                For iRow As Integer = 0 To UBound(aryShift)
                    If oSheet.Cells(iOfstR + (iRow * 5), iOfstC + iCol).text = Nothing Then
                        Continue For
                    End If

                    If oSheet.Cells(iOfstR + (iRow * 5), iOfstC - 1).text = "人数" Then
                        Exit For
                    End If
                    arySort(iRow) = oSheet.Cells(iOfstR + (iRow * 5), iOfstC + iCol).value
                    aryShift(iRow).dtStime = Date.FromOADate(oSheet.Cells(iOfstR + (iRow * 5), iOfstC + iCol).value)
                    aryShift(iRow).dtStimeDbl = oSheet.Cells(iOfstR + (iRow * 5), iOfstC + iCol).value
                    aryShift(iRow).dtEtime = Date.FromOADate(oSheet.Cells(iOfstR + (iRow * 5) + 1, iOfstC + iCol).value)
                    aryShift(iRow).dtEtimeDbl = oSheet.Cells(iOfstR + (iRow * 5) + 1, iOfstC + iCol).value

                    'AM6:00前なら1日追加する（その日の深夜だからね）
                    If arySort(iRow) < 0.25 Then
                        arySort(iRow) = arySort(iRow) + 1.0
                        aryShift(iRow).dtStimeDbl = aryShift(iRow).dtStimeDbl + 1.0
                        aryShift(iRow).dtEtimeDbl = aryShift(iRow).dtEtimeDbl + 1.0
                    End If

                    aryShift(iRow).strName = oSheet.Cells(iOfstR + (iRow * 5), iOfstC - 1).text
                    aryShift(iRow).bcol = oSheet.Cells(iOfstR + (iRow * 5), iOfstC - 1).Interior.Color
                    'No9以下は社員とする
                    If CInt(oSheet.Cells(iOfstR + (iRow * 5), iOfstC - 2).text) < 10 Then
                        aryShift(iRow).bSyain = True
                    Else
                        aryShift(iRow).bSyain = False
                    End If
                Next

                '時間の早い順にソート
                Array.Sort(arySort, aryShift)

                '週間スケジュールのどのラインに出力するか決定する
                Dim i As Integer = 0
                Dim aryEtimeDbl(10) As Double
                Dim bstart As Boolean = False

                For iRow As Integer = 0 To UBound(aryShift)
                    If aryShift(iRow).bSyain = True Then
                        Continue For
                    End If
                    If aryShift(iRow).strName = Nothing Then
                        aryShift(iRow).iNum = -1
                        Continue For
                    End If
                    If bstart = True Then
                        If arysearch(i, aryEtimeDbl, aryShift(iRow).dtStimeDbl) = False Then
                            Me.TopMost = True
                            MessageBox.Show(oSheet.Cells(iOfstR - 2, iOfstC + iCol).text + ":" + aryShift(iRow).strName + "さんの線は手動で移動してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Me.TopMost = False
                            aryShift(iRow).iNum = 0
                            Continue For
                        End If
                    End If
                    bstart = True
                    aryShift(iRow).iNum = i
                    aryEtimeDbl(i) = aryShift(iRow).dtEtimeDbl
                Next

                i = 8
                bstart = False
                For iRow As Integer = 0 To UBound(aryShift)
                    If aryShift(iRow).bSyain = False Then
                        Continue For
                    End If
                    If aryShift(iRow).strName = Nothing Then
                        aryShift(iRow).iNum = -1
                        Continue For
                    End If
                    If bstart = True Then
                        If arysearch2(i, aryEtimeDbl, aryShift(iRow).dtStimeDbl) = False Then
                            Me.TopMost = True
                            MessageBox.Show(oSheet.Cells(iOfstR - 2, iOfstC + iCol).text + ":" + aryShift(iRow).strName + "さんの線は手動で移動してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Me.TopMost = False
                            aryShift(iRow).iNum = 8
                            Continue For
                        End If
                    End If
                    bstart = True
                    aryShift(iRow).iNum = i
                    aryEtimeDbl(i) = aryShift(iRow).dtEtimeDbl
                Next



                oSheet2.Select()

                '日付・曜日を出力
                'oSheet2.Cells(2 + ((iCol + iOfs) * 10), 2).Font.Color = oSheet.Cells(iOfstR - 2, iOfstC + iCol).Font.Color
                'oSheet2.Cells(2 + ((iCol + iOfs) * 10), 3).Font.Color = oSheet.Cells(iOfstR - 1, iOfstC + iCol).Font.Color
                oSheet2.Select()
                oSheet2.Cells(2 + ((iCol + iOfs) * 11), 2) = oSheet.Cells(iOfstR - 2, iOfstC + iCol).text
                oSheet2.Cells(2 + ((iCol + iOfs) * 11), 3) = oSheet.Cells(iOfstR - 1, iOfstC + iCol).text

                Dim xlShapes As Excel.Shapes
                Dim xlShape As Excel.Shape
                xlShapes = oSheet2.Shapes

                '週間スケジュールに描画
                For iRow As Integer = 0 To UBound(aryShift)
                    If aryShift(iRow).strName = Nothing Then
                        Continue For
                    End If
                    Dim idx As Integer = aryShift(iRow).iNum + 2 + ((iCol + iOfs) * 11)
                    Dim Rng As Excel.Range
                    Rng = oSheet2.Range("G" + idx.ToString + ":BB" + idx.ToString)

                    'ブロック矢印を描画
                    xlShape = xlShapes.AddShape(MsoAutoShapeType.msoShapeChevron, Rng.Left + (aryShift(iRow).dtStimeDbl - 0.25) * Rng.Width, Rng.Top, (aryShift(iRow).dtEtimeDbl - aryShift(iRow).dtStimeDbl) * Rng.Width, Rng.Height)

                    'msoShapeRoundedRectangle   msoShapeLeftRightArrow msoShapeChevron

                    'オートシェイプ(ブロック矢印)の背景色と前景色を設定する
                    Dim xlFillFormat As Excel.FillFormat
                    xlFillFormat = xlShape.Fill
                    '  xlFillFormat.ForeColor.RGB = RGB(220, 230, 242)
                    xlFillFormat.ForeColor.RGB = Color_to_RGB(aryShift(iRow).bcol)
                    xlFillFormat.Transparency = 0.25
                    MRComObject(xlFillFormat)

                    Dim xlLineFormat As Excel.LineFormat
                    xlLineFormat = xlShape.Line
                    xlLineFormat.Weight = 0.25
                    xlLineFormat.ForeColor.RGB = RGB(150, 150, 150)
                    xlLineFormat.ForeColor.RGB = RGB(50, 50, 50)
                    MRComObject(xlLineFormat)


                    ''赤色・太さ1.5ポイントの矢印線
                    'Dim BX As Single, BY As Single, EX As Single, EY As Single
                    'BX = Rng.Left + (aryShift(iRow).dtStimeDbl - 0.25) * Rng.Width
                    'BY = Rng.Top + 5
                    'EX = BX + (aryShift(iRow).dtEtimeDbl - aryShift(iRow).dtStimeDbl) * Rng.Width
                    'EY = Rng.Top + 5

                    'With xlShapes.AddLine(BX, BY, EX, EY).Line
                    '    .ForeColor.RGB = Color_to_RGB(aryShift(iRow).bcol)
                    '    .Weight = 1.5
                    '    .EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle
                    'End With



                    xlShape.Select()
                    xlShape.TextFrame.Characters.Text = aryShift(iRow).strName  'オートシェイプに文字列を追加
                    'xlShape.TextFrame.HorizontalOverflow = 0                    '文字列をはみ出して描画
                    'xlShape.TextFrame.VerticalOverflow = 0
                    xlShape.TextFrame.MarginTop = 0
                    xlShape.TextFrame.MarginBottom = 0


                    With xlShape.TextFrame2
                        .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                        .TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter
                        With .TextRange.Font.Fill()
                            .ForeColor.RGB = RGB(0, 0, 0)
                            .Transparency = 0
                            .Solid()
                        End With
                    End With

                    MRComObject(xlShape)

                Next
                MRComObject(xlShapes)

            Next

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet2)
            oSheet2 = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet)
            oSheet = Nothing
            'ファイルを保存する
            Dim FilePath As String = ofd.FileName
            oBook.SaveAs(FilePath, Excel.XlFileFormat.xlOpenXMLWorkbook)
            oBook.Close(False)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
            oBook = Nothing


            oExcel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
            oExcel = Nothing

        Catch ex As Exception
            Me.TopMost = True
            MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK)
            Me.TopMost = False

        Finally

        End Try

    End Sub

    'EXCEL シートの存在チェック
    Private Function wssChk(ByVal oBook As Excel.Workbook, ByVal strName As String) As Boolean

        Dim ws As Excel.Worksheet
        Dim flag As Boolean
        For Each ws In oBook.Worksheets
            If ws.Name = strName Then flag = True
        Next ws
        Return flag

    End Function

    Private Function arysearch(ByRef i As Integer, ByVal aryEtimeDbl() As Double, ByVal dtStimeDbl As Double) As Boolean

        i = i + 1
        If i > 7 Then i = 0

        For iCnt = 1 To 8

            If dtStimeDbl < aryEtimeDbl(i) Then
                i = i + 1
                If i > 7 Then i = 0
            Else
                Return True
            End If

        Next

        Return False

    End Function

    Private Function arysearch2(ByRef i As Integer, ByVal aryEtimeDbl() As Double, ByVal dtStimeDbl As Double) As Boolean

        i = i + 1
        If i > 10 Then i = 8

        For iCnt = 9 To 11

            If dtStimeDbl < aryEtimeDbl(i) Then
                i = i + 1
                If i > 10 Then i = 8
            Else
                Return True
            End If

        Next

        Return False

    End Function

    '曜日により出力位置を変えるためのオフセット値を取得
    Private Function daysofweekofs(ByVal strday As String) As Integer

        Dim iRet As Integer = 0
        Select Case strday
            Case "日"
                iRet = 0
            Case "月"
                iRet = 1
            Case "火"
                iRet = 2
            Case "水"
                iRet = 3
            Case "木"
                iRet = 4
            Case "金"
                iRet = 5
            Case "土"
                iRet = 6
        End Select
        Return iRet

    End Function

    'VB2005/VB2008/VB2010 用
    ''' <summary>
    ''' COMオブジェクトの参照カウントをデクリメントします。
    ''' </summary>
    ''' <typeparam name="T">(省略可能)</typeparam>
    ''' <param name="objCom">
    ''' COM オブジェクト持った変数を指定します。
    ''' このメソッドの呼出し後、この引数の内容は Nothing となります。
    ''' </param>
    ''' <param name="force">
    ''' すべての参照を強制解放する場合はTrue、現在の参照のみを減ずる場合はFalse。
    ''' </param>
    Public Shared Sub MRComObject(Of T As Class)(ByRef objCom As T,
                       Optional ByVal force As Boolean = False)
        If objCom Is Nothing Then
            Return
        End If
        Try
            If System.Runtime.InteropServices.Marshal.IsComObject(objCom) Then
                If force Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objCom)
                Else
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
                End If
            End If
        Finally
            objCom = Nothing
        End Try
    End Sub


    'カラープロパティをＲＧＢに変換
    Private Function Color_to_RGB(ByVal lngColor As Long) As Integer

        Dim iRed As Integer = lngColor Mod 256
        Dim iGreen As Integer = Int(lngColor / 256) Mod 256
        Dim iBlue As Integer = Int(lngColor / 256 / 256)

        Return RGB(iRed, iGreen, iBlue)

    End Function



End Class
