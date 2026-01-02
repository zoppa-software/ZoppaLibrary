Option Explicit On
Option Strict On

Imports System.Drawing
Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 文字列解析ノード。
    ''' </summary>
    NotInheritable Class TerminalNode
        Inherits AnalysisNode

        ''' <summary>リテラル文字列。</summary>
        Private ReadOnly _literal As String

        ''' <summary>シフトテーブル。</summary>
        Private ReadOnly _shiftTable As SortedDictionary(Of Char, Integer)

        ''' <summary>評価範囲。</summary>
        Public Overrides ReadOnly Property Range As ExpressionRange

        ''' <summary>
        ''' 再試行可能かを取得する。
        ''' </summary>
        Public Overrides ReadOnly Property IsRetry As Boolean
            Get
                Return False
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="range">評価範囲。
        Public Sub New(id As Integer, range As ExpressionRange)
            MyBase.New(id)

            ' 文字列を取得
            Me._literal = UnescapedString(range.SubRanges(0).ToString())

            ' シフトテーブルを作成
            Me._shiftTable = New SortedDictionary(Of Char, Integer)()
            For i As Integer = 0 To Me._literal.Length - 1
                Dim c = Me._literal.Chars(i)
                Dim nx = Me._literal.Length - 1 - i
                If Me._shiftTable.ContainsKey(c) Then
                    Me._shiftTable(c) = nx
                Else
                    Me._shiftTable.Add(c, nx)
                End If
            Next
            For Each c As Char In Me._literal
                If Not Me._shiftTable.ContainsKey(c) Then
                    Me._shiftTable(c) = Me._literal.Length
                End If
            Next

            Me.Range = range
        End Sub

        ''' <summary>
        ''' エスケープシーケンスを変換する。
        ''' </summary>
        ''' <param name="str">変換対象の文字列。</param>
        ''' <returns>変換後の文字列。</returns>
        Private Shared Function UnescapedString(str As String) As String
            If str Is Nothing OrElse str.Length = 0 Then
                Return String.Empty
            End If

            Dim sb As New Text.StringBuilder()
            Dim i As Integer = 0
            While i < str.Length
                If str.Chars(i) = "\"c AndAlso i + 1 < str.Length Then
                    i += 1
                    Select Case str.Chars(i)
                        Case "n"c
                            sb.Append(vbLf)
                        Case "r"c
                            sb.Append(vbCr)
                        Case "t"c
                            sb.Append(vbTab)
                        Case "\"c
                            sb.Append("\"c)
                        Case Else
                            sb.Append(str.Chars(i))
                    End Select
                Else
                    sb.Append(str.Chars(i))
                End If
                i += 1
            End While
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>マッチ結果。</returns>
        Public Overrides Function Match(tr As IPositionAdjustReader, env As EBNFEnvironment, ruleName As String) As (success As Boolean, answer As EBNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()
            Dim startPos = tr.Position

            ' バイト配列を読み込み、リテラルと比較する
            Dim buffer = New Char(Me._literal.Length - 1) {}
            Dim count = tr.Read(buffer, 0, buffer.Length)

            Dim unmatched As Integer
            Dim res = EqualString(buffer, count, Me._literal, Me._shiftTable, unmatched)
            If res.sccess Then
                Return (True, New EBNFAnalysisItem("literal", New List(Of EBNFAnalysisItem), tr, startPos, tr.Position))
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me.Range)

            ' 一致しない場合は偽を返す
            snapPos.Restore()
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' 読み取った文字列と指定された文字列が等しいかどうかを判定します。
        ''' </summary>
        ''' <param name="buffer">読み取りバッファ。</param>
        ''' <param name="count">読み取り文字数。</param>
        ''' <param name="stringValue">比較対象の文字列。</param>
        ''' <param name="shiftTb">シフトテーブル。</param>
        ''' <returns>等しい場合は true。それ以外は false。</returns>
        Private Shared Function EqualString(buffer() As Char,
                                            count As Integer,
                                            stringValue As String,
                                            shiftTb As SortedDictionary(Of Char, Integer),
                                            ByRef unmatched As Integer) As (sccess As Boolean, shift As Integer)
            For i As Integer = stringValue.Length - 1 To 0 Step -1
                Dim c = buffer(i)
                If c <> stringValue.Chars(i) Then
                    Dim shift = If(shiftTb.ContainsKey(c), shiftTb(c), stringValue.Length)
                    unmatched = i
                    Return (False, shift)
                End If
            Next
            Return (True, 0)
        End Function

        ''' <summary>
        ''' 次のパターンのマッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public Overrides Function MoveNext(tr As IPositionAdjustReader, env As EBNFEnvironment) As (success As Boolean, answer As EBNFAnalysisItem)
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"Terminal:{Me._literal}"
        End Function

    End Class

End Namespace
