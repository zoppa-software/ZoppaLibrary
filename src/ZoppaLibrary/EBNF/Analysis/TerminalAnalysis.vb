Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 特殊シーケンス解析を表します。
    ''' </summary>
    Public NotInheritable Class TerminalAnalysis
        Implements IAnalysis

        ''' <summary>リテラル文字列。</summary>
        Private ReadOnly _strValue As String

        ''' <summary>評価範囲。</summary>
        Private ReadOnly _range As ExpressionRange

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        Public ReadOnly Property Pattern As List(Of IAnalysis) Implements IAnalysis.Pattern

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(range As ExpressionRange)
            Me._strValue = UnescapedString(range.SubRanges(0).ToString())
            Me._range = range
            Me.Pattern = New List(Of IAnalysis)()
        End Sub

        ''' <summary>
        ''' エスケープシーケンスを変換する。
        ''' </summary>
        ''' <param name="str">変換対象の文字列。</param>
        ''' <returns>変換後の文字列。</returns>
        Private Shared Function UnescapedString(str As String) As String
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
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <param name="ruleTable">ルール解析テーブル。</param>
        ''' <param name="specialMethods">特殊メソッドテーブル。</param>
        ''' <param name="ruleName">現在のルール名。</param>
        ''' <param name="answers">解析結果のリスト。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              env As EBNFEnvironment,
                              ruleTable As SortedDictionary(Of String, RuleAnalysis),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              ruleName As String,
                              answers As List(Of EBNFAnalysisItem)) As Boolean Implements IAnalysis.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of EBNFAnalysisItem)()

            ' リテラル文字列を評価
            Dim hit = False
            Dim buffer = New Char(Me._strValue.Length - 1) {}
            Dim count = tr.Read(buffer, 0, Me._strValue.Length)
            If EqualString(buffer, count, Me._strValue) Then
                answers.Add(New EBNFAnalysisItem("literal", New List(Of EBNFAnalysisItem)(), tr, startPos, tr.Position))
                hit = True
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me._range)

            ' 次のパターンを評価
            If hit Then
                For Each evalExpr In Me.Pattern
                    If evalExpr.Match(tr, env, ruleTable, specialMethods, ruleName, answers) Then
                        Return True
                    End If
                Next
            End If

            ' どれもマッチしなかった場合は偽を返す
            snap.Restore()
            Return False
        End Function

        ''' <summary>
        ''' 読み取った文字列と指定された文字列が等しいかどうかを判定します。
        ''' </summary>
        ''' <param name="buffer">読み取りバッファ。</param>
        ''' <param name="count">読み取り文字数。</param>
        ''' <param name="stringValue">比較対象の文字列。</param>
        ''' <returns>等しい場合は true。それ以外は false。</returns>
        Private Shared Function EqualString(buffer() As Char, count As Integer, stringValue As String) As Boolean
            If count <> stringValue.Length Then
                Return False
            End If
            For i As Integer = 0 To stringValue.Length - 1
                If buffer(i) <> stringValue.Chars(i) Then
                    Return False
                End If
            Next
            Return True
        End Function

    End Class

End Namespace
