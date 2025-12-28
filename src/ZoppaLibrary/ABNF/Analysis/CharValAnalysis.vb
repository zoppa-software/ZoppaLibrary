Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 文字列解析を表します。
    ''' </summary>
    NotInheritable Class CharValAnalysis
        Implements IAnalysis

        ''' <summary>リテラル文字列。</summary>
        Private ReadOnly _strValue As String

        ''' <summary>シフトテーブル。</summary>
        Private ReadOnly _shiftTable As SortedDictionary(Of Char, Integer)

        ''' <summary>評価範囲。</summary>
        Private ReadOnly _range As ExpressionRange

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        Public ReadOnly Property Pattern As List(Of AnalysisRoute)

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(range As ExpressionRange)
            Me._strValue = range.SubRanges(0).ToString()
            Me._shiftTable = New SortedDictionary(Of Char, Integer)()
            For i As Integer = 0 To Me._strValue.Length - 1
                Dim c = Me._strValue.Chars(i)
                Dim nx = Me._strValue.Length - 1 - i
                If Me._shiftTable.ContainsKey(c) Then
                    Me._shiftTable(c) = nx
                Else
                    Me._shiftTable.Add(c, nx)
                End If
            Next

            For Each c As Char In Me._strValue
                If Not Me._shiftTable.ContainsKey(c) Then
                    Me._shiftTable(c) = Me._strValue.Length
                End If
            Next

            Me._range = range
            Me.Pattern = New List(Of AnalysisRoute)()
        End Sub

        ''' <summary>
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <param name="ruleTable">ルール解析テーブル。</param>
        ''' <param name="specialMethods">特殊メソッドテーブル。</param>
        ''' <param name="ruleName">現在のルール名。</param>
        ''' <param name="answers">解析結果のリスト。</param>
        ''' <param name="counter">訪問回数カウンター。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As PositionAdjustBytes,
                              env As ABNFEnvironment,
                              ruleTable As SortedDictionary(Of String, RuleAnalysis),
                              ruleName As String,
                              answers As List(Of ABNFAnalysisItem),
                              counter As Dictionary(Of IAnalysis, Integer)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of ABNFAnalysisItem)()

            ' リテラル文字列を評価
            Dim buffer = New Char(Me._strValue.Length - 1) {}
            Dim count = 0 'tr.Read(buffer, 0, Me._strValue.Length)
            Dim res = EqualString(buffer, count, Me._strValue, Me._shiftTable)
            If res.sccess Then
                answers.Add(New ABNFAnalysisItem("literal", New List(Of ABNFAnalysisItem)(), tr, startPos, tr.Position))
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me._range)

            ' 次のパターンを評価
            If res.sccess Then
                res = Me.AnalysisNextPattern(tr, env, ruleTable, ruleName, answers, counter)
            End If

            ' 解析に失敗した場合は位置を復元
            If Not res.sccess Then
                snap.Restore()
            End If
            Return res
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
                                            shiftTb As SortedDictionary(Of Char, Integer)) As (sccess As Boolean, shift As Integer)
            For i As Integer = stringValue.Length - 1 To 0 Step -1
                Dim c = buffer(i)
                If c <> stringValue.Chars(i) Then
                    Dim shift = If(shiftTb.ContainsKey(c), shiftTb(c), stringValue.Length)
                    Return (False, shift)
                End If
            Next
            Return (True, 0)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"""{Me._range}"""
        End Function

    End Class

End Namespace
