Option Explicit On
Option Strict On

Namespace EBNF

    Public NotInheritable Class EvaluateExpression
        Public startEps As Boolean
        Public endEps As Boolean
        Public range As ExpressionRange
        Public evaluateExpressions As List(Of EvaluateExpression)

        Public Sub New(startEps As Boolean, endEps As Boolean, range As ExpressionRange)
            Me.startEps = startEps
            Me.endEps = endEps
            Me.range = range
            Me.evaluateExpressions = New List(Of EvaluateExpression)()
        End Sub

        Public Function Match(tr As IPositionAdjustReader,
                              ruleTable As SortedDictionary(Of String, RuleCompiledExpression),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              answers As List(Of EBNFAnalysisItem),
                              debugMode As Boolean,
                              messages As DebugMessage) As Boolean
            Dim snap = tr.MemoryPosition()

            ' 開始位置がεの場合はそのまま進む
            If Me.startEps Then
                For Each evalExpr In Me.evaluateExpressions
                    If evalExpr.Match(tr, ruleTable, specialMethods, answers, debugMode, messages) Then
                        Return True
                    End If
                Next
                snap.Restore()
                Return False
            End If

            ' 終了位置がεの場合はそのまま進む
            If Me.endEps Then
                Return True
            End If

            ' 式の種類に応じて処理を分岐
            If Me.MatchRange(tr, ruleTable, specialMethods, answers, debugMode, messages) Then
                For Each evalExpr In Me.evaluateExpressions
                    If evalExpr.Match(tr, ruleTable, specialMethods, answers, debugMode, messages) Then
                        Return True
                    End If
                Next
            End If

            ' どれもマッチしなかった場合は偽を返す
            snap.Restore()
            Return False
        End Function

        Private Function MatchRange(tr As IPositionAdjustReader,
                                    ruleTable As SortedDictionary(Of String, RuleCompiledExpression),
                                    specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                                    answers As List(Of EBNFAnalysisItem),
                                    debugMode As Boolean,
                                    messages As DebugMessage) As Boolean
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of EBNFAnalysisItem)()

            ' 式の種類に応じて処理を分岐
            Select Case Me.range.Expr.GetType()
                Case GetType(IdentifierExpression)
                    Dim identname = Me.range.ToString()
                    If ruleTable.ContainsKey(identname) Then
                        Dim currPos = tr.Position
                        If ruleTable(identname).Pattern.Match(tr, ruleTable, specialMethods, subAnswers, debugMode, messages) Then
                            answers.Add(New EBNFAnalysisItem(identname, subAnswers, tr, startPos, tr.Position))
                            Return True
                        End If
                    Else
                        Throw New KeyNotFoundException($"識別子 '{identname}' はルールテーブルに存在しません。")
                    End If

                Case GetType(SpecialSeqExpression)
                    Dim methodname = Me.range.SubRanges(0).ToString().Trim()
                    If specialMethods.ContainsKey(methodname) AndAlso
                       specialMethods(methodname)(tr) Then
                        answers.Add(New EBNFAnalysisItem(methodname, New List(Of EBNFAnalysisItem)(), tr, startPos, tr.Position))
                        Return True
                    End If

                Case GetType(TerminalExpression)
                    Dim stringValue = Me.range.SubRanges(0).ToString()
                    Dim readbuffer = New Char(stringValue.Length - 1) {}
                    Dim rcnt = tr.Read(readbuffer, 0, stringValue.Length)
                    If EqualString(readbuffer, rcnt, stringValue) Then
                        answers.Add(New EBNFAnalysisItem("literal", New List(Of EBNFAnalysisItem)(), tr, startPos, tr.Position))
                        Return True
                    End If

                Case GetType(FactorExpression)
                    Dim baseExpr = New RuleCompiledExpression("", Me.range.SubRanges(0))
                    Dim subExpr = New RuleCompiledExpression("", Me.range.SubRanges(2))
                    Dim snap = tr.MemoryPosition()
                    If baseExpr.Pattern.Match(tr, ruleTable, specialMethods, subAnswers, debugMode, messages) Then
                        Dim psnap = tr.MemoryPosition()
                        snap.Restore()
                        If Not subExpr.Pattern.Match(tr, ruleTable, specialMethods, subAnswers, debugMode, messages) Then
                            psnap.Restore()
                            Return True
                        End If
                    End If

                Case Else
                    Throw New InvalidOperationException("未知の式の種類です。")
            End Select

            Return False
        End Function

        ''' <summary>
        ''' 読み取った文字列と指定された文字列が等しいかどうかを判定します。
        ''' </summary>
        ''' <param name="readbuffer">読み取りバッファ。</param>
        ''' <param name="rcnt">読み取り文字数。</param>
        ''' <param name="stringValue">比較対象の文字列。</param>
        ''' <returns>等しい場合は true。それ以外は false。</returns>
        Private Shared Function EqualString(readbuffer() As Char, rcnt As Integer, stringValue As String) As Boolean
            If rcnt <> stringValue.Length Then
                Return False
            End If
            For i As Integer = 0 To stringValue.Length - 1
                If readbuffer(i) <> stringValue.Chars(i) Then
                    Return False
                End If
            Next
            Return True
        End Function

    End Class

End Namespace
