Option Explicit On
Option Strict On

Imports System.Runtime.CompilerServices

Namespace Parser

    ''' <summary>
    ''' 構文解析を実行するためのモジュールを表します。
    ''' </summary>
    Public Module SyntaxAnalysis

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As IPositionAdjustReader,
                                        addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                        ident As String,
                                        target As IPositionAdjustReader) As AnalysisEnvironment
            '  メソッドテーブルを作成
            Dim answerEnv = CreateRuleTable(rules, addSpecMethods)

            ' 解析を実行
            If answerEnv.RuleTable.ContainsKey(ident) Then
                Dim startPos = target.Position
                Dim answers As New List(Of AnalysisRange)()
                If answerEnv.RuleTable(ident).Pattern.Match(target, answerEnv.RuleTable, answerEnv.MethodTable, answers) Then
                    If target.Peek() = -1 Then
                        ' 解析成功
                        answerEnv.Answer = New AnalysisRange(ident, answers, target, startPos, target.Position)
                        Return answerEnv
                    End If
                End If
                Throw New ArgumentException($"識別子 '{ident}' の解析に失敗しました。...{target.ToLastString(50)}")
            End If
            Throw New ArgumentException($"指定された識別子 '{ident}' はルールに存在しません。", ident)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As IPositionAdjustReader,
                                        ident As String,
                                        target As IPositionAdjustReader) As AnalysisEnvironment
            Return LexicalAnalysis(rules, Nothing, ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As String,
                                        addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                        ident As String,
                                        target As IPositionAdjustReader) As AnalysisEnvironment
            Return LexicalAnalysis(New PositionAdjustStringReader(rules), addSpecMethods, ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As String,
                                        ident As String,
                                        target As IPositionAdjustReader) As AnalysisEnvironment
            Return LexicalAnalysis(New PositionAdjustStringReader(rules), Nothing, ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As String,
                                        addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                        ident As String,
                                        target As String) As AnalysisEnvironment
            Return LexicalAnalysis(New PositionAdjustStringReader(rules), addSpecMethods, ident, New PositionAdjustStringReader(target))
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As String,
                                        ident As String,
                                        target As String) As AnalysisEnvironment
            Return LexicalAnalysis(New PositionAdjustStringReader(rules), Nothing, ident, New PositionAdjustStringReader(target))
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CreateRuleTable(rules As IPositionAdjustReader,
                                        addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))) As AnalysisEnvironment
            '  メソッドテーブルを作成
            Dim answerEnv As New AnalysisEnvironment()
            If addSpecMethods IsNot Nothing Then
                addSpecMethods(answerEnv.MethodTable)
            End If

            ' ルールテーブルを作成
            Dim expr = New GrammarExpression()
            Dim range = expr.Match(rules)
            For Each kvp In CreateRuleTable(range)
                answerEnv.RuleTable.Add(kvp.Key, kvp.Value)
            Next
            Return answerEnv
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CreateRuleTable(rules As IPositionAdjustReader) As AnalysisEnvironment
            Return CreateRuleTable(rules, Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CreateRuleTable(rules As String,
                                        addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))) As AnalysisEnvironment
            Return CreateRuleTable(New PositionAdjustStringReader(rules), addSpecMethods)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CreateRuleTable(rules As String) As AnalysisEnvironment
            Return CreateRuleTable(New PositionAdjustStringReader(rules), Nothing)
        End Function

        ''' <summary>
        ''' ルールテーブルを作成します。
        ''' </summary>
        ''' <param name="range">ルール群を表す <see cref="ExpressionRange"/>。</param>
        ''' <returns>ルールテーブル。</returns>
        Private Function CreateRuleTable(range As ExpressionRange) As SortedDictionary(Of String, RuleCompiledExpression)
            Dim ruleTable As New SortedDictionary(Of String, RuleCompiledExpression)()
            For Each sr In range.SubRanges
                Dim key = sr.SubRanges(0).ToString()
                If Not ruleTable.ContainsKey(key) Then
                    ruleTable.Add(key, New RuleCompiledExpression(key, sr.SubRanges(1)))
                End If
            Next
            Return ruleTable
        End Function

        ''' <summary>
        ''' 指定された識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        <Extension()>
        Public Function LexicalAnalysis(env As AnalysisEnvironment, ident As String, target As IPositionAdjustReader) As AnalysisRange
            If env.RuleTable.ContainsKey(ident) Then
                Dim startPos = target.Position
                Dim answers As New List(Of AnalysisRange)()
                If env.RuleTable(ident).Pattern.Match(target, env.RuleTable, env.MethodTable, answers) Then
                    If target.Peek() = -1 Then
                        ' 解析成功
                        env.Answer = New AnalysisRange(ident, answers, target, startPos, target.Position)
                        Return env.Answer
                    End If
                End If
                Throw New ArgumentException($"識別子 '{ident}' の解析に失敗しました。...{target.ToLastString(50)}")
            End If
            Throw New ArgumentException($"指定された識別子 '{ident}' はルールに存在しません。", ident)
        End Function

        ''' <summary>
        ''' 指定された識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        <Extension()>
        Public Function LexicalAnalysis(env As AnalysisEnvironment, ident As String, target As String) As AnalysisRange
            Return LexicalAnalysis(env, ident, New PositionAdjustStringReader(target))
        End Function

#Region "特殊メソッド"

        Public ReadOnly Property SpaceMethodName As String = NameOf(Space)

        Private Function Space(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While Char.IsWhiteSpace(ChrW(tr.Peek()))
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        Public ReadOnly Property AlphaMethodName As String = NameOf(Alpha)

        Private Function Alpha(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While tr.Peek() >= AscW("A"c) AndAlso tr.Peek() <= AscW("Z"c) OrElse
                  tr.Peek() >= AscW("a"c) AndAlso tr.Peek() <= AscW("z"c)
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        Public ReadOnly Property DigitMethodName As String = NameOf(Digit)

        Private Function Digit(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While Char.IsDigit(ChrW(tr.Peek()))
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        Public ReadOnly Property HexdigMethodName As String = NameOf(Hexdig)

        Private Function Hexdig(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While Char.IsDigit(ChrW(tr.Peek())) OrElse
                  (tr.Peek() >= AscW("A"c) AndAlso tr.Peek() <= AscW("F"c)) OrElse
                  (tr.Peek() >= AscW("a"c) AndAlso tr.Peek() <= AscW("f"c))
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        Public ReadOnly Property IntegerMethodName As String = NameOf([Integer])

        ''' <summary>
        ''' 整数を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function [Integer](tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim snap = tr.MemoryPosition()
            Dim readAny = False

            ' 符号チェック
            Dim sign = tr.Peek()
            If sign = AscW("+"c) OrElse sign = AscW("-") Then
                tr.Read()
            End If

            ' 数値チェック
            Dim num = tr.Peek()
            If num = AscW("0"c) Then
                ' 0始まり
                tr.Read()
                readAny = True
            Else
                ' 0以外の数字始まり
                If num >= AscW("1"c) AndAlso num <= AscW("9"c) Then
                    tr.Read()
                    readAny = True
                End If

                Do While True
                    ' _ チェック
                    num = tr.Peek()
                    If num = AscW("_"c) Then
                        tr.Read()
                        readAny = True
                    End If

                    ' 数字チェック
                    num = tr.Peek()
                    If num >= AscW("0"c) AndAlso num <= AscW("9"c) Then
                        tr.Read()
                        readAny = True
                    Else
                        Exit Do
                    End If
                    tr.Read()
                    readAny = True
                Loop
            End If

            If readAny Then
                Return True
            Else
                snap.Restore()
                Return False
            End If
        End Function

        Public ReadOnly Property NumberMethodName As String = NameOf(Number)

        Private Function Number(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position

            If Not [Integer](tr) Then
                Return False
            End If

            ' 小数点の判定
            Dim snap = tr.MemoryPosition()
            Dim dec = tr.Peek()
            If dec = AscW("."c) Then
                ' 小数点以下
                tr.Read()

                ' 数字
                Dim num = tr.Peek()
                If num >= AscW("0"c) AndAlso num <= AscW("9"c) Then
                    ReadSeqDidit(tr)
                Else
                    snap.Restore()
                End If
            End If

            Dim snap2 = tr.MemoryPosition()
            Dim ecode = tr.Peek()
            If ecode = AscW("E"c) OrElse ecode = AscW("e"c) Then
                ' 指数部
                tr.Read()

                ' 符号
                Dim sign = tr.Peek()
                If sign = AscW("+"c) OrElse sign = AscW("-"c) Then
                    tr.Read()
                End If

                ' 数字
                Dim num = tr.Peek()
                If num >= AscW("0"c) AndAlso num <= AscW("9"c) Then
                    ReadSeqDidit(tr)
                Else
                    snap2.Restore()
                End If
            End If

            Return True
        End Function

        Private Sub ReadSeqDidit(tr As IPositionAdjustReader)
            tr.Read()
            Do While True
                ' _ チェック
                Dim num = tr.Peek()
                If num = AscW("_"c) Then
                    tr.Read()
                End If

                ' 数字チェック
                num = tr.Peek()
                If num >= AscW("0"c) AndAlso num <= AscW("9"c) Then
                    tr.Read()
                Else
                    Exit Do
                End If
            Loop
        End Sub

#End Region

        ''' <summary>
        ''' 構文解析環境を表します。
        ''' </summary>
        Public NotInheritable Class AnalysisEnvironment

            ''' <summary>
            ''' 特殊メソッドテーブル
            ''' </summary>
            Public ReadOnly Property MethodTable As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))

            ''' <summary>
            ''' ルールテーブル
            ''' </summary>
            Public ReadOnly Property RuleTable As SortedDictionary(Of String, RuleCompiledExpression)

            ''' <summary>
            ''' 解析結果
            ''' </summary>
            Public Property Answer As AnalysisRange

            ''' <summary>
            ''' コンストラクター
            ''' </summary>
            Public Sub New()
                Me.MethodTable = New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
                Me.InnerClearSpecialMethods()

                Me.RuleTable = New SortedDictionary(Of String, RuleCompiledExpression)()
            End Sub

            ''' <summary>
            ''' 特殊メソッドテーブルをクリアします。
            ''' </summary>
            Private Sub InnerClearSpecialMethods()
                Me.MethodTable.Clear()

                ' 標準メソッドを追加
                Me.MethodTable.Add(AlphaMethodName, AddressOf Alpha)
                Me.MethodTable.Add(DigitMethodName, AddressOf Digit)
                Me.MethodTable.Add(HexdigMethodName, AddressOf Hexdig)
                Me.MethodTable.Add(IntegerMethodName, AddressOf [Integer])
                Me.MethodTable.Add(NumberMethodName, AddressOf Number)
                Me.MethodTable.Add(SpaceMethodName, AddressOf Space)
            End Sub

            ''' <summary>
            ''' 特殊メソッドを追加します。
            ''' </summary>
            ''' <param name="name">メソッド名。</param>
            ''' <param name="method">メソッド本体を表すデリゲート。</param>
            Public Sub AddSpecialMethods(name As String, method As Func(Of IPositionAdjustReader, Boolean))
                If Me.MethodTable.Count <= 0 Then
                    InnerClearSpecialMethods()
                End If
                If Not Me.MethodTable.ContainsKey(name) Then
                    Me.MethodTable.Add(name, method)
                End If
            End Sub

        End Class

    End Module

End Namespace
