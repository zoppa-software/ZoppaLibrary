Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 構文解析を実行するためのモジュールを表します。
    ''' </summary>
    Public Module SyntaxAnalysis

        ''' <summary>
        ''' 特殊メソッドテーブル
        ''' </summary>
        Private _methodTable As New Lazy(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))(
            Function()
                Return New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
            End Function
        )

        ''' <summary>
        ''' ルールテーブル
        ''' </summary>
        Private _ruleTable As New Lazy(Of SortedDictionary(Of String, RuleCompiledExpression))(
            Function()
                Return New SortedDictionary(Of String, RuleCompiledExpression)()
            End Function
        )

        ''' <summary>
        ''' 特殊メソッドテーブルをクリアします。
        ''' </summary>
        Public Sub ClearSpecialMethods()
            SyncLock _methodTable.Value
                InnerClearSpecialMethods()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 特殊メソッドテーブルをクリアします。
        ''' </summary>
        Private Sub InnerClearSpecialMethods()
            If _methodTable.IsValueCreated Then
                _methodTable.Value.Clear()
            End If

            ' 標準メソッドを追加
            _methodTable.Value.Add(AlphaMethodName, AddressOf Alpha)
            _methodTable.Value.Add(DigitMethodName, AddressOf Digit)
            _methodTable.Value.Add(HexdigMethodName, AddressOf Hexdig)
            _methodTable.Value.Add(IntegerMethodName, AddressOf [Integer])
            _methodTable.Value.Add(NumberMethodName, AddressOf Number)
            _methodTable.Value.Add(SpaceMethodName, AddressOf Space)
        End Sub

        ''' <summary>
        ''' 特殊メソッドを追加します。
        ''' </summary>
        ''' <param name="name">メソッド名。</param>
        ''' <param name="method">メソッド本体を表すデリゲート。</param>
        Public Sub AddSpecialMethods(name As String, method As Func(Of IPositionAdjustReader, Boolean))
            SyncLock _methodTable.Value
                If _methodTable.Value.Count <= 0 Then
                    ClearSpecialMethods()
                End If
                If Not _methodTable.Value.ContainsKey(name) Then
                    _methodTable.Value.Add(name, method)
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisRange"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As IPositionAdjustReader, ident As String, target As IPositionAdjustReader) As AnalysisRange
            '  メソッドテーブルを作成
            Dim specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))
            SyncLock _methodTable.Value
                If _methodTable.Value.Count <= 0 Then
                    ClearSpecialMethods()
                End If
                specialMethods = New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))(_methodTable.Value)
            End SyncLock

            ' ルールテーブルを作成
            Dim expr = New GrammarExpression()
            Dim range = expr.Match(rules)
            SyncLock _ruleTable.Value
                If _ruleTable.IsValueCreated Then
                    _ruleTable.Value.Clear()
                End If
                For Each kvp In CreateRuleTable(range)
                    _ruleTable.Value.Add(kvp.Key, kvp.Value)
                Next

                ' 解析を実行
                If _ruleTable.Value.ContainsKey(ident) Then
                    Dim startPos = target.Position
                    Dim answers As New List(Of AnalysisRange)()
                    If _ruleTable.Value(ident).Pattern.Match(target, _ruleTable.Value, specialMethods, answers) Then
                        If target.Peek() = -1 Then
                            ' 解析成功
                            Return New AnalysisRange(ident, answers, target, startPos, target.Position)
                        End If
                    End If
                    Throw New ArgumentException($"識別子 '{ident}' の解析に失敗しました。...{target.ToLastString(50)}")
                End If
                Throw New ArgumentException($"指定された識別子 '{ident}' はルールに存在しません。", ident)
            End SyncLock
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisRange"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As String, ident As String, target As IPositionAdjustReader) As AnalysisRange
            Return LexicalAnalysis(New PositionAdjustStringReader(rules), ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す文字列。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(rules As String, ident As String, target As String) As AnalysisRange
            Return LexicalAnalysis(New PositionAdjustStringReader(rules), ident, New PositionAdjustStringReader(target))
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
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisRange"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(ident As String, target As IPositionAdjustReader) As AnalysisRange
            '  メソッドテーブルを作成
            Dim specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))
            SyncLock _methodTable.Value
                If _methodTable.Value.Count <= 0 Then
                    ClearSpecialMethods()
                End If
                specialMethods = New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))(_methodTable.Value)
            End SyncLock

            ' 解析を実行
            SyncLock _ruleTable.Value
                If _ruleTable.Value.ContainsKey(ident) Then
                    Dim startPos = target.Position
                    Dim answers As New List(Of AnalysisRange)()
                    If _ruleTable.Value(ident).Pattern.Match(target, _ruleTable.Value, specialMethods, answers) Then
                        If target.Peek() = -1 Then
                            ' 解析成功
                            Return New AnalysisRange(ident, answers, target, startPos, target.Position)
                        End If
                    End If
                    Throw New ArgumentException($"識別子 '{ident}' の解析に失敗しました。...{target.ToLastString(50)}")
                End If
                Throw New ArgumentException($"指定された識別子 '{ident}' はルールに存在しません。", ident)
            End SyncLock
        End Function

        ''' <summary>
        ''' 指定された識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisRange"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function LexicalAnalysis(ident As String, target As String) As AnalysisRange
            Return LexicalAnalysis(ident, New PositionAdjustStringReader(target))
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

    End Module

End Namespace
