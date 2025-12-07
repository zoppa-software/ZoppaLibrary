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
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As IPositionAdjustReader,
                                          addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                          ident As String,
                                          target As IPositionAdjustReader,
                                          Optional debugMode As Boolean = False) As AnalysisEnvironment
            '  メソッドテーブルを作成
            Dim answerEnv = CompileEnvironment(rules, addSpecMethods, debugMode)

            ' 解析を実行
            If answerEnv.RuleTable.ContainsKey(ident) Then
                Dim startPos = target.Position
                Dim answers As New List(Of AnalysisRange)()
                Dim message As New DebugMessage(target)
                If answerEnv.RuleTable(ident).Pattern.Match(target, answerEnv.RuleTable, answerEnv.MethodTable, answers, debugMode, message) Then
                    If target.Peek() = -1 Then
                        ' 解析成功
                        answerEnv.Answer = New AnalysisRange(ident, answers, target, startPos, target.Position)
                        Return answerEnv
                    End If
                End If
                Throw New ArgumentException($"識別子 '{ident}' の解析に失敗しました。...{message.GetUnmatchedMessage()}")
            End If
            Throw New ArgumentException($"指定された識別子 '{ident}' はルールに存在しません。", ident)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As IPositionAdjustReader,
                                          ident As String,
                                          target As IPositionAdjustReader,
                                          Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileToEvaluate(rules, Nothing, ident, target, debugMode)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                          ident As String,
                                          target As IPositionAdjustReader,
                                          Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), addSpecMethods, ident, target, debugMode)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          ident As String,
                                          target As IPositionAdjustReader,
                                          Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), Nothing, ident, target, debugMode)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                          ident As String,
                                          target As String,
                                          Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), addSpecMethods, ident, New PositionAdjustStringReader(target), debugMode)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>解析結果を表す <see cref="AnalysisEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          ident As String,
                                          target As String,
                                          Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), Nothing, ident, New PositionAdjustStringReader(target), debugMode)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                           Optional debugMode As Boolean = False) As AnalysisEnvironment
            '  メソッドテーブルを作成
            Dim answerEnv As New AnalysisEnvironment(debugMode)
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
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader, Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileEnvironment(rules, Nothing, debugMode)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As String,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                           Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileEnvironment(New PositionAdjustStringReader(rules), addSpecMethods, debugMode)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
        ''' <returns>ルールテーブルを含む <see cref="AnalysisEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As String, Optional debugMode As Boolean = False) As AnalysisEnvironment
            Return CompileEnvironment(New PositionAdjustStringReader(rules), Nothing, debugMode)
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
        Public Function Evaluate(env As AnalysisEnvironment, ident As String, target As IPositionAdjustReader) As AnalysisRange
            If env.RuleTable.ContainsKey(ident) Then
                Dim startPos = target.Position
                Dim answers As New List(Of AnalysisRange)()
                Dim message As New DebugMessage(target)
                If env.RuleTable(ident).Pattern.Match(target, env.RuleTable, env.MethodTable, answers, env.DebugMode, message) Then
                    If target.Peek() = -1 Then
                        ' 解析成功
                        env.Answer = New AnalysisRange(ident, answers, target, startPos, target.Position)
                        Return env.Answer
                    End If
                End If
                Throw New ArgumentException($"識別子 '{ident}' の解析に失敗しました。...{message.GetUnmatchedMessage()}")
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
        Public Function Evaluate(env As AnalysisEnvironment, ident As String, target As String) As AnalysisRange
            Return Evaluate(env, ident, New PositionAdjustStringReader(target))
        End Function

#Region "特殊メソッド"

        ''' <summary>
        ''' 空白文字を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property SpaceMethodName As String = NameOf(Space)

        ''' <summary>
        ''' 空白文字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function Space(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While Char.IsWhiteSpace(ChrW(tr.Peek()))
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        ''' <summary>
        ''' 全ての文字を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property AllCharMethodName As String = NameOf(AllChar)

        ''' <summary>
        ''' 全ての文字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function AllChar(tr As IPositionAdjustReader) As Boolean
            Dim c = tr.Read()
            Return (c > 0)
        End Function

        ''' <summary>
        ''' 英字を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property AlphaMethodName As String = NameOf(Alpha)

        ''' <summary>
        ''' 英字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
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

        ''' <summary>
        ''' ASCII文字を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property AsciiCharMethodName As String = NameOf(AsciiChar)

        ''' <summary>
        ''' ASCII文字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns> 
        Private Function AsciiChar(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While tr.Peek() >= 32 AndAlso tr.Peek() <= 126
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        ''' <summary>
        ''' 数字を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property DigitMethodName As String = NameOf(Digit)

        ''' <summary>
        ''' 数字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function Digit(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While Char.IsDigit(ChrW(tr.Peek()))
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        Public ReadOnly Property HexingMethodName As String = NameOf(Hexing)

        ''' <summary>
        ''' 16進数を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function Hexing(tr As IPositionAdjustReader) As Boolean
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

        ''' <summary>
        ''' 整数を表す特殊メソッド名を取得します。
        ''' </summary>
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

        ''' <summary>
        ''' 数値を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property NumberMethodName As String = NameOf(Number)

        ''' <summary>
        ''' 数値を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
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
                    ReadSeqDigits(tr)
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
                    ReadSeqDigits(tr)
                Else
                    snap2.Restore()
                End If
            End If

            Return True
        End Function

        ''' <summary>
        ''' 連続する数字列を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。
        ''' </summary>
        ''' <param name="tr"></param>
        Private Sub ReadSeqDigits(tr As IPositionAdjustReader)
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
            ''' 特殊メソッドテーブル。
            ''' </summary>
            Public ReadOnly Property MethodTable As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))

            ''' <summary>
            ''' ルールテーブル。
            ''' </summary>
            Public ReadOnly Property RuleTable As SortedDictionary(Of String, RuleCompiledExpression)

            ''' <summary>
            ''' デバッグモードで動作するかどうかを示す値。
            ''' </summary>
            Public ReadOnly Property DebugMode As Boolean

            ''' <summary>
            ''' 解析結果
            ''' </summary>
            Public Property Answer As AnalysisRange

            ''' <summary>
            ''' コンストラクター
            ''' </summary>
            ''' <param name="debugMode">デバッグモードで動作するかどうかを示す値。</param>
            Public Sub New(debugMode As Boolean)
                Me.MethodTable = New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
                Me.InnerClearSpecialMethods()

                Me.RuleTable = New SortedDictionary(Of String, RuleCompiledExpression)()

                Me.DebugMode = debugMode
            End Sub

            ''' <summary>
            ''' 特殊メソッドテーブルをクリアします。
            ''' </summary>
            Private Sub InnerClearSpecialMethods()
                Me.MethodTable.Clear()

                ' 標準メソッドを追加
                Me.MethodTable.Add(AllCharMethodName, AddressOf AllChar)
                Me.MethodTable.Add(AlphaMethodName, AddressOf Alpha)
                Me.MethodTable.Add(DigitMethodName, AddressOf Digit)
                Me.MethodTable.Add(HexingMethodName, AddressOf Hexing)
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

        ''' <summary>
        ''' デバッグメッセージを保持するクラス。
        ''' </summary>
        Public NotInheritable Class DebugMessage

            ''' <summary>
            ''' 文字参照インターフェース。
            ''' </summary>
            Private ReadOnly _tar As IPositionAdjustReader

            ''' <summary>
            ''' メッセージリスト。
            ''' </summary>
            Private ReadOnly _msgs As New List(Of String)

            ''' <summary>
            ''' 不一致メッセージ。
            ''' </summary>
            Private _unmatched As String = ""

            ''' <summary>
            ''' コンストラクタ。
            ''' </summary>
            ''' <param name="target">文字参照。</param>
            Public Sub New(target As IPositionAdjustReader)
                _tar = target
            End Sub

            ''' <summary>
            ''' メッセージを追加する。
            ''' </summary>
            ''' <param name="msg">メッセージ。</param>
            Public Sub Add(msg As String)
                _msgs.Add(msg)
            End Sub

            ''' <summary>
            ''' 不一致メッセージを設定する。
            ''' </summary>
            ''' <param name="msg">不一致メッセージ。</param>
            Public Sub SetUnmatched(msg As String)
                _unmatched = msg
            End Sub

            ''' <summary>
            ''' 不一致メッセージを取得する。
            ''' </summary>
            ''' <returns>不一致メッセージ。</returns>
            Public Function GetUnmatchedMessage() As String
                Return If(_unmatched <> "", _unmatched, _tar.Substring(_tar.Position))
            End Function

        End Class

    End Module

End Namespace
