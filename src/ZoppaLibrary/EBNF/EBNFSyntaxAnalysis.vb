Option Explicit On
Option Strict On

Imports System.IO
Imports System.Runtime.CompilerServices

Namespace EBNF

    ''' <summary>
    ''' 構文解析を実行するためのモジュールを表します。
    ''' </summary>
    Public Module EBNFSyntaxAnalysis

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As IPositionAdjustReader,
                                          addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                          ident As String,
                                          target As IPositionAdjustReader) As EBNFEnvironment
            '  メソッドテーブルを作成
            Dim answerEnv = CompileEnvironment(rules, addSpecMethods)

            ' 対象ルールがあるか確認
            If answerEnv.RuleTable.ContainsKey(ident) Then
                Dim startPos = target.Position
                Dim answers As New List(Of EBNFAnalysisItem)()

                ' 解析実行
                Dim res = answerEnv.RuleTable(ident).Match(target, answerEnv, answerEnv.RuleTable, answerEnv.MethodTable, ident, answers)
                If res.sccess AndAlso target.Peek() = -1 Then
                    answerEnv.Answer = New EBNFAnalysisItem(ident, answers, target, startPos, target.Position)
                    Return answerEnv
                End If
                answerEnv.ThrowFailureException(ident)
            End If
            Throw New EBNFException($"指定された識別子 '{ident}' はルールに存在しません。")
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As IPositionAdjustReader,
                                          ident As String,
                                          target As IPositionAdjustReader) As EBNFEnvironment
            Return CompileToEvaluate(rules, Nothing, ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                          ident As String,
                                          target As IPositionAdjustReader) As EBNFEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), addSpecMethods, ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          ident As String,
                                          target As IPositionAdjustReader) As EBNFEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), Nothing, ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                          ident As String,
                                          target As String) As EBNFEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), addSpecMethods, ident, New PositionAdjustStringReader(target))
        End Function

        ''' <summary>
        ''' 指定されたルール群と識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          ident As String,
                                          target As String) As EBNFEnvironment
            Return CompileToEvaluate(New PositionAdjustStringReader(rules), Nothing, ident, New PositionAdjustStringReader(target))
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))) As EBNFEnvironment
            '  メソッドテーブルを作成
            Dim answerEnv As New EBNFEnvironment()
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
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader) As EBNFEnvironment
            Return CompileEnvironment(rules, Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As String,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))),
                                           Optional debugMode As Boolean = False) As EBNFEnvironment
            Return CompileEnvironment(New PositionAdjustStringReader(rules), addSpecMethods)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As String) As EBNFEnvironment
            Return CompileEnvironment(New PositionAdjustStringReader(rules), Nothing)
        End Function

        ''' <summary>
        ''' ルールテーブルを作成します。
        ''' </summary>
        ''' <param name="range">ルール群を表す <see cref="ExpressionRange"/>。</param>
        ''' <returns>ルールテーブル。</returns>
        Private Function CreateRuleTable(range As ExpressionRange) As SortedDictionary(Of String, RuleAnalysis)
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()
            For Each sr In range.SubRanges
                Dim key = sr.SubRanges(0).ToString()
                If Not ruleTable.ContainsKey(key) Then
                    ruleTable.Add(key, New RuleAnalysis(key, sr.SubRanges(1)))
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
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        <Extension()>
        Public Function Evaluate(env As EBNFEnvironment, ident As String, target As IPositionAdjustReader) As EBNFAnalysisItem
            If env.RuleTable.ContainsKey(ident) Then
                Dim startPos = target.Position
                Dim answers As New List(Of EBNFAnalysisItem)()

                ' 解析実行
                Dim res = env.RuleTable(ident).Match(target, env, env.RuleTable, env.MethodTable, ident, answers)

                ' 解析でき、かつ全て消費した場合は成功
                If res.sccess AndAlso target.Peek() = -1 Then
                    env.Answer = New EBNFAnalysisItem(ident, answers, target, startPos, target.Position)
                    Return env.Answer
                End If
                env.ThrowFailureException(ident)
            End If
            Throw New EBNFException($"指定された識別子 '{ident}' はルールに存在しません。")
        End Function

        ''' <summary>
        ''' 指定された識別子、解析対象に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        <Extension()>
        Public Function Evaluate(env As EBNFEnvironment, ident As String, target As String) As EBNFAnalysisItem
            Return Evaluate(env, ident, New PositionAdjustStringReader(target))
        End Function

        ''' <summary>
        ''' 指定された識別子、解析対象に基づいて検索を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="searchStart">検索を開始する位置。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        <Extension()>
        Public Function Search(env As EBNFEnvironment, ident As String, target As IPositionAdjustReader, Optional searchStart As Integer = 0) As Integer
            If env.RuleTable.ContainsKey(ident) Then
                ' 検索開始位置を移動
                target.Seek(searchStart)

                Dim answers As New List(Of EBNFAnalysisItem)()
                Dim startPos = target.Position

                ' 検索を実行
                Do While target.Peek() <> -1
                    Dim shift As Integer = Integer.MaxValue

                    For Each evalExpr In env.RuleTable(ident).Pattern
                        answers.Clear()
                        Dim res = evalExpr.Match(target, env, env.RuleTable, env.MethodTable, ident, answers)
                        If res.sccess Then
                            env.Answer = New EBNFAnalysisItem(ident, answers, target, startPos, target.Position)
                            Return startPos
                        ElseIf res.shift < shift Then
                            shift = If(res.shift > 0, res.shift, 1)
                        End If
                    Next

                    target.Seek(startPos + shift)
                    startPos = target.Position
                Loop

                Return -1
            End If
            Throw New EBNFException($"指定された識別子 '{ident}' はルールに存在しません。")
        End Function

        ''' <summary>
        ''' 指定された識別子、解析対象に基づいて検索を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析を開始する識別子。</param>
        ''' <param name="target">解析対象を表す文字列。</param>
        ''' <param name="searchStart">検索を開始する位置。</param>
        ''' <returns>解析結果を表す <see cref="EBNFEnvironment"/>。</returns>
        ''' <exception cref="ArgumentException">指定された識別子がルールに存在しない場合。</exception>
        <Extension()>
        Public Function Search(env As EBNFEnvironment, ident As String, target As String, Optional searchStart As Integer = 0) As Integer
            Return Search(env, ident, New PositionAdjustStringReader(target))
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
        ''' 空白文字を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property NotSpaceMethodName As String = "Not Space"

        ''' <summary>
        ''' 空白以外の文字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function NotSpace(tr As IPositionAdjustReader) As Boolean
            Dim startPos = tr.Position
            Dim readAny = False
            While Not Char.IsWhiteSpace(ChrW(tr.Peek()))
                tr.Read()
                readAny = True
            End While
            Return readAny
        End Function

        ''' <summary>
        ''' 全ての文字を表す特殊メソッド名を取得します。
        ''' </summary>
        Public ReadOnly Property AllCharMethodName As String = "All Char"

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
        ''' </param>
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
        Public NotInheritable Class EBNFEnvironment

            ''' <summary>
            ''' 特殊メソッドテーブル。
            ''' </summary>
            Public ReadOnly Property MethodTable As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))

            ''' <summary>
            ''' ルールテーブル。
            ''' </summary>
            Public ReadOnly Property RuleTable As SortedDictionary(Of String, RuleAnalysis)

            ''' <summary>
            ''' 解析結果
            ''' </summary>
            Public Property Answer As EBNFAnalysisItem

            ''' <summary>
            ''' 解析失敗情報：失敗したルール名。
            ''' </summary>
            Private _failRuleName As String = ""

            ''' <summary>
            ''' 解析失敗情報：失敗した位置調整リーダー。
            ''' </summary>
            Private _failTr As IPositionAdjustReader = Nothing

            ''' <summary>
            ''' 解析失敗情報：失敗した位置。
            ''' </summary>
            Private _failPos As Integer = -1

            ''' <summary>
            ''' 解析失敗情報：失敗した評価範囲。
            ''' </summary>
            Private _failRange As ExpressionRange = Nothing

            ''' <summary>
            ''' コンストラクター
            ''' </summary>
            Public Sub New()
                Me.MethodTable = New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
                Me.InnerClearSpecialMethods()

                Me.RuleTable = New SortedDictionary(Of String, RuleAnalysis)()
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
                Me.MethodTable.Add(NotSpaceMethodName, AddressOf NotSpace)
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

            ''' <summary>
            ''' 解析失敗情報を設定します。
            ''' </summary>
            ''' <param name="failRuleName">失敗したルール名。</param>
            ''' <param name="failTr">失敗した位置調整リーダー。</param>
            ''' <param name="failPos">失敗した位置。</param>
            ''' <param name="failRange">失敗した評価範囲。</param>
            Public Sub SetFailureInformation(failRuleName As String,
                                             failTr As IPositionAdjustReader,
                                             failPos As Integer,
                                             failRange As ExpressionRange)
                Me._failRuleName = failRuleName
                Me._failTr = failTr
                Me._failPos = failPos
                Me._failRange = failRange
            End Sub

            ''' <summary>
            ''' 解析失敗例外をスローします。
            ''' </summary>
            ''' <param name="ident">失敗した識別子。</param>
            Public Sub ThrowFailureException(ident As String)
                Throw New EBNFException($"識別子 '{ident}' の解析に失敗しました。 ルール: '{Me._failRuleName}', 評価範囲: {Me._failRange}, 文字列: {Me._failTr.Substring(Me._failPos)}")
            End Sub

            ''' <summary>
            ''' ルールグラフをデバッグ出力します。
            ''' </summary>
            ''' <param name="out">出力先のテキストライター。</param>
            Public Sub DebugRuleGraphPrint(out As TextWriter)
                out.WriteLine("***** EBNF ルール *****")
                For Each kvp In Me.RuleTable
                    out.WriteLine($"ルール名: {kvp.Key}")

                    Dim arrivals As New HashSet(Of IAnalysis)()
                    DebugRuleGraphPrint(out, arrivals, kvp.Value)
                Next
            End Sub

            ''' <summary>
            ''' ルールグラフをデバッグ出力します。
            ''' </summary>
            ''' <param name="out">出力先のテキストライター。</param>
            ''' <param name="arrivals">到達済みノード集合。</param>
            ''' <param name="node">現在のノード。</param>
            Public Sub DebugRuleGraphPrint(out As TextWriter, arrivals As HashSet(Of IAnalysis), node As IAnalysis)
                If Not arrivals.Contains(node) Then
                    arrivals.Add(node)

                    out.Write($"node:{node} -> ")
                    For Each nextNode In node.Pattern
                        out.Write($"{nextNode}, ")
                    Next
                    out.WriteLine()

                    For Each nextNode In node.Pattern
                        If TypeOf nextNode Is CompletedAnalysis Then
                            Continue For
                        End If
                        DebugRuleGraphPrint(out, arrivals, nextNode)
                    Next
                End If
            End Sub

        End Class

    End Module

End Namespace
