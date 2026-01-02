Option Explicit On
Option Strict On

Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Xml.Schema
Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 構文解析機能を提供します（EBNF）
    ''' </summary>
    Public Module EBNFSyntaxAnalysis

        ''' <summary>
        ''' 指定されたルール群に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">EBNF形式のルール群を表す文字列。</param>
        ''' <param name="ident">解析対象のルール名（開始ルール）。</param>
        ''' <param name="target">解析対象。</param>
        ''' <param name="addSpecMethods">
        ''' カスタム特殊メソッドを追加するためのデリゲート。
        ''' Nothing の場合は標準メソッドのみを使用します。
        ''' </param>
        ''' <returns>解析が成功した場合は解析結果、失敗した場合は例外をスローします。</returns>
        ''' <exception cref="eBNFException">解析に失敗した場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          ident As String,
                                          target As IPositionAdjustReader,
                                          Optional addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))) = Nothing) As EBNFAnalysisItem
            Dim env = CompileEnvironment(rules, addSpecMethods)
            Return env.Evaluate(ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">EBNF形式のルール群を表す文字列。</param>
        ''' <param name="ident">解析対象のルール名（開始ルール）。</param>
        ''' <param name="target">解析対象の文字列。</param>
        ''' <param name="addSpecMethods">
        ''' カスタム特殊メソッドを追加するためのデリゲート。
        ''' Nothing の場合は標準メソッドのみを使用します。
        ''' </param>
        ''' <returns>解析が成功した場合は解析結果、失敗した場合は例外をスローします。</returns>
        ''' <exception cref="eBNFException">解析に失敗した場合。</exception>
        Public Function CompileToEvaluate(rules As String,
                                          ident As String,
                                          target As String,
                                          Optional addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))) = Nothing) As EBNFAnalysisItem
            Dim env = CompileEnvironment(rules, addSpecMethods)
            Return env.Evaluate(ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析対象の識別子。</param>
        ''' <param name="target">解析対象の位置調整リーダー。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>解析結果を表す <see cref="eBNFAnalysisItem"/>。</returns>
        Public Function CompileToEvaluate(rules As IPositionAdjustReader,
                                          ident As String,
                                          target As IPositionAdjustReader,
                                          Optional addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))) = Nothing) As EBNFAnalysisItem
            Dim env = CompileEnvironment(rules, addSpecMethods)
            Return env.Evaluate(ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ident">解析対象の識別子。</param>
        ''' <param name="target">解析対象の文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>解析結果を表す <see cref="eBNFAnalysisItem"/>。</returns>
        Public Function CompileToEvaluate(rules As IPositionAdjustReader,
                                          ident As String,
                                          target As String,
                                          Optional addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))) = Nothing) As EBNFAnalysisItem
            Dim env = CompileEnvironment(rules, addSpecMethods)
            Return env.Evaluate(ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群から構文解析環境を作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <returns>構文解析環境。</returns>
        Public Function CompileEnvironment(rules As String) As EBNFEnvironment
            Return CompileEnvironment(New PositionAdjustString(rules), Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As String,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))) As EBNFEnvironment
            Return CompileEnvironment(New PositionAdjustString(rules), addSpecMethods)
        End Function

        ''' <summary>
        ''' 指定されたルール群から構文解析環境を作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>構文解析環境。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader) As EBNFEnvironment
            Return CompileEnvironment(rules, Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))) As EBNFEnvironment
            ' 引数チェック
            If rules Is Nothing Then
                Throw New ArgumentNullException(NameOf(rules))
            End If

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
        ''' ルールテーブルを作成します。
        ''' </summary>
        ''' <param name="range">ルール群を表す <see cref="ExpressionRange"/>。</param>
        ''' <returns>ルールテーブル。</returns>
        Private Function CreateRuleTable(range As ExpressionRange) As SortedDictionary(Of String, RuleAnalysis)
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()

            ' ルール名ごとの RuleAnalysis を作成
            For Each sr In range.SubRanges
                Dim key = sr.SubRanges(0).ToString()
                If Not ruleTable.ContainsKey(key) Then
                    ruleTable.Add(key, New RuleAnalysis(key, sr.SubRanges(1)))
                End If
            Next

            ' 単純ルートかチェックする
            For Each kvp In ruleTable
                kvp.Value.CheckSimpleRoute(ruleTable)
            Next

            Return ruleTable
        End Function

        ''' <summary>
        ''' 指定された識別子のルールで構文解析を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析対象の識別子。</param>
        ''' <param name="target">解析対象の文字列ストリーム。</param>
        ''' <returns>解析結果を表す <see cref="ABNFAnalysisItem"/>。</returns>
        <Extension()>
        Public Function Evaluate(env As EBNFEnvironment, ident As String, target As IO.TextReader) As EBNFAnalysisItem
            Return Evaluate(env, ident, New PositionAdjustStringReader(target))
        End Function

        ''' <summary>
        ''' 指定された識別子のルールで構文解析を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析対象の識別子。</param>
        ''' <param name="target">解析対象の文字列。</param>
        ''' <returns>解析結果を表す <see cref="ABNFAnalysisItem"/>。</returns>
        <Extension()>
        Public Function Evaluate(env As EBNFEnvironment, ident As String, target As String) As EBNFAnalysisItem
            Return Evaluate(env, ident, New PositionAdjustString(target))
        End Function

        ''' <summary>
        ''' 指定された識別子のルールで構文解析を実行します。
        ''' </summary>
        ''' <param name="env">構文解析環境。</param>
        ''' <param name="ident">解析対象の識別子。</param>
        ''' <param name="target">解析対象の位置調整リーダー。</param>
        ''' <returns>解析結果を表す <see cref="ABNFAnalysisItem"/>。</returns>
        <Extension()>
        Public Function Evaluate(env As EBNFEnvironment, ident As String, target As IPositionAdjustReader) As EBNFAnalysisItem
            ' 引数チェック
            If String.IsNullOrWhiteSpace(ident) Then
                Throw New ArgumentException("識別子が空です。", NameOf(ident))
            End If

            If target Is Nothing Then
                Throw New ArgumentNullException(NameOf(target))
            End If

            ' ルールの存在確認
            If Not env.RuleTable.ContainsKey(ident) Then
                Throw New EBNFException($"指定された識別子 '{ident}' はルールに存在しません。")
            End If

            Dim startPos = target.Position
            Dim matcher = env.RuleTable(ident).GetMatcher()
            matcher.ClearCache()

            ' 全パターンを試行
            Do While target.Peek() <> -1
                Dim res = matcher.MoveNext(target, env)

                If res.success Then
                    ' 全て消費した場合は成功
                    If target.Peek() = -1 Then
                        env.Answer = New EBNFAnalysisItem(ident, matcher.GetAnswer(), target, startPos, target.Position)
                        Return env.Answer
                    End If

                    ' 次のパターンを試行
                    target.Seek(startPos)
                Else
                    ' これ以上のパターンがない
                    Exit Do
                End If
            Loop

            ' 全パターン失敗
            env.ThrowFailureException(ident)
            Return Nothing  ' 到達しないが、コンパイラ警告回避
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
        Public Function Search(env As EBNFEnvironment,
                               ident As String,
                               target As IPositionAdjustReader,
                               Optional searchStart As Integer = 0) As (start As Integer, length As Integer)
            ' 引数チェック
            If String.IsNullOrWhiteSpace(ident) Then
                Throw New ArgumentException("識別子が空です。", NameOf(ident))
            End If

            If target Is Nothing Then
                Throw New ArgumentNullException(NameOf(target))
            End If

            ' ルールの存在確認
            If Not env.RuleTable.ContainsKey(ident) Then
                Throw New EBNFException($"指定された識別子 '{ident}' はルールに存在しません。")
            End If

            Dim startPos = target.Position
            Dim matcher = env.RuleTable(ident).GetMatcher()
            matcher.ClearCache()

            ' 全パターンを試行
            Do While target.Peek() <> -1
                Dim res = matcher.MoveNext(target, env)

                If res.success Then
                    ' 成功
                    env.Answer = New EBNFAnalysisItem(ident, matcher.GetAnswer(), target, startPos, target.Position)
                    Return (startPos, env.Answer.End - env.Answer.Start)
                Else
                    target.Seek(startPos + 1)
                    startPos = target.Position
                End If
            Loop

            Return (-1, 0)






            'If env.RuleTable.ContainsKey(ident) Then
            '    ' 検索開始位置を移動
            '    target.Seek(searchStart)

            '    Dim answers As New List(Of EBNFAnalysisItem)()
            '    Dim startPos = target.Position

            '    ' 検索を実行
            '    Do While target.Peek() <> -1
            '        Dim shift As Integer = Integer.MaxValue

            '        'For Each evalExpr In env.RuleTable(ident).Pattern
            '        '    answers.Clear()
            '        '    Dim res = evalExpr.Match(target, env, env.RuleTable, env.MethodTable, ident, answers)
            '        '    If res.sccess Then
            '        '        env.Answer = New EBNFAnalysisItem(ident, answers, target, startPos, target.Position)
            '        '        Return startPos
            '        '    ElseIf res.shift < shift Then
            '        '        shift = If(res.shift > 0, res.shift, 1)
            '        '    End If
            '        'Next

            '        target.Seek(startPos + shift)
            '        startPos = target.Position
            '    Loop

            '    Return -1
            'Else
            '    Throw New EBNFException($"指定された識別子 '{ident}' はルールに存在しません。")
            'End If
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
        Public Function Search(env As EBNFEnvironment,
                               ident As String,
                               target As String,
                               Optional searchStart As Integer = 0) As (start As Integer, length As Integer)
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
                Dim msg As New StringBuilder()
                msg.Append($"識別子 '{ident}' の解析に失敗しました。 ")
                msg.Append($"ルール:{Me._failRuleName}, ")
                If Me._failRange.Enable Then
                    msg.Append($"評価範囲:{Me._failRange}, ")
                End If
                If Me._failPos >= 0 Then
                    msg.Append($"データ位置:{Me._failPos} データ:{Me._failTr.Substring(Me._failPos, 20)}")
                End If
                Throw New EBNFException(msg.ToString())
            End Sub

            ''' <summary>
            ''' ルールグラフをデバッグ出力します。
            ''' </summary>
            ''' <param name="out">出力先のテキストライター。</param>
            Public Sub DebugRuleGraphPrint(out As TextWriter)
                'out.WriteLine("***** EBNF ルール *****")
                'For Each kvp In Me.RuleTable
                '    out.WriteLine($"ルール名: {kvp.Key}")

                '    Dim arrivals As New HashSet(Of IAnalysis)()
                '    DebugRuleGraphPrint(out, arrivals, kvp.Value)
                'Next
            End Sub

            '''' <summary>
            '''' ルールグラフをデバッグ出力します。
            '''' </summary>
            '''' <param name="out">出力先のテキストライター。</param>
            '''' <param name="arrivals">到達済みノード集合。</param>
            '''' <param name="node">現在のノード。</param>
            'Public Sub DebugRuleGraphPrint(out As TextWriter, arrivals As HashSet(Of IAnalysis), node As IAnalysis)
            '    'If Not arrivals.Contains(node) Then
            '    '    arrivals.Add(node)

            '    '    out.Write($"node:{node} -> ")
            '    '    For Each nextNode In node.Pattern
            '    '        out.Write($"{nextNode}, ")
            '    '    Next
            '    '    out.WriteLine()

            '    '    For Each nextNode In node.Pattern
            '    '        If TypeOf nextNode Is CompletedAnalysis Then
            '    '            Continue For
            '    '        End If
            '    '        DebugRuleGraphPrint(out, arrivals, nextNode)
            '    '    Next
            '    'End If
            'End Sub

        End Class

    End Module

End Namespace
