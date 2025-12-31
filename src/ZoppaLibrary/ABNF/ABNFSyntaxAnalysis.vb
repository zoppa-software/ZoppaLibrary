Option Explicit On
Option Strict On

Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 構文解析機能を提供します（ABNF）
    ''' </summary>
    Public Module ABNFSyntaxAnalysis

        ''' <summary>
        ''' 指定されたルール群に基づいて構文解析を実行します。
        ''' </summary>
        ''' <param name="rules">ABNF形式のルール群を表す文字列。</param>
        ''' <param name="ident">解析対象のルール名（開始ルール）。</param>
        ''' <param name="target">解析対象のバイト列。</param>
        ''' <param name="addSpecMethods">
        ''' カスタム特殊メソッドを追加するためのデリゲート。
        ''' Nothing の場合は標準メソッドのみを使用します。
        ''' </param>
        ''' <returns>解析が成功した場合は解析結果、失敗した場合は例外をスローします。</returns>
        ''' <exception cref="ABNFException">解析に失敗した場合。</exception>
        ''' <exception cref="ArgumentNullException">引数がNothingの場合。</exception>
        ''' <example>
        ''' <code>
        ''' Dim rules = "expr = DIGIT *( "+" DIGIT )"
        ''' Dim target = New PositionAdjustBytes("1+2+3")
        ''' Dim result = CompileToEvaluater(rules, "expr", target)
        ''' </code>
        ''' </example>
        Public Function CompileToEvaluater(rules As String,
                                           ident As String,
                                           target As PositionAdjustBytes,
                                           Optional addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of PositionAdjustBytes, Boolean))) = Nothing) As ABNFAnalysisItem
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
        ''' <returns>解析結果を表す <see cref="ABNFAnalysisItem"/>。</returns>
        Public Function CompileToEvaluater(rules As IPositionAdjustReader,
                                           ident As String,
                                           target As PositionAdjustBytes,
                                           Optional addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of PositionAdjustBytes, Boolean))) = Nothing) As ABNFAnalysisItem
            Dim env = CompileEnvironment(rules, addSpecMethods)
            Return env.Evaluate(ident, target)
        End Function

        ''' <summary>
        ''' 指定されたルール群から構文解析環境を作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <returns>構文解析環境。</returns>
        Public Function CompileEnvironment(rules As String) As ABNFEnvironment
            Return CompileEnvironment(New PositionAdjustString(rules), Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As String,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of PositionAdjustBytes, Boolean)))) As ABNFEnvironment
            Return CompileEnvironment(New PositionAdjustString(rules), addSpecMethods)
        End Function

        ''' <summary>
        ''' 指定されたルール群から構文解析環境を作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>構文解析環境。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader) As ABNFEnvironment
            Return CompileEnvironment(rules, Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of PositionAdjustBytes, Boolean)))) As ABNFEnvironment
            ' 引数チェック
            If rules Is Nothing Then
                Throw New ArgumentNullException(NameOf(rules))
            End If

            '  メソッドテーブルを作成
            Dim answerEnv As New ABNFEnvironment()
            If addSpecMethods IsNot Nothing Then
                addSpecMethods(answerEnv.MethodTable)
            End If

            ' ルールテーブルを作成
            Dim expr = New RuleListExpression()
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
        ''' <param name="target">解析対象の位置調整リーダー。</param>
        ''' <returns>解析結果を表す <see cref="ABNFAnalysisItem"/>。</returns>
        <Extension()>
        Public Function Evaluate(env As ABNFEnvironment, ident As String, target As PositionAdjustBytes) As ABNFAnalysisItem
            ' 引数チェック
            If String.IsNullOrWhiteSpace(ident) Then
                Throw New ArgumentException("識別子が空です。", NameOf(ident))
            End If

            If target Is Nothing Then
                Throw New ArgumentNullException(NameOf(target))
            End If

            ' ルールの存在確認
            If Not env.RuleTable.ContainsKey(ident) Then
                Throw New ABNFException($"指定された識別子 '{ident}' はルールに存在しません。")
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
                        env.Answer = New ABNFAnalysisItem(ident, matcher.GetAnswer(), target, startPos, target.Position)
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

#Region "特殊メソッド"

        ''' <summary>
        ''' 英字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function ALPHA(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If c >= AscW("A"c) AndAlso c <= AscW("Z"c) OrElse
               c >= AscW("a"c) AndAlso c <= AscW("z"c) Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 数字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function DIGIT(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If c >= AscW("0"c) AndAlso c <= AscW("9"c) Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 16進数字 (0-9 A-F a-f)を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function HEXDIG(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If (c >= AscW("0"c) AndAlso c <= AscW("9"c)) OrElse
               (c >= AscW("A"c) AndAlso c <= AscW("F"c)) OrElse
               (c >= AscW("a"c) AndAlso c <= AscW("f"c)) Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 二重引用符を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function DQUOTE(tr As PositionAdjustBytes) As Boolean
            If tr.Peek() = &H22 Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 空白を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function SP(tr As PositionAdjustBytes) As Boolean
            If tr.Peek() = &H20 Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 水平タブを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function HTAB(tr As PositionAdjustBytes) As Boolean
            If tr.Peek() = &H9 Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 空白と水平タブを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function WSP(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If c = &H20 OrElse c = &H9 Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 空白線型空白と水平タブを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function LWSP(tr As PositionAdjustBytes) As Boolean
            Dim res = False
            Do While True
                Dim c = tr.Peek()
                If c = &H20 OrElse c = &H9 Then
                    tr.Read()
                    res = True
                ElseIf IsCrLfAndWsp(tr) Then
                    tr.Read()
                    res = True
                Else
                    Exit Do
                End If
            Loop
            Return res
        End Function

        ''' <summary>
        ''' 改行 + 空白を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function IsCrLfAndWsp(tr As PositionAdjustBytes) As Boolean
            Dim snap = tr.MemoryPosition()
            If tr.Peek() = &HD Then
                tr.Read()
                If tr.Peek() = &HA Then
                    tr.Read()
                    Dim c = tr.Peek()
                    If c = &H20 OrElse c = &H9 Then
                        Return True
                    End If
                End If
            End If

            snap.Restore()
            Return False
        End Function

        ''' <summary>
        ''' 印字される文字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function VCHAR(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If c >= &H21 AndAlso c <= &H7E Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' NUL以外の任意の7ビットASCII文字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function CHARMethod(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If c >= &H1 AndAlso c <= &H7F Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 8ビットのデータを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function OCTET(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If c >= &H0 AndAlso c <= &HFF Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 制御文字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function CTL(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If (c >= &H0 AndAlso c <= &H1F) OrElse c = &H7F Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 復帰コードを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function CR(tr As PositionAdjustBytes) As Boolean
            If tr.Peek() = &HD Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 改行コードを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function LF(tr As PositionAdjustBytes) As Boolean
            If tr.Peek() = &HA Then  ' LF = 0x0A
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' インターネットの標準改行コードを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function CRLF(tr As PositionAdjustBytes) As Boolean
            Dim snap = tr.MemoryPosition()
            If tr.Peek() = &HD Then
                tr.Read()
                If tr.Peek() = &HA Then
                    tr.Read()  ' LFも読み取る
                    Return True
                End If
            End If

            snap.Restore()
            Return False
        End Function

        ''' <summary>
        ''' ビットを読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function BIT(tr As PositionAdjustBytes) As Boolean
            Dim c = tr.Peek()
            If c = AscW("0"c) OrElse c = AscW("1"c) Then
                tr.Read()
                Return True
            Else
                Return False
            End If
        End Function

#End Region

        ''' <summary>
        ''' 構文解析環境を表します。
        ''' </summary>
        Public NotInheritable Class ABNFEnvironment

            ''' <summary>
            ''' 特殊メソッドテーブル。
            ''' </summary>
            Public ReadOnly Property MethodTable As SortedDictionary(Of String, Func(Of PositionAdjustBytes, Boolean))

            ''' <summary>
            ''' ルールテーブル。
            ''' </summary>
            Public ReadOnly Property RuleTable As SortedDictionary(Of String, RuleAnalysis)

            ''' <summary>
            ''' 解析結果
            ''' </summary>
            Public Property Answer As ABNFAnalysisItem

            ''' <summary>
            ''' 解析失敗情報：失敗したルール名。
            ''' </summary>
            Private _failRuleName As String = ""

            ''' <summary>
            ''' 解析失敗情報：失敗した位置調整リーダー。
            ''' </summary>
            Private _failTr As PositionAdjustBytes = Nothing

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
                Me.MethodTable = New SortedDictionary(Of String, Func(Of PositionAdjustBytes, Boolean))()
                Me.InnerClearSpecialMethods()

                Me.RuleTable = New SortedDictionary(Of String, RuleAnalysis)()
            End Sub

            ''' <summary>
            ''' 特殊メソッドテーブルをクリアします。
            ''' </summary>
            Private Sub InnerClearSpecialMethods()
                Me.MethodTable.Clear()

                ' 標準メソッドを追加
                Me.MethodTable.Add(NameOf(ALPHA), AddressOf ALPHA)
                Me.MethodTable.Add(NameOf(DIGIT), AddressOf DIGIT)
                Me.MethodTable.Add(NameOf(HEXDIG), AddressOf HEXDIG)
                Me.MethodTable.Add(NameOf(DQUOTE), AddressOf DQUOTE)
                Me.MethodTable.Add(NameOf(SP), AddressOf SP)
                Me.MethodTable.Add(NameOf(HTAB), AddressOf HTAB)
                Me.MethodTable.Add(NameOf(WSP), AddressOf WSP)
                Me.MethodTable.Add(NameOf(LWSP), AddressOf LWSP)
                Me.MethodTable.Add(NameOf(VCHAR), AddressOf VCHAR)
                Me.MethodTable.Add("CHAR", AddressOf CHARMethod)
                Me.MethodTable.Add(NameOf(OCTET), AddressOf OCTET)
                Me.MethodTable.Add(NameOf(CTL), AddressOf CTL)
                Me.MethodTable.Add(NameOf(CR), AddressOf CR)
                Me.MethodTable.Add(NameOf(LF), AddressOf LF)
                Me.MethodTable.Add(NameOf(CRLF), AddressOf CRLF)
                Me.MethodTable.Add(NameOf(BIT), AddressOf BIT)
            End Sub

            ''' <summary>
            ''' 特殊メソッドを追加します。
            ''' </summary>
            ''' <param name="name">メソッド名。</param>
            ''' <param name="method">メソッド本体を表すデリゲート。</param>
            ''' <param name="overwrite">既存のメソッドを上書きする場合はTrue。</param>
            ''' <exception cref="ArgumentException">既に登録済みで上書きがFalseの場合。</exception>
            Public Sub AddSpecialMethods(name As String,
                                         method As Func(Of PositionAdjustBytes, Boolean),
                                         Optional overwrite As Boolean = False)
                If String.IsNullOrWhiteSpace(name) Then
                    Throw New ArgumentException("メソッド名が空です。", NameOf(name))
                End If

                If method Is Nothing Then
                    Throw New ArgumentNullException(NameOf(method))
                End If

                If Me.MethodTable.Count <= 0 Then
                    InnerClearSpecialMethods()
                End If

                If Me.MethodTable.ContainsKey(name) Then
                    If overwrite Then
                        Me.MethodTable(name) = method
                    Else
                        Throw New ArgumentException($"メソッド '{name}' は既に登録されています。", NameOf(name))
                    End If
                Else
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
                                             failTr As PositionAdjustBytes,
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
                Throw New ABNFException(msg.ToString())
            End Sub

            ''' <summary>
            ''' ルールグラフをデバッグ出力します。
            ''' </summary>
            ''' <param name="out">出力先のテキストライター。</param>
            ''' <param name="includeDetails">詳細情報を含める場合はTrue。</param>
            Public Sub DebugRuleGraphPrint(out As TextWriter)
                If out Is Nothing Then
                    Throw New ArgumentNullException(NameOf(out))
                End If

                out.WriteLine("***** ABNF ルール *****")
                out.WriteLine($"登録ルール数: {Me.RuleTable.Count}")
                out.WriteLine($"登録メソッド数: {Me.MethodTable.Count}")
                out.WriteLine()

                For Each kvp In Me.RuleTable
                    out.WriteLine($"ルール名: {kvp.Key}")

                    Dim arrivals As New HashSet(Of AnalysisNode)()
                    For Each route In kvp.Value.Routes
                        DebugRuleGraphPrint(out, arrivals, route.NextNode, indent:=1)
                    Next
                    out.WriteLine()
                Next
            End Sub

            ''' <summary>
            ''' ルールグラフをデバッグ出力します（内部実装）。
            ''' </summary>
            ''' <param name="out">出力先のテキストライター。</param>
            ''' <param name="arrivals">到達済みノード集合。</param>
            ''' <param name="node">現在のノード。</param>
            ''' <param name="indent">インデントレベル。</param>
            Private Sub DebugRuleGraphPrint(out As TextWriter,
                                            arrivals As HashSet(Of AnalysisNode),
                                            node As AnalysisNode,
                                            indent As Integer)
                If Not arrivals.Contains(node) Then
                    arrivals.Add(node)

                    ' インデント出力
                    Dim indentStr = New String(" "c, indent * 2)

                    ' ノード情報
                    out.Write($"{indentStr}・Node:{node.Id}")
                    out.Write($" Type:{node.GetType().Name}")
                    If node.Range.Enable Then
                        out.Write($" Range:{node.Range}")
                    End If

                    ' 接続先
                    If node.Routes.Count > 0 Then
                        out.Write(" → [")
                        out.Write(String.Join(", ", node.Routes.Select(Function(r) r.NextNode.Id.ToString())))
                        out.Write("]")
                    End If
                    out.WriteLine()

                    ' 再帰的に出力
                    For Each route In node.Routes
                        DebugRuleGraphPrint(out, arrivals, route.NextNode, indent + 1)
                    Next
                Else
                    ' 既に訪問済みのノードは参照のみ表示
                    Dim indentStr = New String(" "c, indent * 2)
                    out.WriteLine($"{indentStr}・Node:{node.Id} (既出)")
                End If
            End Sub

        End Class

    End Module

End Namespace
