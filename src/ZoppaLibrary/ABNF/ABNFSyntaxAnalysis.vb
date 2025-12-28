Option Explicit On
Option Strict On

Imports System.IO
Imports System.Net.Sockets
Imports System.Runtime.CompilerServices
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.EBNF
Imports ZoppaLibrary.EBNF.EBNFSyntaxAnalysis
Imports ZoppaLibrary.Strings

Namespace ABNF

    Public Module ABNFSyntaxAnalysis

        ''' <summary>
        ''' 指定されたルール群から構文解析環境を作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>構文解析環境。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader) As ABNFEnvironment
            '  メソッドテーブルを作成
            Dim answerEnv As New ABNFEnvironment()

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
            For Each sr In range.SubRanges
                Dim key = sr.SubRanges(0).ToString()
                If Not ruleTable.ContainsKey(key) Then
                    ruleTable.Add(key, New RuleAnalysis(key, sr.SubRanges(1)))
                End If
            Next
            Return ruleTable
        End Function

        <Extension()>
        Public Function Evaluate(env As ABNFEnvironment, ident As String, target As PositionAdjustBytes) As ABNFAnalysisItem
            If env.RuleTable.ContainsKey(ident) Then
                Dim startPos = target.Position
                Dim iter = env.RuleTable(ident).GetMatcher()
                Dim res = iter.MoveNext(target, env)
                If res.success AndAlso target.Peek() = -1 Then
                    env.Answer = New ABNFAnalysisItem(ident, iter.GetAnswer(), target, startPos, target.Position)
                    Return env.Answer
                End If
                'Dim startPos = target.Position
                'Dim answers As New List(Of ABNFAnalysisItem)()

                '' 解析実行
                'Dim res = env.RuleTable(ident).Match(target, env, env.RuleTable, ident, answers, New Dictionary(Of IAnalysis, Integer)())

                '' 解析でき、かつ全て消費した場合は成功
                'If res.sccess AndAlso target.Peek() = -1 Then
                '    env.Answer = New ABNFAnalysisItem(ident, answers, target, startPos, target.Position)
                '    Return env.Answer
                'End If
                'env.ThrowFailureException(ident)
                Throw New ABNFException("解析失敗")
            End If
            Throw New ABNFException($"指定された識別子 '{ident}' はルールに存在しません。")
        End Function

#Region "特殊メソッド"

        ''' <summary>
        ''' 英字を読み取ります。
        ''' </summary>
        ''' <param name="tr">テキストリーダー。</param>
        ''' <returns>一致したら真。</returns>
        Private Function ALPHA(tr As IPositionAdjustReader) As Boolean
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
        Private Function DIGIT(tr As IPositionAdjustReader) As Boolean
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
        Private Function HEXDIG(tr As IPositionAdjustReader) As Boolean
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
        Private Function DQUOTE(tr As IPositionAdjustReader) As Boolean
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
        Private Function SP(tr As IPositionAdjustReader) As Boolean
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
        Private Function HTAB(tr As IPositionAdjustReader) As Boolean
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
        Private Function WSP(tr As IPositionAdjustReader) As Boolean
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
        Private Function LWSP(tr As IPositionAdjustReader) As Boolean
            Dim res = False
            Do While True
                Dim c = tr.Peek()
                If c = &H20 OrElse c = &H9 Then
                    tr.Read()
                    res = True
                ElseIf IsCrlfAndWsp(tr) Then
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
        Private Function IsCrLfAndWsp(tr As IPositionAdjustReader) As Boolean
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
        Private Function VCHAR(tr As IPositionAdjustReader) As Boolean
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
        Private Function CHARMethod(tr As IPositionAdjustReader) As Boolean
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
        Private Function OCTET(tr As IPositionAdjustReader) As Boolean
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
        Private Function CTL(tr As IPositionAdjustReader) As Boolean
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
        Private Function CR(tr As IPositionAdjustReader) As Boolean
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
        Private Function LF(tr As IPositionAdjustReader) As Boolean
            If tr.Peek() = &HD Then
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
        Private Function CRLF(tr As IPositionAdjustReader) As Boolean
            Dim snap = tr.MemoryPosition()
            If tr.Peek() = &HD Then
                tr.Read()
                If tr.Peek() = &HA Then
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
        Private Function BIT(tr As IPositionAdjustReader) As Boolean
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
            Public ReadOnly Property MethodTable As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))

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
                'Throw New ABNFException($"識別子 '{ident}' の解析に失敗しました。 ルール: '{Me._failRuleName}', 評価範囲: {Me._failRange}, 文字列: {Me._failTr.Substring(Me._failPos)}")
                Throw New ABNFException($"識別子 '{ident}' の解析に失敗しました。 ルール: '{Me._failRuleName}', 評価範囲: {Me._failRange}")
            End Sub

            ''' <summary>
            ''' ルールグラフをデバッグ出力します。
            ''' </summary>
            ''' <param name="out">出力先のテキストライター。</param>
            Public Sub DebugRuleGraphPrint(out As TextWriter)
                out.WriteLine("***** ABNF ルール *****")
                For Each kvp In Me.RuleTable
                    out.WriteLine($"ルール名: {kvp.Key}")

                    Dim arrivals As New HashSet(Of RuleAnalysis)()
                    DebugRuleGraphPrint(out, arrivals, kvp.Value)
                Next
            End Sub

            ''' <summary>
            ''' ルールグラフをデバッグ出力します。
            ''' </summary>
            ''' <param name="out">出力先のテキストライター。</param>
            ''' <param name="arrivals">到達済みノード集合。</param>
            ''' <param name="node">現在のノード。</param>
            Public Sub DebugRuleGraphPrint(out As TextWriter, arrivals As HashSet(Of RuleAnalysis), node As RuleAnalysis)
                If Not arrivals.Contains(node) Then
                    arrivals.Add(node)

                    'out.Write($"node:{node} -> ")
                    'For Each nextNode In node.Pattern
                    '    out.Write($"{nextNode}, ")
                    'Next
                    'out.WriteLine()

                    'For Each nextNode In node.Pattern
                    '    If TypeOf nextNode.ToAnalysis Is CompletedAnalysis Then
                    '        Continue For
                    '    End If
                    '    DebugRuleGraphPrint(out, arrivals, nextNode.ToAnalysis)
                    'Next
                End If
            End Sub

        End Class

    End Module

End Namespace
