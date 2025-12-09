Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' ルールのコンパイル済み式を表します。
    ''' </summary>
    Public NotInheritable Class RuleCompiledExpression

        ''' <summary>
        ''' ルール名を取得する。
        ''' </summary>
        ''' <returns>ルール名。</returns>
        Public ReadOnly Property RuleName As String

        ''' <summary>
        ''' ルールのパターンを取得する。
        ''' </summary>
        ''' <returns>ルールのパターン。</returns>
        Public ReadOnly Property Pattern As ICompiledExpression

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="name">ルール名。</param>
        ''' <param name="pattern">ルールのパターンを表す <see cref="ExpressionRange"/>。</param>
        Public Sub New(name As String, pattern As ExpressionRange)
            Me.RuleName = name
            Me.Pattern = Compile(pattern)
        End Sub

        ''' <summary>
        ''' 指定された式範囲をコンパイルします。
        ''' </summary>
        ''' <param name="target">式範囲。</param>
        ''' <returns>コンパイル済み式。</returns>
        Private Shared Function Compile(target As ExpressionRange) As ICompiledExpression
            Select Case target.Expr.GetType()
                Case GetType(CharacterExpression)
                    Return New TerminalCompiledExpression(target)
                Case GetType(IdentifierExpression)
                    Return New IdentifierCompiledExpression(target)
                Case GetType(TerminalExpression)
                    Return New TerminalCompiledExpression(target.SubRanges(0))
                Case GetType(SpecialSeqExpression)
                    Return New SpecialSeqCompiledExpression(target)
                Case GetType(TermExpression)
                    Select Case target.SubRanges(0).Expr.GetType()
                        Case GetType(TerminalExpression), GetType(IdentifierExpression)
                            Return Compile(target.SubRanges(0))
                        Case Else
                            Dim compiledSubRanges As New List(Of ICompiledExpression)()
                            For Each sr In target.SubRanges
                                compiledSubRanges.Add(Compile(sr))
                            Next
                            Return New TermCompiledExpression(target, compiledSubRanges)
                    End Select
                Case GetType(FactorExpression)
                    Return If(target.SubRanges.Count > 1,
                        New FactorCompiledExpression(target, CompileForSubRanges(target)),
                        Compile(target.SubRanges(0))
                    )
                Case GetType(ConcatenationExpression)
                    Return If(target.SubRanges.Count > 1,
                        New ConcatenationCompiledExpression(target, CompileForSubRanges(target)),
                        Compile(target.SubRanges(0))
                    )
                Case GetType(AlternationExpression)
                    Return If(target.SubRanges.Count > 1,
                        New AlternationCompiledExpression(target, CompileForSubRanges(target)),
                        Compile(target.SubRanges(0))
                    )
                Case Else
                    Throw New Exception("未知の式タイプです。")
            End Select
        End Function

        ''' <summary>
        ''' 指定された式範囲のサブ式をコンパイルします。
        ''' </summary>
        ''' <param name="target">式範囲。</param>
        ''' <returns>コンパイル済み式の列挙。</returns>
        Private Shared Function CompileForSubRanges(target As ExpressionRange) As IEnumerable(Of ICompiledExpression)
            Dim compiledSubRanges As New List(Of ICompiledExpression)()
            For Each sr In target.SubRanges
                compiledSubRanges.Add(Compile(sr))
            Next
            Return compiledSubRanges
        End Function

    End Class

End Namespace
