Option Explicit On
Option Strict On

Namespace Parser

    Public NotInheritable Class RuleCompiledExpression

        Public ReadOnly Property RuleName As String

        Public ReadOnly Property Pattern As ICompiledExpression

        Public Sub New(name As String, pattern As ExpressionRange)
            Me.RuleName = name
            Me.Pattern = Compile(pattern)
        End Sub

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

        Private Shared Function CompileForSubRanges(target As ExpressionRange) As IEnumerable(Of ICompiledExpression)
            Dim compiledSubRanges As New List(Of ICompiledExpression)()
            For Each sr In target.SubRanges
                compiledSubRanges.Add(Compile(sr))
            Next
            Return compiledSubRanges
        End Function

    End Class

End Namespace
