Option Explicit On
Option Strict On

Imports System.Net
Imports System.Text
Imports Xunit
Imports ZoppaLibrary.EBNF

Public Class GrammarExpressionTests2

    <Fact>
    Public Sub NumberTest2()
        Dim input = "" &
"grammar = 'a', ('b' | 'c')+, 'd';"

        Dim compiled = EBNFSyntaxAnalysis.CompileEnvironment(input)
        Dim buf As New StringBuilder()
        Using writer As New IO.StringWriter(buf)
            compiled.DebugRuleGraphPrint(writer)
        End Using

        Dim ans1 = EBNFSyntaxAnalysis.Evaluate(compiled, "grammar", "abcd")

        Assert.Throws(Of EBNFException)(
            Sub()
                EBNFSyntaxAnalysis.Evaluate(compiled, "grammar", "abad")
            End Sub
        )


        Dim ans2 = EBNFSyntaxAnalysis.Search(compiled, "grammar", "123abcd456")
        Assert.Equal(3, ans2)
        Assert.Equal("abcd", compiled.Answer.ToString())
    End Sub

End Class
