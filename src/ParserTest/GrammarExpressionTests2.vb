Option Explicit On
Option Strict On

Imports System.Net
Imports System.Text
Imports Xunit
Imports ZoppaLibrary.EBNF

Public Class GrammarExpressionTests2

    <Fact>
    Public Sub GrammarTest1()
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

    <Fact>
    Public Sub GrammarTest2()
        Dim input = "" &
"SELECT * FROM テーブル名

INSERT INTO テーブル名
   VALUES ('1', 'タカシ', '初ツイート！', '2017/07/05' ,'2017/07/05')

UPDATE テーブル名
 SET カラム名 = 値, カラム名 = 値 
 WHERE id = 1

DELETE FROM テーブル名
    WHERE 条件"

        Dim pattern = "" &
"S = {? Space ?};
ident = ? Not Space ?+;
grammar = 'UPDATE', S, ident, S, 'SET';"

        Dim compiled = EBNFSyntaxAnalysis.CompileEnvironment(pattern)
        Dim ans1 = EBNFSyntaxAnalysis.Search(compiled, "grammar", input)
        Assert.Equal("UPDATE テーブル名
 SET", compiled.Answer.ToString())
        Assert.Equal("テーブル名", compiled.Answer("ident").ToString())
    End Sub

End Class
