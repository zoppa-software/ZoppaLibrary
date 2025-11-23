Imports System
Imports Xunit
Imports ZoppaLibrary
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings

Public Class LexicalTest

    <Fact>
    Public Sub LexicalTest_Example()
        Dim ans = LexicalModule.SplitWords(U8String.NewString("true false and or xor in other"))
        Assert.Equal(7, ans.Length)
        Assert.Equal(WordType.TrueLiteral, ans(0).kind)
        Assert.Equal(WordType.FalseLiteral, ans(1).kind)
        Assert.Equal(WordType.AndOperator, ans(2).kind)
        Assert.Equal(WordType.OrOperator, ans(3).kind)
        Assert.Equal(WordType.XorOperator, ans(4).kind)
        Assert.Equal(WordType.InKeyword, ans(5).kind)
        Assert.Equal(WordType.Identifier, ans(6).kind)
    End Sub

    <Fact>
    Public Sub LexicalTest_number()
        Dim ans = LexicalModule.SplitWords(U8String.NewString("123 456.789 1.123 10_1_2_333"))
        Assert.Equal(4, ans.Length)
        Assert.Equal(WordType.Number, ans(0).kind)
        Assert.Equal(WordType.Number, ans(1).kind)
        Assert.Equal(WordType.Number, ans(2).kind)
        Assert.Equal(WordType.Number, ans(3).kind)

        Assert.Throws(Of AnalysisException)(
            Sub()
                LexicalModule.SplitWords(U8String.NewString("123__456_789"))
            End Sub
        )
    End Sub

    <Fact>
    Public Sub LexicalTest_String()
        Dim ans = LexicalModule.SplitWords(U8String.NewString("""Hello"" ""World!"" ""\"""" ""'"""))
        Assert.Equal(4, ans.Length)
        Assert.Equal(WordType.StringLiteral, ans(0).kind)
        Assert.True(ans(0).str.Equals("""Hello"""))
        Assert.Equal(WordType.StringLiteral, ans(1).kind)
        Assert.True(ans(1).str.Equals("""World!"""))
        Assert.Equal(WordType.StringLiteral, ans(2).kind)
        Assert.True(ans(2).str.Equals("""\"""""))
        Assert.Equal(WordType.StringLiteral, ans(3).kind)
        Assert.True(ans(3).str.Equals("""'"""))

        ' 空文字列のテスト
        Dim emptyAns = LexicalModule.SplitWords(U8String.NewString(""""""))
        Assert.Single(emptyAns)
        Assert.Equal(WordType.StringLiteral, emptyAns(0).kind)

        Dim ans1 = LexicalModule.SplitWords(U8String.NewString("'Hello' 'World!' '\'' ''''"))
        Assert.Equal(4, ans1.Length)
        Assert.Equal(WordType.StringLiteral, ans1(0).kind)
        Assert.True(ans1(0).str.Equals("'Hello'"))
        Assert.Equal(WordType.StringLiteral, ans1(1).kind)
        Assert.True(ans1(1).str.Equals("'World!'"))
        Assert.Equal(WordType.StringLiteral, ans1(2).kind)
        Assert.True(ans1(2).str.Equals("'\''"))
        Assert.Equal(WordType.StringLiteral, ans1(3).kind)
        Assert.True(ans1(3).str.Equals("''''"))
    End Sub

    <Fact>
    Public Sub LexicalTest_Operators()
        Dim ans = LexicalModule.SplitWords(U8String.NewString("= + - * / < <= > >= == <>"))
        Assert.Equal(11, ans.Length)
        Assert.Equal(WordType.Assign, ans(0).kind)
        Assert.Equal(WordType.Plus, ans(1).kind)
        Assert.Equal(WordType.Minus, ans(2).kind)
        Assert.Equal(WordType.Multiply, ans(3).kind)
        Assert.Equal(WordType.Divide, ans(4).kind)
        Assert.Equal(WordType.LessThan, ans(5).kind)
        Assert.Equal(WordType.LessEqual, ans(6).kind)
        Assert.Equal(WordType.GreaterThan, ans(7).kind)
        Assert.Equal(WordType.GreaterEqual, ans(8).kind)
        Assert.Equal(WordType.Equal, ans(9).kind)
        Assert.Equal(WordType.NotEqual, ans(10).kind)
    End Sub

    <Fact>
    Public Sub LexicalTest_Identifiers()
        Dim ans = LexicalModule.SplitWords(U8String.NewString("var1 var_2 var3"))
        Assert.Equal(3, ans.Length)
        Assert.Equal(WordType.Identifier, ans(0).kind)
        Assert.Equal("var1", ans(0).str.ToString())
        Assert.Equal(WordType.Identifier, ans(1).kind)
        Assert.Equal("var_2", ans(1).str.ToString())
        Assert.Equal(WordType.Identifier, ans(2).kind)
        Assert.Equal("var3", ans(2).str.ToString())
    End Sub

End Class
