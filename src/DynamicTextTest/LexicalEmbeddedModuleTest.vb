Imports Xunit
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings

Public Class LexicalEmbeddedModuleTest

    <Fact>
    Public Sub SplitEmbeddedText_OnlyPlainText_ReturnsOneNoneBlock()
        Dim input = U8String.NewString("plain text only")
        Dim result = LexicalEmbeddedModule.SplitEmbeddedText(input)
        Assert.Single(result)
        Assert.Equal(EmbeddedType.None, result(0).kind)
        Assert.Equal("plain text only", result(0).str.ToString())
    End Sub

    <Fact>
    Public Sub SplitEmbeddedText_VariableDefineBlock_ReturnsVariableBlock()
        Dim input = U8String.NewString("${var}")
        Dim result = LexicalEmbeddedModule.SplitEmbeddedText(input)
        Assert.Single(result)
        Assert.Equal(EmbeddedType.VariableDefine, result(0).kind)
        Assert.Equal("${var}", result(0).str.ToString())
    End Sub

    <Fact>
    Public Sub SplitEmbeddedText_MixedTextAndVariableBlock()
        Dim input = U8String.NewString("abc${var}def")
        Dim result = LexicalEmbeddedModule.SplitEmbeddedText(input)
        Assert.Equal(3, result.Length)
        Assert.Equal(EmbeddedType.None, result(0).kind)
        Assert.Equal("abc", result(0).str.ToString())
        Assert.Equal(EmbeddedType.VariableDefine, result(1).kind)
        Assert.Equal("${var}", result(1).str.ToString())
        Assert.Equal(EmbeddedType.None, result(2).kind)
        Assert.Equal("def", result(2).str.ToString())
    End Sub

    <Fact>
    Public Sub SplitEmbeddedText_EscapedBraces_AreNotParsedAsBlocks()
        Dim input = U8String.NewString("a\{b\}c")
        Dim result = LexicalEmbeddedModule.SplitEmbeddedText(input)
        Assert.Single(result)
        Assert.Equal(EmbeddedType.None, result(0).kind)
        Assert.Equal("a\{b\}c", result(0).str.ToString())
    End Sub

    <Fact>
    Public Sub SplitEmbeddedText_UnclosedVariableBlock_ThrowsException()
        Dim input = U8String.NewString("${unclosed")
        Assert.Throws(Of AnalysisException)(Function() LexicalEmbeddedModule.SplitEmbeddedText(input))
    End Sub

End Class