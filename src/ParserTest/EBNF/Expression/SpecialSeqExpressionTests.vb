Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit
Imports ZoppaLibrary.BNF

Public Class SpecialSeqExpressionTests

    <Fact>
    Public Sub Match_SimpleSequence_ReturnsRange()
        Dim tr As New PositionAdjustString("?abc?")
        Dim expr As New SpecialSeqExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(startPos, r.Start)
        Assert.Equal(tr.Position, r.End)
        Assert.Equal("?abc?", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_EmptySequence_ReturnsRange()
        Dim tr As New PositionAdjustString("??rest")
        Dim expr As New SpecialSeqExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal("??", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_EscapedQuestionInside_DoesNotCloseEarly()
        Dim tr As New PositionAdjustString("?a\?b?")
        Dim expr As New SpecialSeqExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal("?a\?b?", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_EscapedBackslashInside_AllowsLiteralBackslash()
        Dim tr As New PositionAdjustString("?\\" & "\" & "?") ' "?\\?"
        Dim expr As New SpecialSeqExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal("?\\" & "\" & "?", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_UnclosedSequence_RestoresPositionAndReturnsInvalid()
        Dim tr As New PositionAdjustString("?unclosed")
        Dim expr As New SpecialSeqExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.False(r.Enable)
        Assert.Equal(startPos, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_NotStartingWithQuestion_RestoresPositionAndReturnsInvalid()
        Dim tr As New PositionAdjustString("x?y?")
        Dim expr As New SpecialSeqExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.False(r.Enable)
        Assert.Equal(startPos, tr.Position)
    End Sub

End Class