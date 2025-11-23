Option Explicit On
Option Strict On

Imports ZoppaLibrary.Parser
Imports Xunit

Public Class CommentExpressionTests

    <Fact>
    Public Sub Match_ValidComment_ReturnsRange()
        Dim tr As New PositionAdjustString("(* comment *)rest")
        Dim expr As New CommentExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(startPos, r.Start)
        Assert.Equal(tr.Position, r.End)
        Assert.Equal("(* comment *)", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_EmptyComment_ReturnsRange()
        Dim tr As New PositionAdjustString("(**)xyz")
        Dim expr As New CommentExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal("(**)", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_NoClosing_RestoresPositionAndReturnsInvalid()
        Dim tr As New PositionAdjustString("(* unclosed")
        Dim expr As New CommentExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.False(r.Enable)
        Assert.Equal(startPos, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_NotAComment_RestoresPositionAndReturnsInvalid()
        Dim tr As New PositionAdjustString("(a)rest")
        Dim expr As New CommentExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.False(r.Enable)
        Assert.Equal(startPos, tr.Position)
    End Sub

End Class