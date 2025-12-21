Option Explicit On
Option Strict On

Imports ZoppaLibrary.ABNF
Imports Xunit
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' <see cref="SpaceExpression"/> のテストクラス。
    ''' </summary>
    Public Class SpaceExpressionTests

        <Fact>
        Public Sub Match_SingleSpace_ReturnsRangeAndAdvancesReader()
            Dim tr = New PositionAdjustStringReader(" ")
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(1, r.[End])
            Assert.Equal(1, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_SingleTab_ReturnsRangeAndAdvancesReader()
            Dim tr = New PositionAdjustStringReader(vbTab)
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(1, r.[End])
            Assert.Equal(1, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_MultipleSpaces_ReturnsRangeAndAdvancesReader()
            Dim tr = New PositionAdjustStringReader("   ")
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(3, r.[End])
            Assert.Equal(3, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_MixedSpacesAndTabs_ReturnsRangeAndAdvancesReader()
            Dim tr = New PositionAdjustStringReader("  " & vbTab & " " & vbTab)
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(5, r.[End])
            Assert.Equal(5, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_SpacesFollowedByNonSpace_ReturnsRangeUpToNonSpace()
            Dim tr = New PositionAdjustStringReader("  abc")
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(2, r.[End])
            Assert.Equal(2, tr.Position)
            Assert.Equal(AscW("a"c), tr.Peek())
        End Sub

        <Fact>
        Public Sub Match_EmptyString_ReturnsInvalid()
            Dim tr = New PositionAdjustStringReader(String.Empty)
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.False(r.Enable)
            Assert.Equal(0, tr.Position)
        End Sub

        <Theory>
        <InlineData("a")>
        <InlineData("0")>
        <InlineData("_")>
        <InlineData("|")>
        <InlineData("abc")>
        Public Sub Match_NonSpaceCharacters_ReturnsInvalidAndDoesNotAdvance(input As String)
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.False(r.Enable)
            Assert.Equal(0, tr.Position)
        End Sub

        <Theory>
        <InlineData(vbLf)>
        <InlineData(vbCr)>
        <InlineData(vbFormFeed)>
        <InlineData(vbBack)>
        Public Sub Match_OtherWhitespaceCharacters_ReturnsInvalidAndDoesNotAdvance(input As String)
            ' ABNF の WSP 定義では SP と HTAB のみが対象
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.False(r.Enable)
            Assert.Equal(0, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_AtEndOfInput_ReturnsInvalid()
            Dim tr = New PositionAdjustStringReader("a")
            tr.Read() ' 'a' を読み進めてEOFに到達
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.False(r.Enable)
            Assert.Equal(1, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_AfterNonSpace_StartsFromCurrentPosition()
            Dim tr = New PositionAdjustStringReader("abc  ")
            tr.Read() ' 'a'
            tr.Read() ' 'b'
            tr.Read() ' 'c'
            Dim expr = New SpaceExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(3, r.[Start])
            Assert.Equal(5, r.[End])
            Assert.Equal(5, tr.Position)
        End Sub

    End Class

End Namespace