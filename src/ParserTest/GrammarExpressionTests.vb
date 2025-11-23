Option Explicit On
Option Strict On

Imports System.Net
Imports Xunit
Imports ZoppaLibrary.Parser

Public Class GrammarExpressionTests

    <Fact>
    Public Sub Match_SingleRule_ReturnsMatchAndAdvancesReader()
        Dim input = "id = 'x';"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New GrammarExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_MultipleRules_ReturnsMatchAndAdvancesReader()
        Dim input = "first = 'a'; second = 'b'."
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New GrammarExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_RulesWithWhitespaceAndNewlines_ReturnsMatchAndAdvancesReader()
        Dim input = "  r1 = 'x' ;" & vbLf & "r2='y';  "
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New GrammarExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_PartialRule_MissingTerminator_ReturnsInvalidAndRestoresReader()
        Dim input = "name = 'val'"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New GrammarExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_GrammarTest()
        Dim input = "" &
"letter   = ""A"" | ""B"" | ""C"" | ""D"" | ""E"" | ""F"" | ""G""
       | ""H"" | ""I"" | ""J"" | ""K"" | ""L"" | ""M"" | ""N""
       | ""O"" | ""P"" | ""Q"" | ""R"" | ""S"" | ""T"" | ""U""
       | ""V"" | ""W"" | ""X"" | ""Y"" | ""Z"" | ""a"" | ""b""
       | ""c"" | ""d"" | ""e"" | ""f"" | ""g"" | ""h"" | ""i""
       | ""j"" | ""k"" | ""l"" | ""m"" | ""n"" | ""o"" | ""p""
       | ""q"" | ""r"" | ""s"" | ""t"" | ""u"" | ""v"" | ""w""
       | ""x"" | ""y"" | ""z"" ;

digit = ""0"" | ""1"" | ""2"" | ""3"" | ""4"" | ""5"" | ""6"" | ""7"" | ""8"" | ""9"" ;

symbol = ""["" | ""]"" | ""{"" | ""}"" | ""("" | "")"" | ""<"" | "">""
       | ""'"" | '""' | ""="" | ""|"" | ""."" | "","" | "";"" | ""-"" 
       | ""+"" | ""*"" | ""?"" | """ & vbLf & """ | """ & vbTab & """ | """ & vbCr & """ | """ & vbFormFeed & """ | """ & vbBack & """ ;

character = letter | digit | symbol | ""_"" | "" "" ;
identifier = letter , { letter | digit | ""_"" } ;

S = { "" "" | """ & vbLf & """ | """ & vbTab & """ | """ & vbCr & """ | """ & vbFormFeed & """ | """ & vbBack & """ } ;

terminal = ""'"" , character - ""'"" , { character - ""'"" } , ""'""
         | '""' , character - '""' , { character - '""' } , '""' ;

terminator = "";"" | ""."" ;

term = ""("" , S , rhs , S , "")""
     | ""["" , S , rhs , S , ""]""
     | ""{"" , S , rhs , S , ""}""
     | terminal
     | identifier ;

factor = term , S , ""?""
       | term , S , ""*""
       | term , S , ""+""
       | term , S , ""-"" , S , term
       | term , S ;

concatenation = ( S , factor , S , "","" ? ) + ;
alternation = ( S , concatenation , S , ""|"" ? ) + ;

rhs = alternation ;
lhs = identifier ;

rule = lhs , S , ""="" , S , rhs , S , terminator ;

grammar = ( S , rule , S ) * ;"
        Dim answer = SyntaxAnalysis.LexicalAnalysis(input, "grammar", input)
    End Sub

    <Fact>
    Public Sub Match_GrammarTest2()
        Dim input = "" &
"digit = ? local_number ?;
add_or_sub = digit, S, ('+' | '-'), S, digit;
S = { ' ' } ;
grammar = add_or_sub;"
        SyntaxAnalysis.AddSpecialMethods(
            "local_number",
            Function(tr As IPositionAdjustReader) As Boolean
                Dim startPos = tr.Position
                Dim readAny = False
                While Char.IsDigit(ChrW(tr.Peek()))
                    tr.Read()
                    readAny = True
                End While
                Return readAny
            End Function
        )

        Dim answer = SyntaxAnalysis.LexicalAnalysis(input, "grammar", "100 + 200")
    End Sub


    <Fact>
    Public Sub NumberTest1()
        Dim input = "number = ? Number ?;"
        Dim ans1 = SyntaxAnalysis.LexicalAnalysis(input, "number", "+1.0")
        Assert.Equal("+1.0", ans1.ToString())
        Dim ans2 = SyntaxAnalysis.LexicalAnalysis("number", "3.1415")
        Assert.Equal("3.1415", ans2.ToString())
        Dim ans3 = SyntaxAnalysis.LexicalAnalysis("number", "-0.01")
        Assert.Equal("-0.01", ans3.ToString())
        Dim ans4 = SyntaxAnalysis.LexicalAnalysis("number", "5e+22")
        Assert.Equal("5e+22", ans4.ToString())
        Dim ans5 = SyntaxAnalysis.LexicalAnalysis("number", "1e06")
        Assert.Equal("1e06", ans5.ToString())
        Dim ans6 = SyntaxAnalysis.LexicalAnalysis("number", "-2E-2")
        Assert.Equal("-2E-2", ans6.ToString())
        Dim ans7 = SyntaxAnalysis.LexicalAnalysis("number", "6.626e-34")
        Assert.Equal("6.626e-34", ans7.ToString())

        Assert.Throws(Of ArgumentException)(
            Sub()
                SyntaxAnalysis.LexicalAnalysis("number", ".7")
            End Sub
        )
        Assert.Throws(Of ArgumentException)(
            Sub()
                SyntaxAnalysis.LexicalAnalysis("number", "7.")
            End Sub
        )
        Assert.Throws(Of ArgumentException)(
            Sub()
                SyntaxAnalysis.LexicalAnalysis("number", "3.e+20")
            End Sub
        )
    End Sub

    <Fact>
    Public Sub NumberTest2()
        Dim input = "" &
"number = ? Number ?;
backet = '(' , S , add_or_sub , S , ')' ;
term = number | backet ;
multi_or_div = term, {S, ('*' | '/'), S, term};
add_or_sub = multi_or_div, {S, ('+' | '-'), S, multi_or_div};
S = {? Space ?};
grammar = add_or_sub;"

        Dim analysised1 = SyntaxAnalysis.LexicalAnalysis(input, "grammar", "1 + 2 * (3 - 4 + 5)")
        Dim answer1 = ExpressionEvaluate.Run(Of Integer)(analysised1, AddressOf Evaluate2)
        Assert.Equal(9, answer1)
    End Sub

    Private Function Evaluate2(expr As AnalysisRange, values As IEnumerable(Of EvaluateAnswer)) As EvaluateAnswer
        Dim filtered = values.Where(Function(v) v.Range IsNot Nothing AndAlso v.Range.Identifier <> "S").ToList()

        Select Case expr.Identifier
            Case "Number", "Space", "S"
                ' 評価しない
                Return Nothing
            Case "number"
                Return New EvaluateAnswer(expr, Double.Parse(expr.SubRanges(0).ToString()))
            Case "backet"
                Return New EvaluateAnswer(expr, filtered(1).Value)
            Case "term"
                Return New EvaluateAnswer(expr, filtered(0).Value)

            Case "literal"
                Return New EvaluateAnswer(expr, expr.ToString())

            Case "add_or_sub"
                Dim asans = CDbl(filtered(0).Value)
                For i As Integer = 1 To filtered.Count - 1 Step 2
                    Dim op = filtered(i).Value.ToString()
                    Dim right = CDbl(filtered(i + 1).Value)
                    Select Case op
                        Case "+"
                            asans += right
                        Case "-"
                            asans -= right
                    End Select
                Next
                Return New EvaluateAnswer(expr, asans)

            Case "multi_or_div"
                Dim mdans = CDbl(filtered(0).Value)
                For i As Integer = 1 To filtered.Count - 1 Step 2
                    Dim op = filtered(i).Value.ToString()
                    Dim right = CDbl(filtered(i + 1).Value)
                    Select Case op
                        Case "*"
                            mdans *= right
                        Case "/"
                            mdans /= right
                    End Select
                Next
                Return New EvaluateAnswer(expr, mdans)

            Case "grammar"
                Return New EvaluateAnswer(expr, filtered(0).Value)

            Case Else
                Throw New InvalidOperationException($"未知の識別子: {expr.Identifier}")
        End Select
    End Function

End Class