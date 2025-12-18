Option Explicit On
Option Strict On

Imports System.Net
Imports Xunit
Imports ZoppaLibrary.EBNF

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
        Dim answer = EBNFSyntaxAnalysis.CompileToEvaluate(input, "grammar", input)
    End Sub

    <Fact>
    Public Sub Match_GrammarTest2()
        Dim input = "" &
"digit = ? local_number ?;
add_or_sub = digit, { S, ('+' | '-'), S, digit };
S = { ' ' } ;
grammar = add_or_sub;"
        Dim answer = EBNFSyntaxAnalysis.CompileToEvaluate(
            input,
            Sub(env)
                env.Add(
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
            End Sub,
            "grammar",
            "10 + 20 - 5"
        )
    End Sub


    <Fact>
    Public Sub NumberTest1()
        Dim input = "number = ? Number ?;"
        Dim env = EBNFSyntaxAnalysis.CompileToEvaluate(input, "number", "+1.0")
        Assert.Equal("+1.0", env.Answer.ToString())
        Dim ans2 = env.Evaluate("number", "3.1415")
        Assert.Equal("3.1415", ans2.ToString())
        Dim ans3 = env.Evaluate("number", "-0.01")
        Assert.Equal("-0.01", ans3.ToString())
        Dim ans4 = env.Evaluate("number", "5e+22")
        Assert.Equal("5e+22", ans4.ToString())
        Dim ans5 = env.Evaluate("number", "1e06")
        Assert.Equal("1e06", ans5.ToString())
        Dim ans6 = env.Evaluate("number", "-2E-2")
        Assert.Equal("-2E-2", ans6.ToString())
        Dim ans7 = env.Evaluate("number", "6.626e-34")
        Assert.Equal("6.626e-34", ans7.ToString())

        Assert.Throws(Of ArgumentException)(
            Sub()
                env.Evaluate("number", ".7")
            End Sub
        )
        Assert.Throws(Of ArgumentException)(
            Sub()
                env.Evaluate("number", "7.")
            End Sub
        )
        Assert.Throws(Of ArgumentException)(
            Sub()
                env.Evaluate("number", "3.e+20")
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

        Dim analysised1 = EBNFSyntaxAnalysis.CompileToEvaluate(input, "grammar", "1 + 2 * (3 - 4 + 5)")
        Dim answer1 = EBNFEvaluate.Run(Of Integer)(analysised1, AddressOf Evaluate2)
        Assert.Equal(9, answer1)
    End Sub

    Private Function Evaluate2(expr As EBNFAnalysisItem, values As IEnumerable(Of EBNFEvaluateAnswer)) As EBNFEvaluateAnswer
        Dim filtered = values.Where(Function(v) v.Range IsNot Nothing AndAlso v.Range.Identifier <> "S").ToList()

        Select Case expr.Identifier
            Case "Number", "Space", "S"
                ' 評価しない
                Return Nothing
            Case "number"
                Return New EBNFEvaluateAnswer(expr, Double.Parse(expr.SubRanges(0).ToString()))
            Case "backet"
                Return New EBNFEvaluateAnswer(expr, filtered(1).Value)
            Case "term"
                Return New EBNFEvaluateAnswer(expr, filtered(0).Value)

            Case "literal"
                Return New EBNFEvaluateAnswer(expr, expr.ToString())

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
                Return New EBNFEvaluateAnswer(expr, asans)

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
                Return New EBNFEvaluateAnswer(expr, mdans)

            Case "grammar"
                Return New EBNFEvaluateAnswer(expr, filtered(0).Value)

            Case Else
                Throw New InvalidOperationException($"未知の識別子: {expr.Identifier}")
        End Select
    End Function

    <Fact>
    Public Sub StringTest()
        Dim input = "" &
"S = {? Space ?};
quo_blk = ""'"", (? AllChar ? - ""'"")+, ""'"";
add_or_sub = quo_blk, {S, '+', S, quo_blk};
grammar = add_or_sub;"

        Dim analysised1 = EBNFSyntaxAnalysis.CompileToEvaluate(input, "grammar", "'あいう' + 'えお'")
        Dim answer1 = EBNFEvaluate.Run(Of String)(analysised1, AddressOf Evaluate3)
        Assert.Equal("あいうえお", answer1)
    End Sub

    Private Function Evaluate3(expr As EBNFAnalysisItem, values As IEnumerable(Of EBNFEvaluateAnswer)) As EBNFEvaluateAnswer
        Dim filtered = values.Where(Function(v) v.Range IsNot Nothing AndAlso v.Range.Identifier <> "S").ToList()

        Select Case expr.Identifier
            Case "AllChar", "Space", "S"
                ' 評価しない
                Return Nothing
            Case "quo_blk"
                Return New EBNFEvaluateAnswer(expr, expr.ToString().Trim("'"c))
            Case "literal"
                Return New EBNFEvaluateAnswer(expr, expr.ToString())
            Case "add_or_sub"
                Dim asans = filtered(0).Value.ToString()
                For i As Integer = 1 To filtered.Count - 1 Step 2
                    Dim op = filtered(i).Value.ToString()
                    Dim right = filtered(i + 1).Value.ToString()
                    Select Case op
                        Case "+"
                            asans &= right
                    End Select
                Next
                Return New EBNFEvaluateAnswer(expr, asans)
            Case "grammar"
                Return New EBNFEvaluateAnswer(expr, filtered(0).Value)

            Case Else
                Throw New InvalidOperationException($"未知の識別子: {expr.Identifier}")
        End Select
    End Function

    <Fact>
    Public Sub NumberTest3()
        Dim input = "" &
"multi_or_div = term, {S, ('*' | '/'), S, term};
number = ? Number ?;
backet = '(' , S , add_or_sub , S , ')' ;
term = number | backet ;
add_or_sub = multi_or_div, {S, ('+' | '-'), S, multi_or_div};
S = {? Space ?};
grammar = add_or_sub;"

        'Dim expr = New GrammarExpression()
        'Dim range = expr.Match(New PositionAdjustStringReader(input))
        'CreateRuleTable(range)

        Dim analysised1 = EBNFSyntaxAnalysis.CompileToEvaluate(input, "grammar", "1 + 2 * (3 - 4 + 5)")
    End Sub

End Class