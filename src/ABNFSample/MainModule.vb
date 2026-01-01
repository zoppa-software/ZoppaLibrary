Option Explicit On
Option Strict On

Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.BNF

Module MainModule

    Sub Main()
        Dim env As ABNFSyntaxAnalysis.ABNFEnvironment
        Using sr As New IO.StreamReader("JSON_ABNF.txt")
            env = CompileEnvironment(New PositionAdjustStringReader(sr))
        End Using

        Dim input = "" &
"{
""name"": ""Tanaka"",
""age"": 30,
""isStudent"": false
}"
        Dim ans = env.Evaluate("JSON-text", New PositionAdjustBytes(Text.Encoding.UTF8.GetBytes(input)))
        ans.PrintAnalysisTree(Console.Out)
    End Sub

End Module
