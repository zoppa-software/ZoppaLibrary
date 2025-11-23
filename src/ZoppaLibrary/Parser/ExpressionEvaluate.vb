Option Explicit On
Option Strict On

Namespace Parser

    Public Module ExpressionEvaluate

        Public Function Run(Of T)(expr As AnalysisRange,
                                  evaMethod As Func(Of AnalysisRange, IEnumerable(Of EvaluateAnswer), EvaluateAnswer)) As T
            Return CType(RunSubroutine(expr, evaMethod).Value, T)
        End Function

        Private Function RunSubroutine(expr As AnalysisRange,
                                       evaMethod As Func(Of AnalysisRange, IEnumerable(Of EvaluateAnswer), EvaluateAnswer)) As EvaluateAnswer
            Dim values As New List(Of EvaluateAnswer)
            For Each sexpr In expr.SubRanges
                Dim seva = RunSubroutine(sexpr, evaMethod)
                If seva IsNot Nothing Then
                    values.Add(seva)
                End If
            Next
            Return evaMethod(expr, values)
        End Function

    End Module

End Namespace

