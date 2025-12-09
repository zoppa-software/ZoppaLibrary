Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' EBNF解析結果の評価を行うモジュール。
    ''' </summary>
    Public Module EBNFEvaluate

        ''' <summary>
        ''' EBNF解析結果を評価します。 
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expr"></param>
        ''' <param name="evaMethod"></param>
        ''' <returns></returns>
        Public Function Run(Of T)(expr As EBNFEnvironment,
                                  evaMethod As Func(Of EBNFAnalysisItem, IEnumerable(Of EBNFEvaluateAnswer), EBNFEvaluateAnswer)) As T
            Return CType(RunSubroutine(expr.Answer, evaMethod).Value, T)
        End Function

        ''' <summary>
        ''' EBNF解析結果を評価します。 
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expr"></param>
        ''' <param name="evaMethod"></param>
        ''' <returns></returns>
        Public Function Run(Of T)(expr As EBNFAnalysisItem,
                                  evaMethod As Func(Of EBNFAnalysisItem, IEnumerable(Of EBNFEvaluateAnswer), EBNFEvaluateAnswer)) As T
            Return CType(RunSubroutine(expr, evaMethod).Value, T)
        End Function

        ''' <summary>
        ''' サブルーチンを実行します。
        ''' </summary>
        ''' <param name="expr">評価対象の解析結果。</param>
        ''' <param name="evaMethod">評価メソッド。</param>
        ''' <returns>評価結果。</returns>
        Private Function RunSubroutine(expr As EBNFAnalysisItem,
                                       evaMethod As Func(Of EBNFAnalysisItem, IEnumerable(Of EBNFEvaluateAnswer), EBNFEvaluateAnswer)) As EBNFEvaluateAnswer
            Dim values As New List(Of EBNFEvaluateAnswer)
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

