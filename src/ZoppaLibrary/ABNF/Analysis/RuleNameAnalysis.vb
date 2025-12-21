Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    NotInheritable Class RuleNameAnalysis
        Implements IAnalysis

        ''' <summary>識別子名。</summary>
        Private ReadOnly _name As String

        ''' <summary>評価範囲。</summary>
        Private ReadOnly _range As ExpressionRange

        Public ReadOnly Property Pattern As List(Of IAnalysis.Link) Implements IAnalysis.Pattern

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(range As ExpressionRange)
            Me._name = range.ToString()
            Me._range = range
            Me.Pattern = New List(Of IAnalysis.Link)()
        End Sub

        Public Function Match(tr As IPositionAdjustReader, env As ABNFEnvironment, ruleTable As SortedDictionary(Of String, RuleAnalysis), ruleName As String, answers As List(Of ABNFAnalysisItem)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            Throw New NotImplementedException()
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"<{Me._name}>"
        End Function

    End Class

End Namespace
