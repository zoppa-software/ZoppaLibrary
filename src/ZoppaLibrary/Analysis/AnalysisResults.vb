Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' ParseAnswerクラスは、解析結果を表すクラスです。
    ''' このクラスは、解析されたデータや結果を格納するために使用されます。
    ''' </summary>
    Public NotInheritable Class AnalysisResults
        Implements IComparable(Of AnalysisResults)

        ' 解析した文字列

        ''' <summary>
        ''' 解析した文字列を取得します。
        ''' このプロパティは、解析された文字列を表すU8String型の値を返します。
        ''' 解析結果の文字列は、後で参照するために使用されます。
        ''' </summary>
        ''' <returns>解析した文字列。</returns>
        Public ReadOnly Property InputString As U8String

        ''' <summary>
        ''' 解析結果の式を取得します。
        ''' このプロパティは、解析された式を表すIExpression型の値を返します。
        ''' </summary>
        ''' <returns>解析結果の式。</returns>
        Public ReadOnly Property Expression As IExpression

        ' 解析結果

        ’'' <summary>コンストラクタ。</summary>
        ''' <param name="input">解析した文字列。</param>
        ''' <param name="expression">解析結果の式。</param>
        ''' <remarks>
        ''' このコンストラクタは、解析された文字列とその結果の式を初期化します。
        ''' </remarks>
        Public Sub New(input As U8String, expression As IExpression)
            Me.InputString = input
            Me.Expression = expression
        End Sub

        ''' <summary>比較を行います。</summary>
        ''' <param name="other">比較対象。</param>
        ''' <returns>比較結果。</returns>
        Public Function CompareTo(other As AnalysisResults) As Integer Implements IComparable(Of AnalysisResults).CompareTo
            Return Me.InputString.CompareTo(other.InputString)
        End Function

        ''' <summary>文字列表現を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"{InputString}"
        End Function

    End Class

End Namespace
