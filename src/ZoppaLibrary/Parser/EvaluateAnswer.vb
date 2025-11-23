Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 評価結果を表す構造体。
    ''' </summary>
    Public NotInheritable Class EvaluateAnswer

        ''' <summary>
        ''' 評価範囲を取得する。
        ''' </summary>
        ''' <returns>評価範囲。</returns>
        Public ReadOnly Property Range As AnalysisRange

        ''' <summary>
        ''' 評価値を取得する。
        ''' </summary>
        ''' <returns>評価値。</returns>
        Public ReadOnly Property Value As Object

        ''' <summary>
        ''' 評価値の型を取得する。
        ''' </summary>
        ''' <returns>評価値の型。</returns>
        Public ReadOnly Property ValueType As Type
            Get
                If Value Is Nothing Then
                    Return GetType(Object)
                Else
                    Return Value.GetType()
                End If
            End Get
        End Property

        ''' <summary>
        ''' 新しい EvaluateAnswer のインスタンスを初期化する。
        ''' </summary>
        ''' <param name="rng">評価範囲。</param>
        ''' <param name="value">評価値。</param>
        Public Sub New(rng As AnalysisRange, value As Object)
            Me.Range = rng
            Me.Value = value
        End Sub

    End Class

End Namespace
