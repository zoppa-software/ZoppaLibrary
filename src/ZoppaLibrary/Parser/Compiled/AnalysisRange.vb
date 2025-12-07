Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 解析の範囲を表します。
    ''' </summary>
    Public NotInheritable Class AnalysisRange

        ''' <summary>
        ''' 範囲の識別子を取得します。
        ''' </summary>
        Public ReadOnly Property Identifier As String

        ''' <summary>
        ''' 範囲内の解析結果のリストを取得します。
        ''' </summary>
        Private _answers As List(Of AnalysisRange)

        ''' <summary>
        ''' 位置調整リーダーを取得します。
        ''' </summary>
        Private _tr As IPositionAdjustReader

        ''' <summary>
        ''' 範囲の開始位置（0 ベースのインデックス）を取得します。
        ''' </summary>
        Public ReadOnly Property [Start] As Integer

        ''' <summary>
        ''' 範囲の終了位置（0 ベースのインデックス、開始位置より大きいと有効）を取得します。
        ''' </summary>
        Public ReadOnly Property [End] As Integer

        ''' <summary>
        ''' 範囲内のサブレンジのリストを取得します。
        ''' </summary>
        Public ReadOnly Property SubRanges As IEnumerable(Of AnalysisRange)
            Get
                Return Me._answers
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="ident">範囲の識別子。</param>
        ''' <param name="answers">範囲内の解析結果のリスト。</param>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="startPos">範囲の開始位置。</param>
        ''' <param name="endPos">範囲の終了位置。</param>
        Public Sub New(ident As String,
                       answers As List(Of AnalysisRange),
                       tr As IPositionAdjustReader,
                       startPos As Integer,
                       endPos As Integer)
            Me.Identifier = ident
            Me._answers = answers
            Me._tr = tr
            Me.Start = startPos
            Me.End = endPos
        End Sub

        ''' <summary>
        ''' この範囲の文字列を取得します。
        ''' </summary>
        ''' <returns>範囲の文字列。</returns>
        Public Overrides Function ToString() As String
            Return $"{Me._tr.Substring(Me.Start, Me.End - Me.Start)}"
        End Function

    End Class

End Namespace
