Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 式の範囲を表します。
    ''' </summary>
    Public Structure ExpressionRange

        ''' <summary>
        ''' 空の式リスト。
        ''' </summary>
        Private Shared ReadOnly _emptyRanges As New Lazy(Of IEnumerable(Of ExpressionRange))(
            Function()
                Return New ExpressionRange() {}
            End Function
        )

        ''' <summary>
        ''' 無効式。
        ''' </summary>
        Private Shared ReadOnly _invalid As New Lazy(Of ExpressionRange)(
            Function()
                Return New ExpressionRange(Nothing, Nothing, -1, -1, New List(Of ExpressionRange))
            End Function
        )

        ''' <summary>
        ''' 式のテキストリーダー。
        ''' </summary>
        Private ReadOnly _tr As IPositionAdjustReader

        ''' <summary>
        ''' 範囲の開始位置（0 ベースのインデックス）。
        ''' </summary>
        Public ReadOnly Property [Start] As Integer

        ''' <summary>
        ''' 範囲の終了位置（0 ベースのインデックス、開始位置より大きいと有効）。
        ''' </summary>
        Public ReadOnly Property [End] As Integer

        ''' <summary>
        ''' この範囲に対応する式。
        ''' </summary>
        Public ReadOnly Property Expr As IExpression

        ''' <summary>
        ''' サブレンジのリスト。
        ''' </summary>
        Public ReadOnly Property SubRanges As IEnumerable(Of ExpressionRange)

        ''' <summary>
        ''' 範囲が有効かを示します。開始位置が終了位置より小さい場合 True を返します。
        ''' </summary>
        Public ReadOnly Property Enable As Boolean
            Get
                Return Me.[Start] < Me.[End] AndAlso Me.[Start] >= 0
            End Get
        End Property

        ''' <summary>
        ''' 空の式リストを取得する。
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property EmptyRanges As IEnumerable(Of ExpressionRange)
            Get
                Return _emptyRanges.Value
            End Get
        End Property

        ''' <summary>
        ''' 無効な範囲を表す定数（StartPos = -1, EndPos = -1）。
        ''' </summary>
        Public Shared ReadOnly Property Invalid As ExpressionRange
            Get
                Return _invalid.Value
            End Get
        End Property

        ''' <summary>
        ''' 指定した開始位置と終了位置で新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="expr">この範囲に対応する式。</param>
        ''' <param name="tr">式のテキストリーダー。</param>
        ''' <param name="startPos">開始位置（0 ベース）。</param>
        ''' <param name="endPos">終了位置（0 ベース）。</param>
        ''' <param name="subRanges">サブレンジのリスト。</param>
        Public Sub New(expr As IExpression, tr As IPositionAdjustReader,
                       startPos As Integer, endPos As Integer, subRanges As IEnumerable(Of ExpressionRange))
            Me.Expr = expr
            Me._tr = tr
            Me.[Start] = startPos
            Me.[End] = endPos
            Me.SubRanges = subRanges
        End Sub

        ''' <summary>
        ''' このインスタンスの文字列表現を取得します。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Overrides Function ToString() As String
            If Me.Start >= 0 AndAlso Me.End >= 0 Then
                Return Me._tr.Substring(Me.[Start], Me.[End] - Me.[Start])
            Else
                Return ""
            End If
        End Function

    End Structure

End Namespace
