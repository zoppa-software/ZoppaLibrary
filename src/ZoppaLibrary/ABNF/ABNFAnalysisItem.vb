Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 解析の範囲を表します。
    ''' </summary>
    Public NotInheritable Class ABNFAnalysisItem

        ''' <summary>
        ''' 範囲の識別子を取得します。
        ''' </summary>
        Public ReadOnly Property Identifier As String

        ''' <summary>
        ''' 範囲内の解析結果のリストを取得します。
        ''' </summary>
        Private _answers As List(Of ABNFAnalysisItem)

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
        Public ReadOnly Property SubRanges As IEnumerable(Of ABNFAnalysisItem)
            Get
                Return Me._answers
            End Get
        End Property

        ''' <summary>
        ''' 指定したインデックスまたは識別子に対応する範囲を取得します。
        ''' </summary>
        ''' <param name="index">インデックス。</param>
        ''' <returns>対応する範囲。</returns>
        Default Public ReadOnly Property Item(index As Integer) As ABNFAnalysisItem
            Get
                If index >= 0 AndAlso index < Me._answers.Count Then
                    Return Me._answers(index)
                Else
                    Throw New IndexOutOfRangeException($"インデックス '{index}' が範囲外です。")
                End If
            End Get
        End Property

        ''' <summary>
        ''' 指定した識別子に対応する範囲を取得します。
        ''' </summary>
        ''' <param name="ident">識別子。</param>
        ''' <returns>対応する範囲。</returns>
        Default Public ReadOnly Property Item(ident As String) As ABNFAnalysisItem
            Get
                For Each ans As ABNFAnalysisItem In Me._answers
                    If ans.Identifier = ident Then
                        Return ans
                    End If
                Next
                Throw New KeyNotFoundException($"識別子 '{ident}' の範囲が見つかりません。")
            End Get
        End Property

        ''' <summary>
        ''' 範囲内の解析結果の数を取得します。
        ''' </summary>
        Public ReadOnly Property Count As Integer
            Get
                Return Me._answers.Count
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
                       answers As List(Of ABNFAnalysisItem),
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
        ''' 指定した識別子に対応する範囲をすべて取得します。
        ''' </summary>
        ''' <param name="ident">識別子。</param>
        ''' <returns>対応する範囲のリスト。</returns>
        Public Function SearchByName(ident As String) As ABNFAnalysisItem
            Dim res As New List(Of ABNFAnalysisItem)
            For Each ans As ABNFAnalysisItem In Me._answers
                If ans.Identifier = ident Then
                    res.Add(ans)
                End If
            Next
            Return New ABNFAnalysisItem(ident, res, Me._tr, Me.Start, Me.End)
        End Function

        ''' <summary>
        ''' この範囲の文字列を取得します。
        ''' </summary>
        ''' <returns>範囲の文字列。</returns>
        Public Overrides Function ToString() As String
            Return $"{Me._tr.Substring(Me.Start, Me.End - Me.Start)}"
        End Function

    End Class

End Namespace
