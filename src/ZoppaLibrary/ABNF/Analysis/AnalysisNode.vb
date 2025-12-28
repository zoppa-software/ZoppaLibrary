Option Explicit On
Option Strict On

Imports System.Text
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ABNF解析ノード。
    ''' </summary>
    Public Class AnalysisNode

        ''' <summary>識別値。</summary>
        Public ReadOnly Property Id As Integer

        ''' <summary>評価範囲。</summary>
        Public ReadOnly Property Range As ExpressionRange

        ''' <summary>接続ルート。</summary>
        Public ReadOnly Property Routes As List(Of Route)

        ''' <summary>
        ''' 再試行可能かを取得する。
        ''' </summary>
        Public Overridable ReadOnly Property IsRetry As Boolean
            Get
                Return False
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(id As Integer, range As ExpressionRange)
            Me.Id = id
            Me.Range = range
            Me.Routes = New List(Of Route)()
        End Sub

        ''' <summary>
        ''' インスタンスを生成する。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="range">評価範囲。</param>
        ''' <returns>生成されたインスタンス。</returns>
        Public Shared Function Create(id As Integer, range As ExpressionRange) As AnalysisNode
            Select Case range.Expr?.GetType()
                Case GetType(CharValExpression)
                    Return New CharValNode(id, range)
                Case GetType(NumValExpression)
                    Return New NumValNode(id, range)
                Case GetType(RuleNameExpression)
                    Return New RuleNameNode(id, range)
                Case Else
                    Return New AnalysisNode(id, range)
            End Select
        End Function

        ''' <summary>
        ''' ルートを追加する。
        ''' </summary>
        ''' <param name="nextNode">次のノード。</param>
        ''' <param name="required">必要訪問回数。</param>
        ''' <param name="limited">制限訪問回数。</param>
        Public Sub AddRoute(nextNode As AnalysisNode,
                            required As Integer,
                            limited As Integer)
            Me.Routes.Add(New Route(nextNode, required, limited))
        End Sub

        ''' <summary>
        ''' 次のルートが存在するかどうかを取得する。
        ''' </summary>
        ''' <param name="position">現在の位置。</param>
        ''' <param name="route">ルート番号。</param>
        ''' <returns>次のルートが存在する場合に True を返します。</returns>
        Public Function HasNext(position As Integer, route As Integer) As Boolean
            Return route < Me.Routes.Count
        End Function

        ''' <summary>
        ''' 次のルートが存在するかどうかを取得する。
        ''' </summary>
        ''' <param name="position">現在の位置。</param>
        ''' <param name="route">ルート番号。</param>
        ''' <returns>次のルートが存在する場合に True を返します。</returns>
        Public Overridable Function Match(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()

            ' 解析を実行
            If Me.Range.Expr Is Nothing Then
                ' 空文字列
                Return (True, Nothing)
            Else
                Throw New NotImplementedException()
            End If

            snapPos.Restore()
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param
        ''' <returns>マッチ結果。</returns>
        Public Overridable Function MoveNext(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, isRetry As Boolean, answer As ABNFAnalysisItem)
            Return (False, False, Nothing)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Dim buf As New StringBuilder()
            For Each n In Me.Routes
                If buf.Length > 0 Then buf.Append(", ")
                buf.Append($"{n.NextNode.Id}({n.RequiredVisits},{n.LimitedVisits})")
            Next
            Return $"{Me.Id} {Me.Range} -> {buf}"
        End Function

        ''' <summary>接続ルート情報。</summary>
        Public Structure Route

            ''' <summary>次のノード。</summary>
            Public ReadOnly Property NextNode As AnalysisNode

            ''' <summary>必要訪問回数。</summary>
            Public ReadOnly Property RequiredVisits As Integer

            ''' <summary>制限訪問回数。</summary>
            Public ReadOnly Property LimitedVisits As Integer

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="nextNode">次のノード。</param>
            ''' <param name="required">必要訪問回数。</param>
            ''' <param name="limited">制限訪問回数。</param>
            Public Sub New(nextNode As AnalysisNode, required As Integer, limited As Integer)
                Me.NextNode = nextNode
                Me.RequiredVisits = required
                Me.LimitedVisits = limited
            End Sub
        End Structure

    End Class

End Namespace
