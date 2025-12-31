Option Explicit On
Option Strict On

Imports System.Text
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ABNF解析ノード基底クラス。
    ''' </summary>
    ''' <remarks>
    ''' <para>このクラスは解析グラフのノードを表現します。</para>
    ''' <para>主要なサブクラス:</para>
    ''' <list type="bullet">
    ''' <item>EpsilonNode: ε遷移ノード（空遷移）</item>
    ''' <item>CharValNode: 文字リテラルのマッチング</item>
    ''' <item>NumValNode: 数値範囲のマッチング</item>
    ''' <item>RuleNameNode: ルール参照のマッチング</item>
    ''' <item>MethodNode: カスタムメソッドによるマッチング</item>
    ''' </list>
    ''' </remarks>
    Public MustInherit Class AnalysisNode

        ''' <summary>識別値。</summary>
        Public ReadOnly Property Id As Integer

        ''' <summary>評価範囲。</summary>
        Public MustOverride ReadOnly Property Range As ExpressionRange

        ''' <summary>接続ルート。</summary>
        Public ReadOnly Property Routes As List(Of Route)

        ''' <summary>
        ''' 再試行可能かを取得する。
        ''' </summary>
        Public MustOverride ReadOnly Property IsRetry As Boolean

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(id As Integer)
            Me.Id = id
            Me.Routes = New List(Of Route)()
        End Sub

        ''' <summary>
        ''' インスタンスを生成する。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="range">評価範囲。</param>
        ''' <returns>生成されたインスタンス。</returns>
        ''' <remarks>
        ''' 式の型に応じて以下のサブクラスを返します:
        ''' <list type="bullet">
        ''' <item><see cref="CharValExpression"/> → <see cref="CharValNode"/></item>
        ''' <item><see cref="NumValExpression"/> → <see cref="NumValNode"/></item>
        ''' <item><see cref="RuleNameExpression"/> → <see cref="RuleNameNode"/></item>
        ''' <item>その他 → <see cref="EpsilonNode"/> (空ノード)</item>
        ''' </list>
        ''' </remarks>
        Public Shared Function Create(id As Integer, range As ExpressionRange) As AnalysisNode
            Select Case range.Expr?.GetType()
                Case GetType(CharValExpression)
                    Return New CharValNode(id, range)
                Case GetType(NumValExpression)
                    Return New NumValNode(id, range)
                Case GetType(RuleNameExpression)
                    Return New RuleNameNode(id, range)
                Case Else
                    Return New EpsilonNode(id)
            End Select
        End Function

        ''' <summary>
        ''' インスタンスを生成する。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="name">名前。</param>
        ''' <param name="method">マッチ対象を判定する関数。</param>
        ''' <returns>生成されたインスタンス。</returns>
        Public Shared Function Create(id As Integer, name As String, method As Func(Of PositionAdjustBytes, Boolean)) As AnalysisNode
            Return New MethodNode(id, name, method)
        End Function

        ''' <summary>
        ''' キャッシュをクリアします。
        ''' </summary>
        ''' <param name="idHash">クリア済みノードIDハッシュセット。</param>
        Friend Sub ClearCache(idHash As HashSet(Of Integer))
            Me.ClearCacheImpl()
            For Each route In Me.Routes
                If Not idHash.Contains(route.NextNode.Id) Then
                    idHash.Add(route.NextNode.Id)
                    route.NextNode.ClearCache(idHash)
                End If
            Next
        End Sub

        ''' <summary>
        ''' キャッシュをクリアします。（実装）
        ''' </summary>
        Protected Overridable Sub ClearCacheImpl()
            ' 何もしない
        End Sub

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
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public MustOverride Function Match(tr As PositionAdjustBytes,
                                           env As ABNFEnvironment,
                                           ruleName As String) As (success As Boolean, answer As ABNFAnalysisItem)

        ''' <summary>
        ''' 次のパターンのマッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public MustOverride Function MoveNext(tr As PositionAdjustBytes,
                                              env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Dim buf As New StringBuilder()
            For Each n In Me.Routes
                If buf.Length > 0 Then buf.Append(", ")
                buf.Append($"{If(n.NextNode?.Id.ToString(), "")}({n.RequiredVisits},{n.LimitedVisits})")
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
