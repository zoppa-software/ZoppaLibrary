Option Explicit On
Option Strict On

Imports System.Text
Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.ABNF.ABNFSyntaxAnalysis
Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' EBNF解析ノード基底クラス。
    ''' </summary>
    ''' <remarks>
    ''' <para>このクラスは解析グラフのノードを表現します。</para>
    ''' <para>主要なサブクラス:</para>
    ''' <list type="bullet">
    ''' <item><see cref="ExclusiveNode"/>: 要素式ノード。</item>
    ''' <item><see cref="SpecialSeqNode"/>: 特殊シーケンスノード。</item>
    ''' <item><see cref="TerminalNode"/>: 終端記号ノード。</item>
    ''' <item><see cref="IdentifierNode"/>: 識別子ノード。</item>
    ''' <item><see cref="EpsilonNode"/>: 空ノード（ε遷移）。</item>
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
        ''' <item><see cref="FactorExpression"/> → <see cref="ExclusiveNode"/></item>
        ''' <item><see cref="SpecialSeqExpression"/> → <see cref="SpecialSeqNode"/></item>
        ''' <item><see cref="TerminalExpression"/> → <see cref="TerminalNode"/></item>
        ''' <item><see cref="IdentifierExpression"/> → <see cref="IdentifierNode"/></item>
        ''' <item>その他 → <see cref="EpsilonNode"/> (空ノード)</item>
        ''' </list>
        ''' </remarks>
        Public Shared Function Create(id As Integer, range As ExpressionRange) As AnalysisNode
            Select Case range.Expr?.GetType()
                Case GetType(SpecialSeqExpression) ' 特殊式
                    Return New SpecialSeqNode(id, range)
                Case GetType(TerminalExpression) ' 終端記号
                    Return New TerminalNode(id, range)
                    'Return New NumValNode(id, range)
                Case GetType(IdentifierExpression) ' 識別子式
                    Return New IdentifierNode(id, range)
                    'Return New RuleNameNode(id, range)
                Case GetType(FactorExpression) ' 要素式
                    Return New ExclusiveNode(id, range)
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
        Public Shared Function Create(id As Integer, name As String, method As Func(Of IPositionAdjustReader, Boolean)) As AnalysisNode
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
        Public Sub AddRoute(nextNode As AnalysisNode)
            Me.Routes.Add(New Route(nextNode))
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public MustOverride Function Match(tr As IPositionAdjustReader,
                                           env As EBNFEnvironment,
                                           ruleName As String) As (success As Boolean, answer As EBNFAnalysisItem)

        ''' <summary>
        ''' 次のパターンのマッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public MustOverride Function MoveNext(tr As IPositionAdjustReader,
                                              env As EBNFEnvironment) As (success As Boolean, answer As EBNFAnalysisItem)

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Dim buf As New StringBuilder()
            For Each n In Me.Routes
                If buf.Length > 0 Then buf.Append(", ")
                buf.Append($"{If(n.NextNode?.Id.ToString(), "")}")
            Next
            Return $"{Me.Id} {Me.Range} -> {buf}"
        End Function

        ''' <summary>接続ルート情報。</summary>
        Public Structure Route

            ''' <summary>次のノード。</summary>
            Public ReadOnly Property NextNode As AnalysisNode

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="nextNode">次のノード。</param>
            Public Sub New(nextNode As AnalysisNode)
                Me.NextNode = nextNode
            End Sub

        End Structure
    End Class

End Namespace
