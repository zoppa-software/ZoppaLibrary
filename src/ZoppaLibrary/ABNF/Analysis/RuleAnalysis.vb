Option Explicit On
Option Strict On

Imports System.Security.Cryptography
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ルールのコンパイル済み式を表します。
    ''' </summary>
    Public NotInheritable Class RuleAnalysis

        ''' <summary>
        ''' 解析ノードのルート。
        ''' </summary>
        Private ReadOnly _root As AnalysisNode

        ''' <summary>
        ''' ルール名を取得する。
        ''' </summary>
        ''' <returns>ルール名。</returns>
        Public ReadOnly Property RuleName As String

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="name">ルール名。</param>
        ''' <param name="targets">ルールのパターンを表す <see cref="ExpressionRange"/>。</param>
        Public Sub New(name As String, targets As ExpressionRange)
            Me.RuleName = name

            ' ルートを作成
            Dim nodes As New NodeList()
            Dim startNode = nodes.NewNode()
            Dim routes = CreateRoute(nodes, targets)
            Dim endNode = nodes.NewNode()

            startNode.Routes.Add(routes.st)
            routes.ed.Routes.Add(endNode)

            ' ノードのリンクを作成
            Dim pattern As New SortedDictionary(Of Integer, NodeLink)()
            CreatePattern(pattern, startNode, endNode)
            For i As Integer = 1 To nodes.Count - 2
                If Not nodes(i).IsEpsilon OrElse Not nodes(i).IsEpsilon Then
                    CreatePattern(pattern, nodes(i), endNode)
                End If
            Next
            pattern.Add(endNode.Id, New NodeLink(endNode))

            ' 評価用グラフを作成
            ' 1. 評価ノードを作成
            ' 2. ルートを接続
            ' 3. ルートを並び変え
            Dim analysis As New SortedDictionary(Of Integer, AnalysisNode)() ' 1
            For Each kvp In pattern
                With kvp.Value.StartNode
                    Dim ana = AnalysisNode.Create(.Id, .Range)
                    analysis.Add(ana.Id, ana)
                End With
            Next
            For Each kvp In pattern ' 2
                For Each edge In kvp.Value.EndNodes
                    If analysis.ContainsKey(edge.Item1.Id) Then
                        analysis(kvp.Key).AddRoute(analysis(edge.Item1.Id), edge.Item2, edge.Item3)
                    End If
                Next
            Next
            For Each kvp In analysis ' 3
                kvp.Value.MoveEndRoute(endNode.Id)
            Next
            Me._root = analysis(startNode.Id)
        End Sub

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="name">ルール名。</param>
        ''' <param name="method">マッチ対象を判定する関数。</param>
        Public Sub New(name As String, method As Func(Of PositionAdjustBytes, Boolean))
            Me.RuleName = name

            ' ルートを作成
            Dim startNode = AnalysisNode.Create(0, ExpressionRange.Invalid)
            Dim methodNode = AnalysisNode.Create(1, name, method)
            Dim endNode = AnalysisNode.Create(2, ExpressionRange.Invalid)
            startNode.AddRoute(methodNode, 0, Integer.MaxValue)
            methodNode.AddRoute(endNode, 0, Integer.MaxValue)

            Me._root = startNode
        End Sub

#Region "ルート作成"

        ''' <summary>
        ''' ルートを作成します。
        ''' </summary>
        ''' <param name="nodes">ノードリスト。</param>
        ''' <param name="target">式の範囲。</param>
        ''' <returns>接続点。</returns>
        Private Shared Function CreateRoute(nodes As NodeList, target As ExpressionRange) As (st As Node, ed As Node)
            Select Case target.Expr.GetType()
                Case GetType(AlternationExpression)
                    ' 選択式
                    Return If(target.SubRanges.Count > 1,
                              AlternationRoute(nodes, target),
                              CreateRoute(nodes, target.SubRanges(0)))

                Case GetType(CharValExpression)
                    ' 文字式
                    Return DirectRoute(nodes, target)

                Case GetType(ConcatenationExpression)
                    ' 連結式
                    Return If(target.SubRanges.Count > 1,
                              ConcatenationRoute(nodes, target),
                              CreateRoute(nodes, target.SubRanges(0)))

                Case GetType(GroupExpression)
                    ' グループ式
                    Return CreateRoute(nodes, target.SubRanges(0))

                Case GetType(OptionExpression)
                    ' オプション式
                    Return RangeRoute(nodes, target.SubRanges(0), 0, 1)

                Case GetType(RepetitionExpression)
                    ' 反復式
                    If target.SubRanges.Count > 1 Then
                        Dim minRange = target.SubRanges(0).SubRanges(0)
                        Dim maxRange = target.SubRanges(0).SubRanges(1)

                        Dim minCount = If(minRange.Enable, Integer.Parse(minRange.ToString()), 0)
                        Dim maxCount = If(maxRange.Enable, Integer.Parse(maxRange.ToString()), Integer.MaxValue)

                        Return RangeRoute(nodes, target.SubRanges(1), minCount, maxCount)
                    Else
                        Return CreateRoute(nodes, target.SubRanges(0))
                    End If

                Case GetType(RuleNameExpression)
                    ' ルール名式
                    Return DirectRoute(nodes, target)

                Case GetType(ProseValExpression)
                    ' 散文式
                    Return DirectRoute(nodes, target)

                Case GetType(NumValExpression)
                    ' 数値式
                    Return DirectRoute(nodes, target)

                Case Else
                    Throw New Exception("未知の式タイプです。")
            End Select
        End Function

        ''' <summary>
        ''' 単純比較ルートを作成します。
        ''' </summary>
        ''' <param name="nodes">ノードリスト。</param>
        ''' <param name="target">式の範囲。</param>
        ''' <returns>接続点。</returns>
        Private Shared Function DirectRoute(nodes As NodeList, target As ExpressionRange) As (st As Node, ed As Node)
            Dim startNode = nodes.NewNode(target)
            Dim endNode = nodes.NewNode()

            ' 一致接続
            startNode.Routes.Add(endNode)

            Return (startNode, endNode)
        End Function

        ''' <summary>
        ''' 選択ルートを作成します。
        ''' </summary>
        ''' <param name="nodes">ノードリスト。</param>
        ''' <param name="target">式の範囲。</param>
        ''' <returns>接続点。</returns>
        Private Shared Function AlternationRoute(nodes As NodeList, target As ExpressionRange) As (st As Node, ed As Node)
            Dim startNode = nodes.NewNode()
            Dim endNode = nodes.NewNode()

            ' 開始点と終了点の間に選択肢を接続
            For Each subRange In target.SubRanges
                Dim subRoute = CreateRoute(nodes, subRange)
                startNode.Routes.Add(subRoute.st)
                subRoute.ed.Routes.Add(endNode)
            Next
            Return (startNode, endNode)
        End Function

        ''' <summary>
        ''' 連結ルートを作成します。
        ''' </summary>
        ''' <param name="nodes">ノードリスト。</param>
        ''' <param name="target">式の範囲。</param>
        ''' <returns>接続点。</returns>
        Private Shared Function ConcatenationRoute(nodes As NodeList, target As ExpressionRange) As (st As Node, ed As Node)
            ' 最初のルートを作成
            Dim curNode = CreateRoute(nodes, target.SubRanges(0))

            ' それ以降のルートを連結
            For i As Integer = 1 To target.SubRanges.Count - 1
                Dim subRoute = CreateRoute(nodes, target.SubRanges(i))
                curNode.ed.Routes.Add(subRoute.st)
                curNode = (curNode.st, subRoute.ed)
            Next
            Return curNode
        End Function

        ''' <summary>
        ''' 範囲ルートを作成します。
        ''' </summary>
        ''' <param name="nodes">ノードリスト。</param>
        ''' <param name="target">式の範囲。</param>
        ''' <param name="minCount">最小回数。</param>
        ''' <param name="maxCount">最大回数。</param>
        ''' <returns>接続点。</returns>
        Private Shared Function RangeRoute(nodes As NodeList, target As ExpressionRange, minCount As Integer, maxCount As Integer) As (st As Node, ed As Node)
            Dim startNode = nodes.NewNode()
            Dim midRoute = CreateRoute(nodes, target)
            Dim endNode1 = nodes.NewNode()
            Dim endNode2 = nodes.NewNode()

            ' 開始点から中間点、中間点から終了点、開始点と終了点の相互へ接続
            startNode.Routes.Add(midRoute.st)
            midRoute.st.MinLimit = 0
            midRoute.st.MaxLimit = maxCount

            ' 開始点から終了点へ直接接続
            If minCount <= 0 Then
                startNode.Routes.Add(endNode2)
            End If

            ' 中間点から終了点へ接続
            midRoute.ed.Routes.Add(endNode2)
            midRoute.ed.Routes.Add(endNode1)

            ' 終了点から開始点へ接続（ループ）
            If maxCount > 1 Then
                endNode1.Routes.Add(startNode)
            End If

            ' 回数制限を設定
            endNode2.MinLimit = minCount
            endNode2.MaxLimit = Integer.MaxValue

            Return (startNode, endNode2)
        End Function

#End Region

#Region "パターン作成"

        ''' <summary>
        ''' パターンを作成します。
        ''' </summary>
        ''' <param name="pattern">パターン格納用辞書。</param>
        ''' <param name="startNode">開始ノード。</param>
        ''' <param name="endNode">終了ノード。</param>
        Private Shared Sub CreatePattern(pattern As SortedDictionary(Of Integer, NodeLink),
                                         startNode As Node,
                                         endNode As Node)
            Dim res As New NodeLink(startNode)
            Dim arrived As New HashSet(Of Integer)()
            CreatePattern(res, arrived, startNode, endNode, 0, Integer.MaxValue)
            pattern.Add(startNode.Id, res)
        End Sub


        ''' <summary>
        ''' パターンを作成します。
        ''' </summary>
        ''' <param name="pattern">パターン格納用。</param>
        ''' <param name="arrived">到達済みノードIDセット。</param>
        ''' <param name="startNode">開始ノード。</param>
        ''' <param name="endNode">終了ノード。</param>
        Private Shared Sub CreatePattern(pattern As NodeLink,
                                         arrived As HashSet(Of Integer),
                                         startNode As Node,
                                         endNode As Node,
                                         minLimit As Integer,
                                         maxLimit As Integer)
            For Each nd In startNode.Routes
                If Not arrived.Contains(nd.Id) Then
                    arrived.Add(nd.Id)

                    Dim minLmt = Math.Max(minLimit, nd.MinLimit)
                    Dim maxLmt = Math.Min(maxLimit, nd.MaxLimit)

                    If nd.Id = endNode.Id Then
                        pattern.EndNodes.Add((nd, minLmt, maxLmt))
                    ElseIf nd.IsEpsilon Then
                        CreatePattern(pattern, arrived, nd, endNode, minLmt, maxLmt)
                    Else
                        pattern.EndNodes.Add((nd, minLmt, maxLmt))
                    End If
                End If
            Next
        End Sub

#End Region

        ''' <summary>
        ''' マッチャーを取得する。
        ''' </summary>
        ''' <returns>マッチャー。</returns>
        Public Function GetMatcher() As AnalysisMatcher
            Return New AnalysisMatcher(Me._root, Me.RuleName)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"<{Me.RuleName}>"
        End Function

        ''' <summary>
        ''' キャッシュをクリアします。
        ''' </summary>
        ''' <param name="idHash">クリア済みノードIDセット。</param>
        Friend Sub ClearCache(idHash As HashSet(Of Integer))
            For Each route In Me._root.Routes
                If Not idHash.Contains(route.NextNode.Id) Then
                    idHash.Add(route.NextNode.Id)
                    route.NextNode.ClearCache(idHash)
                End If
            Next
        End Sub

        ''' <summary>評価ノード。</summary>
        Private NotInheritable Class Node

            ''' <summary>識別値。</summary>
            Public ReadOnly Property Id As Integer

            ''' <summary>εか。</summary>
            Public ReadOnly Property IsEpsilon As Boolean

            ''' <summary>評価範囲。</summary>
            Public ReadOnly Property Range As ExpressionRange

            ''' <summary>接続ルート。</summary>
            Public ReadOnly Property Routes As List(Of Node)

            ''' <summary>最小出現回数。</summary>
            Public Property MinLimit As Integer

            ''' <summary>最大出現回数。</summary>
            Public Property MaxLimit As Integer

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="id">識別値。</param>
            Public Sub New(id As Integer)
                Me.Id = id
                Me.IsEpsilon = True
                Me.Range = ExpressionRange.Invalid
                Me.Routes = New List(Of Node)()
                Me.MinLimit = 0
                Me.MaxLimit = Integer.MaxValue
            End Sub

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="id">識別値。</param>
            ''' <param name="range">評価範囲。</param>
            Public Sub New(id As Integer, range As ExpressionRange)
                Me.Id = id
                Me.IsEpsilon = False
                Me.Range = range
                Me.Routes = New List(Of Node)()
                Me.MinLimit = 0
                Me.MaxLimit = Integer.MaxValue
            End Sub

            ''' <summary>
            ''' 文字列表現を取得する。
            ''' </summary>
            ''' <returns>文字列表現。</returns>
            Overrides Function ToString() As String
                Return $"-> {Me.Id} {Me.Range} [{Me.MinLimit}, {Me.MaxLimit}]"
            End Function

        End Class

        ''' <summary>ノードリスト。</summary>
        Private NotInheritable Class NodeList
            Inherits List(Of Node)

            ''' <summary>評価範囲の設定して評価ノードを新規作成してリストに追加します。</summary>
            ''' <param name="range">評価範囲。</param>
            ''' <returns>評価ノード。</returns>
            Public Function NewNode(range As ExpressionRange) As Node
                Dim nd As New Node(Me.Count, range)
                Me.Add(nd)
                Return nd
            End Function

            ''' <summary>評価ノードを新規作成してリストに追加します。</summary>
            Public Function NewNode() As Node
                Dim nd As New Node(Me.Count)
                Me.Add(nd)
                Return nd
            End Function

        End Class

        ''' <summary>ノードパターン。</summary>
        Private NotInheritable Class NodeLink

            ''' <summary>開始ノード。</summary>
            Public ReadOnly Property StartNode As Node

            ''' <summary>終了ノードリスト。</summary>
            Public ReadOnly Property EndNodes As List(Of (Node, Integer, Integer))

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="startNode">開始ノード。</param>
            Public Sub New(startNode As Node)
                Me.StartNode = startNode
                Me.EndNodes = New List(Of (Node, Integer, Integer))()
            End Sub

        End Class

    End Class

End Namespace
