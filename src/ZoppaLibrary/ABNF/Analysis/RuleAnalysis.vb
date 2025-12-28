Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ルールのコンパイル済み式を表します。
    ''' </summary>
    Public NotInheritable Class RuleAnalysis
        Implements IAnalysis

        ''' <summary>
        ''' 解析ノードのルート。
        ''' </summary>
        Private _root As AnalysisNode

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
            Dim analysis As New SortedDictionary(Of Integer, AnalysisNode)() ' 1
            For Each kvp In pattern
                With kvp.Value.StartNode
                    Dim ana = AnalysisNode.Create(.Id, .Range)
                    analysis.Add(ana.Id, ana)
                End With
            Next
            For Each kvp In pattern ' 2
                For Each endEdge In kvp.Value.EndNodes
                    If analysis.ContainsKey(endEdge.Item1.Id) Then
                        analysis(kvp.Key).AddRoute(analysis(endEdge.Item1.Id), endEdge.Item2, endEdge.Item3)
                    End If
                Next
            Next
            Me._root = analysis(startNode.Id)
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

            If minCount <= 0 Then
                startNode.Routes.Add(endNode2)
            End If

            midRoute.ed.Routes.Add(endNode1)
            midRoute.ed.Routes.Add(endNode2)

            If maxCount > 1 Then
                endNode1.Routes.Add(startNode)
            End If

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
            Return New AnalysisMatcher(Me._root)
        End Function

        ''' <summary>
        ''' 解析イテレーター。
        ''' </summary>
        Public NotInheritable Class AnalysisMatcher

            ''' <summary>ルートノード。</summary>
            Private _root As AnalysisNode

            ''' <summary>解析スタック。</summary>
            Private _stack As New Stack(Of (AnalysisNode, Integer, Integer, ABNFAnalysisItem))()

            ''' <summary>到達回数記録。</summary>
            Private _arrived As New SortedDictionary(Of Integer, Integer)()

            ''' <summary>
            ''' コンストラクタ。
            ''' </summary>
            ''' <param name="root">ルートノード。</param>
            Public Sub New(root As AnalysisNode)
                Me._root = root
            End Sub

            Public Function Match(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, shift As Integer)
                If Me._stack.Count = 0 Then
                    ' 初回開始
                    Me._arrived.Clear()
                    Return Me.Tracking(Me._root, 0, tr, env)
                Else
                    ' 継続解析
                    Dim cur = Me._stack.Pop()
                    tr.Seek(cur.Item3)
                    Me.CountArrived()
                    Return Me.Tracking(cur.Item1, cur.Item2, tr, env)
                End If
            End Function

            ''' <summary>
            ''' 次の解析ステップを実行する。
            ''' </summary>
            ''' <param name="tr">位置調整リーダー。</param>
            ''' <param name="env">解析環境。</param>
            ''' <returns>解析が成功した場合に True を返します。</returns>
            Public Function MoveNext(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, shift As Integer)
                If Me._stack.Count = 0 Then
                    ' 初回開始
                    Me._arrived.Clear()
                    Return Me.Tracking(Me._root, 0, tr, env)
                Else
                    ' 継続解析
                    Dim cur = Me._stack.Pop()
                    tr.Seek(cur.Item3)
                    Me.CountArrived()
                    Return Me.Tracking(cur.Item1, cur.Item2 + 1, tr, env)
                End If
            End Function

            Private Function Tracking(node As AnalysisNode,
                                      route As Integer,
                                      tr As PositionAdjustBytes,
                                      env As ABNFEnvironment) As (success As Boolean, shift As Integer)
                Dim currentPosition = tr.Position

start_label:
                Do While route < node.Routes.Count
                        Dim nextNode = node.Routes(route).NextNode
                        Dim fromArrived = Me.GetArrived(node.Id)
                        Dim toArrived = Me.GetArrived(nextNode.Id)
                        Dim minLmt = node.Routes(route).RequiredVisits
                        Dim maxLmt = node.Routes(route).LimitedVisits

                        ' 最小訪問回数に達していない場合は次のルートへ
                        If fromArrived < minLmt Then
                            route += 1
                            Continue Do
                        End If

                        ' 訪問回数が上限を超えている場合は次のルートへ
                        If toArrived >= maxLmt Then
                            route += 1
                            Continue Do
                        End If

                        ' 対象ノードが一致するか判定
                        Dim matched = nextNode.Match(tr, env)
                        If matched.success Then
                            ' 最終ノードに到達した場合は成功
                            If nextNode.Routes.Count = 0 Then
                                Return (True, 0)
                            End If

                            ' 次のノードへ進む
                            Me._stack.Push((nextNode, route, currentPosition, matched.answer))
                            currentPosition = tr.Position
                            node = nextNode
                            route = 0
                            Me.IncrementArrived(nextNode.Id)
                        Else
                            route += 1
                            tr.Seek(currentPosition)
                            '' ひとつ前のノードへ戻る
                            'If Me._stack.Count > 0 Then
                            '    Dim preview = Me._stack.Pop()
                            '    node = preview.Item1
                            '    route = preview.Item2 + 1
                            '    tr.Seek(preview.Item3)
                            '    Me.DecrementArrived(node.Id)
                            'Else
                            '    ' 候補ルートが存在しない場合は失敗
                            '    Return (False, 0)
                            'End If
                        End If
                    Loop


                Do While True
                    If Me._stack.Count > 0 Then
                        Dim preview = Me._stack.Pop()
                        node = preview.Item1
                        tr.Seek(preview.Item3)

                        Dim retry = node.MoveNext(tr, env)
                        If retry.success Then
                            Me._stack.Push((node, route, preview.Item3, retry.answer))
                            route = 0
                            currentPosition = tr.Position
                            GoTo start_label
                        ElseIf preview.Item2 + 1 < node.Routes.Count Then
                            route = preview.Item2 + 1
                            currentPosition = preview.Item3
                            Me.DecrementArrived(node.Id)
                            GoTo start_label
                        Else
                            currentPosition = preview.Item3
                            Me.DecrementArrived(node.Id)
                            If TypeOf preview.Item1 IsNot AnalysisNode.RuleNameNode Then
                                route = preview.Item2 + 1
                                Exit Do
                            End If
                        End If
                    Else
                        ' 候補ルートが存在しない場合は失敗
                        Exit Do
                    End If
                Loop


                'node = preview.Item1
                'route = preview.Item2 + 1
                'currentPosition = preview.Item3
                'Me.DecrementArrived(node.Id)

                Return (False, 0)
            End Function

            Private Sub CountArrived()
                Dim buf As New SortedDictionary(Of Integer, Integer)()
                For Each item In Me._stack
                    If buf.ContainsKey(item.Item1.Id) Then
                        buf(item.Item1.Id) += 1
                    Else
                        buf.Add(item.Item1.Id, 1)
                    End If
                Next
                Me._arrived = buf
            End Sub

            Private Sub IncrementArrived(nodeId As Integer)
                If Me._arrived.ContainsKey(nodeId) Then
                    Me._arrived(nodeId) += 1
                Else
                    Me._arrived.Add(nodeId, 1)
                End If
            End Sub

            Private Sub DecrementArrived(nodeId As Integer)
                Me._arrived(nodeId) -= 1
                If Me._arrived(nodeId) <= 0 Then
                    Me._arrived.Remove(nodeId)
                End If
            End Sub

            Private Function GetArrived(nodeId As Integer) As Integer
                Return If(Me._arrived.ContainsKey(nodeId), Me._arrived(nodeId), 0)
            End Function

            ''' <summary>
            ''' 解析結果を取得する。
            ''' </summary>
            ''' <returns>解析結果リスト。</returns>
            Function GetAnswer() As List(Of ABNFAnalysisItem)
                Dim res As New List(Of ABNFAnalysisItem)()
                For Each item In Me._stack
                    If item.Item4 IsNot Nothing Then
                        res.Add(item.Item4)
                    End If
                Next
                res.Reverse()
                Return res
            End Function

        End Class

        ''' <summary>
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <param name="ruleTable">ルール解析テーブル。</param>
        ''' <param name="ruleName">現在のルール名。</param>
        ''' <param name="answers">解析結果のリスト。</param>
        ''' <param name="counter">訪問回数カウンター。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As PositionAdjustBytes,
                              env As ABNFEnvironment,
                              ruleTable As SortedDictionary(Of String, RuleAnalysis),
                              ruleName As String,
                              answers As List(Of ABNFAnalysisItem),
                              counter As Dictionary(Of IAnalysis, Integer)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            'Dim snap = tr.MemoryPosition()

            '' ルールパターンを順に評価
            'Dim shift As Integer = Integer.MaxValue
            'For Each evalExpr In Me.Pattern
            '    answers.Clear()

            '    If evalExpr.MinLimit <= 0 Then
            '        counter.Clear()
            '        counter.Add(evalExpr.ToAnalysis, 1)

            '        Dim res = evalExpr.Match(tr, env, ruleTable, ruleName, answers, counter)
            '        If res.sccess Then
            '            Return (True, 0)
            '        ElseIf res.shift < shift Then
            '            shift = res.shift
            '        End If
            '    End If
            'Next

            'snap.Restore()
            Return (False, 0)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"<{Me.RuleName}>"
        End Function

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
            ''' <returns>評価ノード。</returns>
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
