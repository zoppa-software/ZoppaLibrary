Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ルールのコンパイル済み式を表します。
    ''' </summary>
    Public NotInheritable Class RuleAnalysis
        Implements IAnalysis

        ''' <summary>
        ''' ルール名を取得する。
        ''' </summary>
        ''' <returns>ルール名。</returns>
        Public ReadOnly Property RuleName As String

        ''' <summary>
        ''' ルールのパターンを取得する。
        ''' </summary>
        ''' <returns>ルールのパターン。</returns>
        Public ReadOnly Property Pattern As List(Of IAnalysis.Link) Implements IAnalysis.Pattern

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

            startNode.Routes.Add(New Edge(routes.st))
            routes.ed.Routes.Add(New Edge(endNode))

            ' ノードのリンクを作成
            Dim pattern As New SortedDictionary(Of Integer, NodeLink)()
            CreatePattern(pattern, startNode, endNode)
            For i As Integer = 1 To nodes.Count - 2
                If Not nodes(i).IsEpsilon Then
                    CreatePattern(pattern, nodes(i), endNode)
                End If
            Next
            pattern.Add(endNode.Id, New NodeLink(endNode))

            ' 評価用グラフを作成
            Dim links As New SortedDictionary(Of Integer, IAnalysis)()
            For Each kvp In pattern
                With kvp.Value.StartNode
                    Dim ana As IAnalysis = Nothing
                    If .Id = endNode.Id Then
                        ana = CompletedAnalysis.Instance
                    ElseIf .Id = startNode.Id Then
                        ana = New BeginAnalysis(.Range)
                    Else
                        Select Case .Range.Expr.GetType()
                            'Case GetType(FactorExpression) ' 要素式
                            '    ana = New FactorAnalysis(.Range)
                            Case GetType(RuleNameExpression) ' 識別子式
                                ana = New RuleNameAnalysis(.Range)
                            'Case GetType(SpecialSeqExpression) ' 特殊式
                            '    ana = New SpecialSeqAnalysis(.Range)
                            Case GetType(CharValExpression) ' 文字列
                                ana = New CharValAnalysis(.Range)
                            Case Else
                                Throw New NotSupportedException($"サポートされていない式タイプです: { .Range.Expr.GetType().FullName }")
                        End Select
                    End If
                    links.Add(.Id, ana)
                End With
            Next
            For Each kvp In pattern
                For Each endEdge In kvp.Value.EndNodes
                    If links.ContainsKey(endEdge.ToNode.Id) Then
                        links(kvp.Key).Pattern.Add(New IAnalysis.Link(links(endEdge.ToNode.Id), endEdge.MinLimit, endEdge.MaxLimit))
                    End If
                Next
            Next

            Me.Pattern = links(startNode.Id).Pattern
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

                    'Case GetType(FactorExpression)
                    '    ' 要素式
                    '    If target.SubRanges.Count > 1 Then
                    '        Select Case target.SubRanges(1).ToString()
                    '            Case "?"c
                    '                Return ZeroOrOneRoute(nodes, target.SubRanges(0))
                    '            Case "*"c
                    '                Return ZeroOrMoreRoute(nodes, target.SubRanges(0))
                    '            Case "+"c
                    '                Return OneOrMoreRoute(nodes, target.SubRanges(0))
                    '            Case Else
                    '                ' 否定式
                    '                Return DirectRoute(nodes, target)
                    '        End Select
                    '    Else
                    '        Return CreateRoute(nodes, target.SubRanges(0))
                    '    End If

                    'Case GetType(IdentifierExpression)
                    '    ' 識別子式
                    '    Return DirectRoute(nodes, target)

                    'Case GetType(SpecialSeqExpression)
                    '    ' 特殊式
                    '    Return DirectRoute(nodes, target)

                    'Case GetType(TermExpression)
                    '    ' 終端式
                    '    Select Case target.SubRanges(0).Expr.GetType()
                    '        Case GetType(TerminalExpression), GetType(IdentifierExpression)
                    '            Return CreateRoute(nodes, target.SubRanges(0))
                    '        Case Else
                    '            Select Case target.SubChar(0)
                    '                Case "["c
                    '                    Return ZeroOrOneRoute(nodes, target.SubRanges(0))
                    '                Case "{"c
                    '                    Return ZeroOrMoreRoute(nodes, target.SubRanges(0))
                    '                Case Else
                    '                    Return CreateRoute(nodes, target.SubRanges(0))
                    '            End Select
                    '    End Select

                    'Case GetType(TerminalExpression)
                    '    ' 終端記号
                    '    Return DirectRoute(nodes, target)

                Case GetType(GroupExpression)
                    ' グループ式
                    Return CreateRoute(nodes, target.SubRanges(0))

                Case GetType(OptionExpression)
                    ' オプション式
                    Return RangeRoute(nodes, target.SubRanges(0), 0, 1)

                Case GetType(RepetitionExpression)
                    ' 反復式
                    If target.SubRanges.Count > 1 Then
                        Dim minRange = target.SubRanges(0).GetRange(0)
                        Dim maxRange = target.SubRanges(0).GetRange(1)

                        Dim minCount = If(minRange.Enable, Integer.Parse(minRange.ToString()), 0)
                        Dim maxCount = If(maxRange.Enable, Integer.Parse(maxRange.ToString()), Integer.MaxValue)

                        Return RangeRoute(nodes, target.SubRanges(1), minCount, maxCount)
                    Else
                        Return CreateRoute(nodes, target.SubRanges(0))
                    End If

                Case GetType(RuleNameExpression)
                    ' ルール名式
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
            startNode.Routes.Add(New Edge(endNode))

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
                startNode.Routes.Add(New Edge(subRoute.st))
                subRoute.ed.Routes.Add(New Edge(endNode))
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
                curNode.ed.Routes.Add(New Edge(subRoute.st))
                curNode = (curNode.st, subRoute.ed)
            Next
            Return curNode
        End Function

        '''' <summary>
        '''' 0回または1回のルートを作成します。
        '''' </summary>
        '''' <param name="nodes">ノードリスト。</param>
        '''' <param name="target">式の範囲。</param>
        '''' <returns>接続点。</returns>
        'Private Shared Function ZeroOrOneRoute(nodes As NodeList, target As ExpressionRange) As (st As Node, ed As Node)
        '    Dim startNode = nodes.NewNode()
        '    Dim midRoute = CreateRoute(nodes, target)
        '    Dim endNode = nodes.NewNode()

        '    ' 開始点から中間点、中間点から終了点、開始点から終了点へ接続
        '    startNode.Routes.Add(midRoute.st)
        '    midRoute.ed.Routes.Add(endNode)
        '    startNode.Routes.Add(endNode)

        '    Return (startNode, endNode)
        'End Function

        '''' <summary>
        '''' 0回以上のルートを作成します。
        '''' </summary>
        '''' <param name="nodes">ノードリスト。</param>
        '''' <param name="target">式の範囲。</param>
        '''' <returns>接続点。</returns>
        'Private Shared Function ZeroOrMoreRoute(nodes As NodeList, target As ExpressionRange) As (st As Node, ed As Node)
        '    Dim startNode = nodes.NewNode()
        '    Dim midRoute = CreateRoute(nodes, target)
        '    Dim endNode = nodes.NewNode()

        '    ' 開始点から中間点、中間点から終了点、開始点と終了点の相互へ接続
        '    startNode.Routes.Add(midRoute.st)
        '    startNode.Routes.Add(endNode)
        '    midRoute.ed.Routes.Add(endNode)
        '    endNode.Routes.Add(startNode)

        '    Return (startNode, endNode)
        'End Function

        '''' <summary>
        '''' 1回以上のルートを作成します。
        '''' </summary>
        '''' <param name="nodes">ノードリスト。</param>
        '''' <param name="target">式の範囲。</param>
        '''' <returns>接続点。</returns>
        'Private Shared Function OneOrMoreRoute(nodes As NodeList, target As ExpressionRange) As (st As Node, ed As Node)
        '    Dim startNode = nodes.NewNode()
        '    Dim midRoute = CreateRoute(nodes, target)
        '    Dim endNode = nodes.NewNode()

        '    ' 開始点から中間点、中間点から終了点、終了点から開始点へ接続
        '    startNode.Routes.Add(midRoute.st)
        '    midRoute.ed.Routes.Add(endNode)
        '    endNode.Routes.Add(startNode)

        '    Return (startNode, endNode)
        'End Function

        Private Shared Function RangeRoute(nodes As NodeList, target As ExpressionRange, minCount As Integer, maxCount As Integer) As (st As Node, ed As Node)
            Dim startNode = nodes.NewNode()
            Dim midRoute = CreateRoute(nodes, target)
            Dim endNode = nodes.NewNode()

            ' 開始点から中間点、中間点から終了点、開始点と終了点の相互へ接続
            startNode.Routes.Add(New Edge(midRoute.st, 0, maxCount))
            startNode.Routes.Add(New Edge(endNode, minCount, Integer.MaxValue))
            midRoute.ed.Routes.Add(New Edge(endNode, minCount, Integer.MaxValue))
            endNode.Routes.Add(New Edge(startNode, minCount, maxCount))

            Return (startNode, endNode)
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
                If Not arrived.Contains(nd.ToNode.Id) Then
                    arrived.Add(nd.ToNode.Id)
                    If nd.ToNode.Id = endNode.Id Then
                        pattern.EndNodes.Add(New Edge(nd.ToNode, minLimit, maxLimit))
                    ElseIf nd.ToNode.IsEpsilon Then
                        CreatePattern(pattern, arrived, nd.ToNode, endNode, Math.Max(minLimit, nd.MinLimit), Math.Min(maxLimit, nd.MaxLimit))
                    Else
                        pattern.EndNodes.Add(New Edge(nd.ToNode, minLimit, maxLimit))
                    End If
                End If
            Next
        End Sub

#End Region


        Public Function Match(tr As IPositionAdjustReader, env As ABNFEnvironment, ruleTable As SortedDictionary(Of String, RuleAnalysis), ruleName As String, answers As List(Of ABNFAnalysisItem)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            Throw New NotImplementedException()
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"<{Me.RuleName}>"
        End Function

        ''' <summary>評価ノード。</summary>
        Private Structure Node

            ''' <summary>識別値。</summary>
            Public ReadOnly Property Id As Integer

            ''' <summary>εか。</summary>
            Public ReadOnly Property IsEpsilon As Boolean

            ''' <summary>評価範囲。</summary>
            Public ReadOnly Property Range As ExpressionRange

            ''' <summary>接続ルート。</summary>
            Public ReadOnly Property Routes As List(Of Edge)

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="id">識別値。</param>
            Public Sub New(id As Integer)
                Me.Id = id
                Me.IsEpsilon = True
                Me.Range = ExpressionRange.Invalid
                Me.Routes = New List(Of Edge)()
            End Sub

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="id">識別値。</param>
            ''' <param name="range">評価範囲。</param>
            Public Sub New(id As Integer, range As ExpressionRange)
                Me.Id = id
                Me.IsEpsilon = False
                Me.Range = range
                Me.Routes = New List(Of Edge)()
            End Sub
        End Structure

        Private Structure Edge

            Public ReadOnly Property ToNode As Node

            Public ReadOnly Property MinLimit As Integer

            Public ReadOnly Property MaxLimit As Integer

            Public Sub New(toNode As Node, minLimit As Integer, maxLimit As Integer)
                Me.ToNode = toNode
                Me.MinLimit = minLimit
                Me.MaxLimit = maxLimit
            End Sub

            Public Sub New(toNode As Node)
                Me.ToNode = toNode
                Me.MinLimit = 0
                Me.MaxLimit = Integer.MaxValue
            End Sub

        End Structure

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
            Public ReadOnly Property EndNodes As List(Of Edge)

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="startNode">開始ノード。</param>
            Public Sub New(startNode As Node)
                Me.StartNode = startNode
                Me.EndNodes = New List(Of Edge)()
            End Sub

        End Class

#If False Then

        ''' <summary>
        ''' ルール名を取得する。
        ''' </summary>
        ''' <returns>ルール名。</returns>
        Public ReadOnly Property RuleName As String

        ''' <summary>
        ''' ルールのパターンを取得する。
        ''' </summary>
        ''' <returns>ルールのパターン。</returns>
        Public ReadOnly Property Pattern As List(Of IAnalysis) Implements IAnalysis.Pattern

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
                If Not nodes(i).IsEpsilon Then
                    CreatePattern(pattern, nodes(i), endNode)
                End If
            Next
            pattern.Add(endNode.Id, New NodeLink(endNode))

            ' 評価用グラフを作成
            Dim links As New SortedDictionary(Of Integer, IAnalysis)()
            For Each kvp In pattern
                With kvp.Value.StartNode
                    Dim ana As IAnalysis = Nothing
                    If .Id = endNode.Id Then
                        ana = CompletedAnalysis.Instance
                    ElseIf .Id = startNode.Id Then
                        ana = New BeginAnalysis(.Range)
                    Else
                        Select Case .Range.Expr.GetType()
                            Case GetType(FactorExpression) ' 要素式
                                ana = New FactorAnalysis(.Range)
                            Case GetType(IdentifierExpression) ' 識別子式
                                ana = New IdentifierAnalysis(.Range)
                            Case GetType(SpecialSeqExpression) ' 特殊式
                                ana = New SpecialSeqAnalysis(.Range)
                            Case GetType(TerminalExpression) ' 終端記号
                                ana = New TerminalAnalysis(.Range)
                            Case Else
                                Throw New NotSupportedException($"サポートされていない式タイプです: { .Range.Expr.GetType().FullName }")
                        End Select
                    End If
                    links.Add(.Id, ana)
                End With
            Next
            For Each kvp In pattern
                For Each endNode In kvp.Value.EndNodes
                    If links.ContainsKey(endNode.Id) Then
                        links(kvp.Key).Pattern.Add(links(endNode.Id))
                    End If
                Next
            Next

            Me.Pattern = links(startNode.Id).Pattern
        End Sub

        ''' <summary>
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <param name="ruleTable">ルール解析テーブル。</param>
        ''' <param name="specialMethods">特殊メソッドテーブル。</param>
        ''' <param name="ruleName">現在のルール名。</param>
        ''' <param name="answers">解析結果のリスト。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              env As EBNFEnvironment,
                              ruleTable As SortedDictionary(Of String, RuleAnalysis),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              ruleName As String,
                              answers As List(Of EBNFAnalysisItem)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            Dim snap = tr.MemoryPosition()

            ' ルールパターンを順に評価
            Dim shift As Integer = Integer.MaxValue
            For Each evalExpr In Me.Pattern
                answers.Clear()

                Dim res = evalExpr.Match(tr, env, ruleTable, specialMethods, ruleName, answers)
                If res.sccess Then
                    Return (True, 0)
                ElseIf res.shift < shift Then
                    shift = res.shift
                End If
            Next

            snap.Restore()
            Return (False, shift)
        End Function

#End If

    End Class

End Namespace
