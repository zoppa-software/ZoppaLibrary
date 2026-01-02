Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ABNF 解析マッチャー。
    ''' </summary>
    Public NotInheritable Class AnalysisMatcher
        Implements IAnalysisMatcher

        ''' <summary>最大反復回数。</summary>
        Private Const MaxIterations As Integer = 10000

        ''' <summary>最大スタック深度。</summary>
        Private Const MaxStackDepth As Integer = 10000

        ''' <summary>実用的な量詞上限（メモリ枯渇防止のため）。</summary>
        ''' <remarks>無制限量詞(*)に対して適用される現実的な上限値</remarks>
        Private Const PracticalQuantifierLimit As Integer = 10000

        ''' <summary>遡りアクション。</summary>
        Private Enum BacktrackAction

            ''' <summary>マッチをリトライする。</summary>
            RetryMatch

            ''' <summary>MoveNext をリトライする。</summary>
            RetryMoveNext

            ''' <summary>追跡を終了する。</summary>
            ExitTracking

        End Enum

        ''' <summary>ルートノード。</summary>
        Private ReadOnly _root As AnalysisNode

        ''' <summary>ルール名。</summary>
        Private ReadOnly _ruleName As String

        ''' <summary>解析スタック。</summary>
        Private ReadOnly _stack As New Stack(Of StackState)()

        ''' <summary>到達回数記録。</summary>
        Private ReadOnly _arrived As New Dictionary(Of Integer, Integer)()

        ''' <summary>直前の結果。</summary>
        Private _previewValue As (startPosition As Integer, endPosition As Integer) = (-1, 0)

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="root">ルートノード。</param>
        Public Sub New(root As AnalysisNode, ruleName As String)
            Me._root = root
            Me._ruleName = ruleName
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As PositionAdjustBytes,
                              env As ABNFEnvironment) As (success As Boolean, shift As Integer) Implements IAnalysisMatcher.Match
            If Me._stack.Count = 0 Then
                ' 初回開始
                Me._arrived.Clear()
                Me.IncrementArrived(Me._root.Id)
                Return Me.Tracking(Me._root, 0, tr, env)
            Else
                ' 継続解析
                Dim cur = Me.PopStack()
                tr.Seek(cur.StartPosition)
                Return Me.Tracking(cur.FromNode, cur.Route, tr, env)
            End If
        End Function

        ''' <summary>
        ''' 次の解析ステップを実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function MoveNext(tr As PositionAdjustBytes,
                                 env As ABNFEnvironment) As (success As Boolean, shift As Integer) Implements IAnalysisMatcher.MoveNext
            If Me._stack.Count = 0 Then
                ' 初回開始
                Me._arrived.Clear()
                Me.IncrementArrived(Me._root.Id)
                Return Me.Tracking(Me._root, 0, tr, env)
            Else
                ' 継続解析
                Dim cur = Me.PopStack()
                tr.Seek(cur.StartPosition)
                Return Me.Tracking(cur.FromNode, cur.Route + 1, tr, env)
            End If
        End Function

        ''' <summary>
        ''' 解析を追跡する。
        ''' </summary>
        ''' <param name="node">現在のノード。</param>
        ''' <param name="route">現在のルート番号。</param>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>
        ''' 解析が成功した場合に True を返します。
        ''' shift パラメータは将来の拡張用で、現在は常に0を返します。
        ''' </returns>
        Private Function Tracking(node As AnalysisNode,
                                  route As Integer,
                                  tr As PositionAdjustBytes,
                                  env As ABNFEnvironment) As (success As Boolean, shift As Integer)
            Dim iterationCount As Integer = 0
            Dim startPosition = tr.Position
            Dim currentPosition = tr.Position
            Dim action As BacktrackAction

            Do
                iterationCount += 1
                If iterationCount > MaxIterations Then
                    Throw New ABNFException($"解析が最大反復回数({MaxIterations})を超過しました。無限ループの可能性があります。ルール:{Me._ruleName} 位置:{currentPosition}")
                End If

                If Me._stack.Count > MaxStackDepth Then
                    Throw New ABNFException($"解析スタックが最大深度({MaxStackDepth})を超過しました。ルール:{Me._ruleName} 位置:{currentPosition}")
                End If

                ' ルートを順次試行、一致を確認
                Do While route < node.Routes.Count
                    Dim nextNode = node.Routes(route).NextNode
                    Dim fromArrived = Me.GetArrived(node.Id)
                    Dim toArrived = Me.GetArrived(nextNode.Id)
                    Dim required = node.Routes(route).RequiredVisits
                    Dim limited = node.Routes(route).LimitedVisits

                    ' 最小訪問回数に達していない場合は次のルートへ
                    If fromArrived < required Then
                        route += 1
                        Continue Do
                    End If

                    ' 訪問回数が上限を超えている場合は次のルートへ
                    If IsVisitLimitExceeded(toArrived, limited) Then
                        route += 1
                        Continue Do
                    End If

                    ' 対象ノードが一致するか判定
                    Dim matched = nextNode.Match(tr, env, Me._ruleName)
                    If matched.success Then
                        ' 評価をスタックに保存
                        Me.PushStack(node, nextNode, route, currentPosition, matched.answer)

                        ' 最終ノードに到達した場合は成功
                        If nextNode.Routes.Count = 0 Then
                            If Me._previewValue.startPosition = startPosition AndAlso
                               Me._previewValue.endPosition = tr.Position Then
                                Return (False, 0)
                            End If
                            Return (True, 0)
                        End If

                        ' 次のノードへ進む
                        currentPosition = tr.Position
                        node = nextNode
                        route = 0
                    Else
                        ' ノードが一致しなかった場合は次のルートへ
                        route += 1
                        tr.Seek(currentPosition)
                    End If
                Loop

                ' 全てのルートを試行しても一致しない場合はノードを遡る
                Do
                    ' ノードを遡る
                    Dim backtrack = Me.BacktrackNode(route, tr, env)

                    Select Case backtrack.Action
                        Case BacktrackAction.ExitTracking
                            ' 追跡終了の場合は失敗
                            Return (False, 0)
                        Case Else
                            ' 上記以外は継続
                            action = backtrack.Action
                            node = backtrack.Node
                            route = backtrack.Route
                            currentPosition = backtrack.Position
                    End Select
                Loop While action = BacktrackAction.RetryMoveNext
            Loop While action = BacktrackAction.RetryMatch
        End Function

        ''' <summary>
        ''' 訪問制限をチェックする。
        ''' </summary>
        Private Function IsVisitLimitExceeded(toArrived As Integer, limited As Integer) As Boolean
            If limited = Integer.MaxValue Then
                ' 無制限量詞(*や+)の場合は実用的な上限を適用
                Return toArrived >= PracticalQuantifierLimit
            Else
                Return toArrived >= limited
            End If
        End Function

        ''' <summary>
        ''' ノードを遡る。
        ''' </summary>
        ''' <param name="route">現在のルート番号。</param>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>追跡状態。</returns>
        Private Function BacktrackNode(route As Integer,
                                       tr As PositionAdjustBytes,
                                       env As ABNFEnvironment) As TrackingState
            If Me._stack.Count > 0 Then
                Dim selectedRoute = route
                Dim currentPosition = tr.Position

                ' ひとつ前の評価を取得
                Dim preview = Me.PopStack()
                tr.Seek(preview.StartPosition)

                ' 再評価を行い、どのように遡るか判断する
                ' 1. 現在のノード（ひとつ前の次のノード）のリトライを実施し、リトライ成功の場合は現在のノードの最初のルートから再開
                ' 2. 現在のノードの選択肢が存在する場合は次の選択肢へ
                ' 3. 選択肢が存在しない、かつリトライ可能な場合はリトライへ
                Dim retry = preview.ToNode.MoveNext(tr, env)
                If retry.success Then
                    ' 1. リトライ成功の場合は現在のノードの最初のルートから再開
                    Me.PushStack(preview.FromNode, preview.ToNode, preview.Route, preview.StartPosition, retry.answer)
                    Return New TrackingState(BacktrackAction.RetryMatch, preview.ToNode, 0, tr.Position)

                ElseIf preview.Route + 1 < preview.FromNode.Routes.Count Then
                    ' 2. 選択肢が存在する場合は次の選択肢へ
                    Return New TrackingState(BacktrackAction.RetryMatch, preview.FromNode, preview.Route + 1, preview.StartPosition)

                Else
                    ' 3. 選択肢が存在しない、かつリトライ可能な場合はリトライへ。そうでない場合は終了
                    currentPosition = preview.StartPosition
                    If preview.FromNode.IsRetry Then
                        Return New TrackingState(BacktrackAction.RetryMoveNext, preview.FromNode, selectedRoute, tr.Position)
                    Else
                        Return New TrackingState(BacktrackAction.ExitTracking, Nothing, 0, 0)
                    End If
                End If
            Else
                ' 遡るノードが存在しない場合は終了
                Return New TrackingState(BacktrackAction.ExitTracking, Nothing, 0, 0)
            End If
        End Function

        ''' <summary>
        ''' 解析スタックに要素をプッシュする。
        ''' </summary>
        ''' <param name="fromNode">開始ノード。</param>
        ''' <param name="toNode">終了ノード。</param>
        ''' <param name="route">ルート番号。</param>
        ''' <param name="startPosition">開始位置。</param>
        ''' <param name="answer">解析結果。</param>
        Private Sub PushStack(fromNode As AnalysisNode,
                              toNode As AnalysisNode,
                              route As Integer,
                              startPosition As Integer,
                              answer As ABNFAnalysisItem)
            ' スタックに要素をプッシュする。
            Me._stack.Push(New StackState(fromNode, toNode, route, startPosition, answer))

            ' 到達回数をインクリメント
            Me.IncrementArrived(toNode.Id)
        End Sub

        ''' <summary>
        ''' 指定ノードの訪問回数をインクリメントする。
        ''' </summary>
        ''' <param name="nodeId">ノードID。</param>
        Private Sub IncrementArrived(nodeId As Integer)
            Dim currentCount As Integer
            If Me._arrived.TryGetValue(nodeId, currentCount) Then
                Me._arrived(nodeId) = currentCount + 1
            Else
                Me._arrived.Add(nodeId, 1)
            End If
        End Sub

        ''' <summary>
        ''' 解析スタックから要素をポップする。
        ''' </summary>
        ''' <returns>ポップしたスタック要素。</returns>
        Private Function PopStack() As StackState
            ' 解析スタックから要素をポップする。
            Dim item = Me._stack.Pop()

            ' 到達回数をデクリメント
            Me._arrived(item.ToNode.Id) -= 1
            If Me._arrived(item.ToNode.Id) <= 0 Then
                Me._arrived.Remove(item.ToNode.Id)
            End If

            Return item
        End Function

        ''' <summary>
        ''' 指定ノードの訪問回数を取得する。
        ''' </summary>
        ''' <param name="nodeId">ノードID。</param>
        ''' <returns>訪問回数。</returns>
        Private Function GetArrived(nodeId As Integer) As Integer
            Dim count As Integer = 0
            Me._arrived.TryGetValue(nodeId, count)
            Return count
        End Function

        ''' <summary>
        ''' 解析結果を取得する。
        ''' </summary>
        ''' <returns>解析結果リスト。</returns>
        Public Function GetAnswer() As List(Of ABNFAnalysisItem) Implements IAnalysisMatcher.GetAnswer
            ' 解析スタックが空の場合は空リストを返す
            If Me._stack.Count = 0 Then
                Return New List(Of ABNFAnalysisItem)()
            End If

            ' 解析スタックから解析結果を収集する
            Dim res As New List(Of ABNFAnalysisItem)(Me._stack.Count)
            For Each item In Me._stack
                If item.Answer IsNot Nothing Then
                    res.Add(item.Answer)
                End If
            Next
            res.Reverse()
            Return res
        End Function

        ''' <summary>
        ''' キャッシュをクリアします。
        ''' </summary>
        Public Sub ClearCache() Implements IAnalysisMatcher.ClearCache
            Dim idHash As New HashSet(Of Integer)()
            Me._root.ClearCache(idHash)
        End Sub

        ''' <summary>解析スタック状態。</summary>
        Private Structure StackState

            ''' <summary>開始ノード。</summary>
            Public ReadOnly Property FromNode As AnalysisNode

            ''' <summary>終了ノード。</summary>
            Public ReadOnly Property ToNode As AnalysisNode

            ''' <summary>ルート番号。</summary>
            Public ReadOnly Property Route As Integer

            ''' <summary>開始位置。</summary>
            Public ReadOnly Property StartPosition As Integer

            ''' <summary>解析結果。</summary>
            Public ReadOnly Property Answer As ABNFAnalysisItem

            ''' <summary>
            ''' コンストラクタ。
            ''' </summary>
            ''' <param name="fromNode">開始ノード。</param>
            ''' <param name="toNode">終了ノード。</param>
            ''' <param name="route">ルート番号。</param>
            ''' <param name="startPosition">開始位置。</param>
            ''' <param name="answer">解析結果。</param>
            Public Sub New(fromNode As AnalysisNode,
                           toNode As AnalysisNode,
                           route As Integer,
                           startPosition As Integer,
                           answer As ABNFAnalysisItem)
                Me.FromNode = fromNode
                Me.ToNode = toNode
                Me.Route = route
                Me.StartPosition = startPosition
                Me.Answer = answer
            End Sub

            ''' <summary>
            ''' 文字列表現を取得する。
            ''' </summary>
            ''' <returns>文字列表現。</returns>
            Overrides Function ToString() As String
                Return $"From:{Me.FromNode.Id}, To:{Me.ToNode.Id}, Route:{Me.Route}, Start:{Me.StartPosition}"
            End Function

        End Structure

        ''' <summary>追跡状態。</summary>
        Private Structure TrackingState

            ''' <summary>バックトラックアクションを取得します。</summary>
            Public ReadOnly Property Action As BacktrackAction

            ''' <summary>対象ノードを取得します。</summary>
            Public ReadOnly Property Node As AnalysisNode

            ''' <summary>対象ルート番号を取得します。</summary>
            Public ReadOnly Property Route As Integer

            ''' <summary>対象位置を取得します。</summary>
            Public ReadOnly Property Position As Integer

            ''' <summary>
            ''' コンストラクタ。
            ''' </summary>
            ''' <param name="action">アクション。</param>
            ''' <param name="node">対象ノード。</param>
            ''' <param name="route">対象ルート番号。</param>
            ''' <param name="position">対象位置。</param>
            Public Sub New(action As BacktrackAction,
                           node As AnalysisNode,
                           route As Integer,
                           position As Integer)
                Me.Action = action
                Me.Node = node
                Me.Route = route
                Me.Position = position
            End Sub

        End Structure

    End Class

End Namespace
