Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Imports ZoppaLibrary.ABNF.AnalysisNode
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.EBNF

Namespace ABNF

    ''' <summary>
    ''' ABNF 解析マッチャー。
    ''' </summary>
    Public NotInheritable Class AnalysisMatcher

        ''' <summary>最大反復回数。</summary>
        Private Const MaxIterations As Integer = 10000

        ''' <summary>最大スタック深度。</summary>
        Private Const MaxStackDepth As Integer = 10000

        ''' <summary>実用的な量詞上限。</summary>
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
        Private ReadOnly _arrived As New SortedDictionary(Of Integer, Integer)()

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
                              env As ABNFEnvironment) As (success As Boolean, shift As Integer)
            If Me._stack.Count = 0 Then
                ' 初回開始
                Me._arrived.Clear()
                Me.IncrementArrived(Me._root.Id)
                Return Me.Tracking(Me._root, 0, tr, env)
            Else
                ' 継続解析
                Dim cur = Me._stack.Pop()
                Me.DecrementArrived(cur.ToNode.Id)
                tr.Seek(cur.StartPosition)
                Return Me.Tracking(cur.ToNode, cur.Route, tr, env)
            End If
        End Function

        ''' <summary>
        ''' 次の解析ステップを実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function MoveNext(tr As PositionAdjustBytes,
                                 env As ABNFEnvironment) As (success As Boolean, shift As Integer)
            If Me._stack.Count = 0 Then
                ' 初回開始
                Me._arrived.Clear()
                Me.IncrementArrived(Me._root.Id)
                Return Me.Tracking(Me._root, 0, tr, env)
            Else
                ' 継続解析
                Dim cur = Me._stack.Pop()
                Me.DecrementArrived(cur.ToNode.Id)
                tr.Seek(cur.StartPosition)
                Return Me.Tracking(cur.ToNode, cur.Route + 1, tr, env)
            End If
        End Function

        ''' <summary>
        ''' 解析を追跡する。
        ''' </summary>
        ''' <param name="node">現在のノード。</param>
        ''' <param name="route">現在のルート番号。</param>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Private Function Tracking(node As AnalysisNode,
                                  route As Integer,
                                  tr As PositionAdjustBytes,
                                  env As ABNFEnvironment) As (success As Boolean, shift As Integer)
            Dim startPosition = tr.Position
            Dim iterationCount As Integer = 0
            Dim currentPosition = tr.Position
            Dim action As BacktrackAction

            Do
                iterationCount += 1
                If iterationCount > MaxIterations Then
                    Throw New ABNFException($"解析が最大反復回数({MaxIterations})を超過しました。無限ループの可能性があります。")
                End If

                If Me._stack.Count > MaxStackDepth Then
                    Throw New ABNFException($"解析スタックが最大深度({MaxStackDepth})を超過しました。")
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
                        ' 次のノードへ進む
                        Me._stack.Push(New StackState(node, nextNode, route, currentPosition, tr.Position, matched.answer))
                        Me.IncrementArrived(nextNode.Id)

                        ' 最終ノードに到達した場合は成功
                        If nextNode.Routes.Count = 0 Then
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

                    ' 追跡状態を取得
                    action = backtrack.Action
                    node = backtrack.Node
                    route = backtrack.Route
                    currentPosition = backtrack.Position
                Loop While action = BacktrackAction.RetryMoveNext

            Loop While action = BacktrackAction.RetryMatch

            ' 全てのルートを試行したがマッチしなかった場合は失敗
            Return (False, 0)
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

                ' ひとつ前のノード、位置へ戻る
                Dim preview = Me._stack.Pop()
                Dim selectedNode = preview.FromNode
                Me.DecrementArrived(preview.ToNode.Id)
                tr.Seek(preview.startPosition)

                ' リトライを試みる
                Dim retry = preview.ToNode.MoveNext(tr, env)
                If retry.success Then
                    ' リトライ成功の場合はそのまま進む
                    Me._stack.Push(New StackState(preview.FromNode, preview.ToNode, preview.Route, preview.StartPosition, tr.Position, retry.answer))
                    Me.IncrementArrived(preview.ToNode.Id)
                    Return New TrackingState(BacktrackAction.RetryMatch, preview.ToNode, 0, tr.Position)

                ElseIf preview.Route + 1 < selectedNode.Routes.Count Then
                    ' 選択肢が存在する場合は次の選択肢へ
                    Return New TrackingState(BacktrackAction.RetryMatch, selectedNode, preview.Route + 1, preview.StartPosition)

                Else
                    ' 選択肢が存在しない、かつリトライ可能な場合はリトライへ
                    ' そうでない場合は終了
                    currentPosition = preview.StartPosition
                    If preview.ToNode.IsRetry Then
                        Return New TrackingState(BacktrackAction.RetryMoveNext, selectedNode, selectedRoute, tr.Position)
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
        ''' 指定ノードの訪問回数をインクリメントする。
        ''' </summary>
        ''' <param name="nodeId">ノードID。</param>
        Private Sub IncrementArrived(nodeId As Integer)
            If Me._arrived.ContainsKey(nodeId) Then
                Me._arrived(nodeId) += 1
            Else
                Me._arrived.Add(nodeId, 1)
            End If
        End Sub

        ''' <summary>
        ''' 指定ノードの訪問回数をデクリメントする。
        ''' </summary>
        ''' <param name="nodeId">ノードID。</param>
        Private Sub DecrementArrived(nodeId As Integer)
            Me._arrived(nodeId) -= 1
            If Me._arrived(nodeId) <= 0 Then
                Me._arrived.Remove(nodeId)
            End If
        End Sub

        ''' <summary>
        ''' 指定ノードの訪問回数を取得する。
        ''' </summary>
        ''' <param name="nodeId">ノードID。</param>
        ''' <returns>訪問回数。</returns>
        Private Function GetArrived(nodeId As Integer) As Integer
            Return If(Me._arrived.ContainsKey(nodeId), Me._arrived(nodeId), 0)
        End Function

        ''' <summary>
        ''' 解析結果を取得する。
        ''' </summary>
        ''' <returns>解析結果リスト。</returns>
        Public Function GetAnswer() As List(Of ABNFAnalysisItem)
            Dim res As New List(Of ABNFAnalysisItem)()
            For Each item In Me._stack
                If item.answer IsNot Nothing Then
                    res.Add(item.answer)
                End If
            Next
            res.Reverse()
            Return res
        End Function

        Friend Sub ClearCache()
            Dim idHash As New HashSet(Of Integer)()
            Me._root.ClearCache(idHash)
        End Sub

        Private Structure StackState
            Public ReadOnly Property FromNode As AnalysisNode
            Public ReadOnly Property ToNode As AnalysisNode
            Public ReadOnly Property Route As Integer
            Public ReadOnly Property StartPosition As Integer
            Public ReadOnly Property Answer As ABNFAnalysisItem
            Public Sub New(fromNode As AnalysisNode,
                           toNode As AnalysisNode,
                           route As Integer,
                           startPosition As Integer,
                           endPosition As Integer,
                           answer As ABNFAnalysisItem)
                Me.FromNode = fromNode
                Me.ToNode = toNode
                Me.Route = route
                Me.StartPosition = startPosition
                Me.Answer = answer
            End Sub
            Overrides Function ToString() As String
                Return $"From:{Me.FromNode.Id}, To:{Me.ToNode.Id}, Route:{Me.Route}, Start:{Me.StartPosition}"
            End Function
        End Structure

        ''' <summary>追跡状態。</summary>
        Private Structure TrackingState

            ''' <summary>対象のノードを取得します。</summary>
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
