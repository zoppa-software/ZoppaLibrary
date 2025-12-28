Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    Public Class AnalysisMatcher
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
                        If Not preview.Item1.IsRetry Then
                            route = preview.Item2 + 1
                            Exit Do
                        End If
                        'If TypeOf preview.Item1 IsNot RuleNameNode Then
                        '    route = preview.Item2 + 1
                        '    Exit Do
                        'End If
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

End Namespace
