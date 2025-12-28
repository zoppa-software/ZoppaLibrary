Option Explicit On
Option Strict On

Imports System.Text
Imports ZoppaLibrary.ABNF.NumValAnalysis
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.EBNF

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
            If TypeOf range.Expr Is RuleNameExpression Then
                Return New RuleNameNode(id, range)
            Else
                Return New AnalysisNode(id, range)
            End If
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

        Public Function Match(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()

            ' 解析を実行
            Dim startPos = tr.Position
            If TypeOf Me.Range.Expr Is RuleNameExpression Then
                ' サブルールの解析を実行
                Dim res = Me.MatchExpression(tr, env)
                If res.success Then
                    Return res
                End If

            ElseIf TypeOf Me.Range.Expr Is NumValExpression Then
                ' 数値の種類を判定
                Dim numType As NumType
                Select Case Me.Range.SubRanges(0).Expr.GetType()
                    Case GetType(NumValExpression.Range)
                        numType = NumType.Range
                    Case GetType(NumValExpression.Concat)
                        numType = NumType.Concat
                    Case Else
                        numType = NumType.One
                End Select

                ' 数値の配列を取得
                Dim list As New List(Of UInteger)()
                Select Case Me.Range.SubChar(0)
                    Case "x"c
                        ' 16進数
                        For Each rng In Me.Range.SubRanges
                            list.Add(Convert.ToUInt32(rng.ToString(), 16))
                        Next
                    Case "b"c
                        ' 2進数
                        For Each rng In Me.Range.SubRanges
                            list.Add(Convert.ToUInt32(rng.ToString(), 2))
                        Next
                    Case Else
                        ' 10進数
                        For Each rng In Me.Range.SubRanges
                            list.Add(Convert.ToUInt32(rng.ToString(), 10))
                        Next
                End Select
                Dim numValues = list.ToArray()

                ' 数値判定
                Select Case numType
                    Case NumType.One
                        ' 単一の数値
                        Dim readByte = tr.Read()
                        If readByte = numValues(0) Then
                            Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
                        End If
                    Case NumType.Range
                        ' 範囲指定の数値
                        Dim readByte = tr.Read()
                        If readByte >= numValues(0) AndAlso readByte <= numValues(1) Then
                            Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
                        End If
                    Case NumType.Concat
                        ' 連結指定の数値
                        Dim initialPos = tr.Position
                        Dim success As Boolean = True
                        For Each val As UInteger In numValues
                            Dim readByte = tr.Read()
                            If readByte <> val Then
                                success = False
                                Exit For
                            End If
                        Next
                        If success Then
                            Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
                        End If
                End Select


            ElseIf TypeOf Me.Range.Expr Is CharValExpression Then
                ' 文字列判定
                Dim strValue = Encoding.UTF8.GetBytes(Me.Range.SubRanges(0).ToString())
                Dim buf = New Byte(strValue.Length - 1) {}
                tr.Read(buf, 0, buf.Length)
                If EqualBytes(strValue, buf) Then
                    Return (True, New ABNFAnalysisItem("char-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
                End If

            ElseIf Me.Range.Expr Is Nothing Then
                ' 空文字列
                Return (True, Nothing)
            Else
                Throw New NotImplementedException()
            End If

            snapPos.Restore()
            Return (False, Nothing)
        End Function

        'Public Function Match(tr As PositionAdjustBytes,
        '                      env As ABNFEnvironment,
        '                      route As Integer,
        '                      arrived As SortedDictionary(Of Integer, Integer)) As (success As Boolean, answer As ABNFAnalysisItem, nextNode As AnalysisNode)
        '    If route < Me.Routes.Count Then

        '    End If

        '    Dim nextNode = Me.Routes(route).NextNode
        '    Dim fromArrived = GetArrivedCount(arrived, Me.Id)
        '    Dim toArrived = GetArrivedCount(arrived, nextNode.Id)
        '    Dim minLmt = Me.Routes(route).RequiredVisits
        '    Dim maxLmt = Me.Routes(route).LimitedVisits

        '    ' 開始ノードは常に次のノードへ進む
        '    If nextNode.Id = 0 Then
        '        Return (False, Nothing, Nothing)
        '    End If

        '    ' 最小訪問回数に達している場合は成功
        '    If fromArrived < minLmt Then
        '        Return (False, Nothing, Nothing)
        '    End If

        '    ' 訪問回数が上限を超えている場合は失敗
        '    If toArrived >= maxLmt Then
        '        Return (False, Nothing, Nothing)
        '    End If

        '    ' 解析を実行
        '    Dim startPos = tr.Position
        '    If TypeOf nextNode.Range.Expr Is RuleNameExpression Then
        '        'Dim ident = nextNode.Range.ToString()
        '        'Dim iter = env.RuleTable(ident).GetMatcher()
        '        'Dim res = iter.MoveNext(tr, env)
        '        'If res.success Then
        '        '    Return (True, New ABNFAnalysisItem(ident, iter.GetAnswer(), tr, startPos, tr.Position), nextNode)
        '        'Else
        '        '    Return (False, Nothing, Nothing)
        '        'End If
        '        'Return nextNode.MatchExpression(tr, env, route, nextNode)

        '    ElseIf TypeOf nextNode.Range.Expr Is NumValExpression Then
        '        ' 数値の種類を判定
        '        Dim numType As NumType
        '        Select Case nextNode.Range.SubRanges(0).Expr.GetType()
        '            Case GetType(NumValExpression.Range)
        '                numType = NumType.Range
        '            Case GetType(NumValExpression.Concat)
        '                numType = NumType.Concat
        '            Case Else
        '                numType = NumType.One
        '        End Select

        '        ' 数値の配列を取得
        '        Dim list As New List(Of UInteger)()
        '        Select Case nextNode.Range.SubChar(0)
        '            Case "x"c
        '                ' 16進数
        '                For Each rng In nextNode.Range.SubRanges
        '                    list.Add(Convert.ToUInt32(rng.ToString(), 16))
        '                Next
        '            Case "b"c
        '                ' 2進数
        '                For Each rng In nextNode.Range.SubRanges
        '                    list.Add(Convert.ToUInt32(rng.ToString(), 2))
        '                Next
        '            Case Else
        '                ' 10進数
        '                For Each rng In nextNode.Range.SubRanges
        '                    list.Add(Convert.ToUInt32(rng.ToString(), 10))
        '                Next
        '        End Select
        '        Dim numValues = list.ToArray()

        '        ' 数値判定
        '        Select Case numType
        '            Case NumType.One
        '                ' 単一の数値
        '                Dim readByte = tr.Read()
        '                If readByte = numValues(0) Then
        '                    Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position), nextNode)
        '                End If
        '            Case NumType.Range
        '                ' 範囲指定の数値
        '                Dim readByte = tr.Read()
        '                If readByte >= numValues(0) AndAlso readByte <= numValues(1) Then
        '                    Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position), nextNode)
        '                End If
        '            Case NumType.Concat
        '                ' 連結指定の数値
        '                Dim initialPos = tr.Position
        '                Dim success As Boolean = True
        '                For Each val As UInteger In numValues
        '                    Dim readByte = tr.Read()
        '                    If readByte <> val Then
        '                        success = False
        '                        Exit For
        '                    End If
        '                Next
        '                If success Then
        '                    Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position), nextNode)
        '                End If
        '        End Select
        '        Return (False, Nothing, Nothing)

        '    ElseIf TypeOf nextNode.Range.Expr Is CharValExpression Then
        '        Dim strValue = Encoding.UTF8.GetBytes(nextNode.Range.SubRanges(0).ToString())
        '        Dim buf = New Byte(strValue.Length - 1) {}
        '        tr.Read(buf, 0, buf.Length)
        '        If EqualBytes(strValue, buf) Then
        '            Return (True, New ABNFAnalysisItem("char-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position), nextNode)
        '        Else
        '            Return (False, Nothing, Nothing)
        '        End If
        '    ElseIf nextNode.Range.Expr Is Nothing Then
        '        ' 空文字列
        '        Return (True, New ABNFAnalysisItem("nil", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position), nextNode)
        '    Else
        '        Throw New NotImplementedException()
        '    End If
        'End Function

        Public Overridable Function MatchExpression(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)
            Return (False, Nothing)
        End Function

        Public Overridable Function MoveNext(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, isRetry As Boolean, answer As ABNFAnalysisItem)
            Return (False, False, Nothing)
        End Function

        ''' <summary>
        ''' 指定したノードの訪問回数を取得する。
        ''' </summary>
        ''' <param name="id">ノードID。</param>
        ''' <param name="arrived">訪問回数リスト。</param>
        ''' <returns>到達回数。</returns>
        Protected Shared Function GetArrivedCount(arrived As SortedDictionary(Of Integer, Integer), id As Integer) As Integer
            Return If(arrived.ContainsKey(id), arrived(id), 0)
        End Function

        Protected Shared Function EqualBytes(a As Byte(), b As Byte()) As Boolean
            If a.Length <> b.Length Then
                Return False
            End If
            For i As Integer = 0 To a.Length - 1
                If a(i) <> b(i) Then
                    Return False
                End If
            Next
            Return True
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

        Public NotInheritable Class RuleNameNode
            Inherits AnalysisNode

            Private _matchers As New SortedDictionary(Of Integer, RuleAnalysis.AnalysisMatcher)()

            'Private _hasNext As New SortedDictionary(Of Integer, Boolean)()

            Private ReadOnly _ident As String

            Public Sub New(id As Integer, range As ExpressionRange)
                MyBase.New(id, range)
                Me._ident = range.ToString()
            End Sub

            'Public Overrides Function HasNext(position As Integer, route As Integer) As Boolean
            '    If Me._hasNext.ContainsKey(position) Then
            '        If Me._hasNext(position) Then
            '            Return True
            '        Else
            '            Me._hasNext.Remove(position)
            '            Return False
            '        End If
            '    Else
            '        Me._hasNext.Add(position, True)
            '        Return True
            '    End If
            'End Function

            Public Overrides Function MatchExpression(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)
                Dim position = tr.Position
                Dim iterator As RuleAnalysis.AnalysisMatcher
                If Me._matchers.ContainsKey(position) Then
                    iterator = Me._matchers(position)
                Else
                    iterator = env.RuleTable(Me._ident).GetMatcher()
                    Me._matchers.Add(position, iterator)
                End If

                Dim res = iterator.Match(tr, env)
                If res.success Then
                    Return (True, New ABNFAnalysisItem(Me._ident, iterator.GetAnswer(), tr, position, tr.Position))
                Else
                    Me._matchers.Remove(position)
                    Return (False, Nothing)
                End If
            End Function

            Public Overrides Function MoveNext(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, isRetry As Boolean, answer As ABNFAnalysisItem)
                Dim position = tr.Position
                Dim iterator As RuleAnalysis.AnalysisMatcher
                If Me._matchers.ContainsKey(position) Then
                    iterator = Me._matchers(position)
                Else
                    iterator = env.RuleTable(Me._ident).GetMatcher()
                    Me._matchers.Add(position, iterator)
                End If

                Dim res = iterator.MoveNext(tr, env)
                If res.success Then
                    Return (True, True, New ABNFAnalysisItem(Me._ident, iterator.GetAnswer(), tr, position, tr.Position))
                Else
                    Me._matchers.Remove(position)
                    Return (False, True, Nothing)
                End If
            End Function

        End Class

    End Class

End Namespace
