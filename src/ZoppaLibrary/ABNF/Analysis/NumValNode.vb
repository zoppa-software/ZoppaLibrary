Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 数値ノード。
    ''' </summary>
    NotInheritable Class NumValNode
        Inherits AnalysisNode

        ''' <summary>数値の種類を表します。</summary>
        Public Enum NumType
            ''' <summary>単一の数値。</summary>
            One = 1
            ''' <summary>範囲指定の数値。</summary>
            Range = 2
            ''' <summary>連結指定の数値。</summary>
            Concat = 3
        End Enum

        ''' <summary>
        ''' 数値の種類を表します。
        ''' </summary>
        Private ReadOnly _numType As NumType

        ''' <summary>
        ''' 数値の配列。
        ''' </summary>
        Private ReadOnly _numValues As UInteger()

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ノードID。</param>
        ''' <param name="range">式範囲。</param>
        Public Sub New(id As Integer, range As ExpressionRange)
            MyBase.New(id, range)

            ' 数値の種類を判定
            Select Case range.SubRanges(0).Expr.GetType()
                Case GetType(NumValExpression.Range)
                    Me._numType = NumType.Range
                Case GetType(NumValExpression.Concat)
                    Me._numType = NumType.Concat
                Case Else
                    Me._numType = NumType.One
            End Select

            ' 数値の配列を取得
            Dim list As New List(Of UInteger)()
            Select Case range.SubChar(0)
                Case "x"c
                    ' 16進数
                    For Each rng In range.SubRanges
                        list.Add(Convert.ToUInt32(rng.ToString(), 16))
                    Next
                Case "b"c
                    ' 2進数
                    For Each rng In range.SubRanges
                        list.Add(Convert.ToUInt32(rng.ToString(), 2))
                    Next
                Case Else
                    ' 10進数
                    For Each rng In range.SubRanges
                        list.Add(Convert.ToUInt32(rng.ToString(), 10))
                    Next
            End Select
            Me._numValues = list.ToArray()
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param
        ''' <returns>マッチ結果。</returns>
        Public Overrides Function Match(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()
            Dim startPos = tr.Position

            ' 数値判定
            Select Case Me._numType
                Case NumType.One
                    ' 単一の数値
                    Dim readByte = tr.Read()
                    If readByte = Me._numValues(0) Then
                        Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
                    End If

                Case NumType.Range
                    ' 範囲指定の数値
                    Dim readByte = tr.Read()
                    If readByte >= Me._numValues(0) AndAlso readByte <= Me._numValues(1) Then
                        Return (True, New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
                    End If

                Case NumType.Concat
                    ' 連結指定の数値
                    Dim initialPos = tr.Position
                    Dim success As Boolean = True
                    For Each val As UInteger In Me._numValues
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

            ' 一致しない場合は偽を返す
            snapPos.Restore()
            Return (False, Nothing)
        End Function

    End Class

End Namespace
