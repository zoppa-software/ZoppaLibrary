Option Explicit On
Option Strict On

Imports System.IO
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.EBNF

Namespace ABNF

    ''' <summary>
    ''' 数値値解析を表します。
    ''' </summary>
    NotInheritable Class NumValAnalysis
        Implements IAnalysis

        ''' <summary>数値の種類を表します。</summary>
        Public Enum NumType
            ''' <summary>単一の数値。</summary>
            One = 1
            ''' <summary>範囲指定の数値。</summary>
            Range = 2
            ''' <summary>連結指定の数値。</summary>
            Concat = 3
        End Enum

        ''' <summary>評価範囲。</summary>
        Private ReadOnly _range As ExpressionRange

        ''' <summary>数値の種類。</summary>
        Private ReadOnly _type As NumType

        ''' <summary>数値の配列。</summary>
        Private ReadOnly _values As UInteger()

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        Public ReadOnly Property Pattern As List(Of AnalysisRoute)

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(range As ExpressionRange)
            Me._range = range

            ' 数値の種類を判定
            Select Case range.SubRanges(0).Expr.GetType()
                Case GetType(NumValExpression.Range)
                    Me._type = NumType.Range
                Case GetType(NumValExpression.Concat)
                    Me._type = NumType.Concat
                Case Else
                    Me._type = NumType.One
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
            Me._values = list.ToArray()

            Me.Pattern = New List(Of AnalysisRoute)()
        End Sub

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
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of ABNFAnalysisItem)()

            ' 数値判定
            Select Case Me._type
                Case NumType.One
                    ' 単一の数値
                    Dim readByte = tr.Read()
                    If readByte = Me._values(0) Then
                        answers.Add(New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem)(), tr, startPos, tr.Position))
                        Return (True, tr.Position - startPos)
                    End If
                Case NumType.Range
                    ' 範囲指定の数値
                    Dim readByte = tr.Read()
                    If readByte >= Me._values(0) AndAlso readByte <= Me._values(1) Then
                        answers.Add(New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem)(), tr, startPos, tr.Position))
                        Return (True, tr.Position - startPos)
                    End If
                Case NumType.Concat
                    ' 連結指定の数値
                    Dim initialPos = tr.Position
                    Dim success As Boolean = True
                    For Each val As UInteger In Me._values
                        Dim readByte = tr.Read()
                        If readByte <> val Then
                            success = False
                            Exit For
                        End If
                    Next
                    If success Then
                        answers.Add(New ABNFAnalysisItem("num-val", New List(Of ABNFAnalysisItem)(), tr, startPos, tr.Position))
                        Return (True, tr.Position - startPos)
                    End If
            End Select

            snap.Restore()
            Return (False, 1)
        End Function
    End Class

End Namespace
