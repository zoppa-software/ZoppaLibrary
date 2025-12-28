Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 解析ルートを表します。
    ''' </summary>
    Public NotInheritable Class AnalysisRoute
        Implements IAnalysis

        ''' <summary>到達先解析ノード。</summary>
        Public ReadOnly Property ToAnalysis As IAnalysis

        ''' <summary>最小出現回数。</summary>
        Public ReadOnly Property MinLimit As Integer

        ''' <summary>最大出現回数。</summary>
        Public ReadOnly Property MaxLimit As Integer

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        Public ReadOnly Property Pattern As List(Of AnalysisRoute)
            Get
                Return Nothing
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="toNode">到達先解析ノード。</param>
        ''' <param name="minLimit">最小出現回数。</param>
        ''' <param name="maxLimit">最大出現回数。</param>
        Public Sub New(toNode As IAnalysis, minLimit As Integer, maxLimit As Integer)
            Me.ToAnalysis = toNode
            Me.MinLimit = minLimit
            Me.MaxLimit = maxLimit
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
            Return Me.ToAnalysis.Match(tr, env, ruleTable, ruleName, answers, counter)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"Min:{Me.MinLimit}, Max:{Me.MaxLimit} -> {Me.ToAnalysis.ToString()}"
        End Function

    End Class

End Namespace