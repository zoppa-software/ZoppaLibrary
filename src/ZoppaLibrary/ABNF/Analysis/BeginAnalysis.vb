Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 端点解析（開始）を表します。
    ''' </summary>
    Public NotInheritable Class BeginAnalysis
        Implements IAnalysis

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        ''' <returns>解析パターン。</returns>
        Public ReadOnly Property Pattern As List(Of IAnalysis.Link) Implements IAnalysis.Pattern

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(range As ExpressionRange)
            Me.Pattern = New List(Of IAnalysis.Link)()
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
                              env As ABNFEnvironment,
                              ruleTable As SortedDictionary(Of String, RuleAnalysis),
                              ruleName As String,
                              answers As List(Of ABNFAnalysisItem)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            Return (True, 0)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "<開始>"
        End Function

    End Class

End Namespace
