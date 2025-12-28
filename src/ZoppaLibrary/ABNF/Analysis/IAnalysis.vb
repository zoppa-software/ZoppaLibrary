Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 解析パターンを表すインターフェイス。
    ''' </summary>
    Public Interface IAnalysis

        ''' <summary>
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。入力ストリームの現在位置を管理します。</param>
        ''' <param name="env">解析環境。ルールテーブルとメソッドテーブルを含みます。</param>
        ''' <param name="ruleTable">ルール名とその解析ロジックのマッピング。</param>
        ''' <param name="ruleName">現在評価中のルール名。デバッグ用。</param>
        ''' <param name="answers">解析結果を格納するリスト。マッチした要素が追加されます。</param>
        ''' <param name="counter">訪問回数を示すカウンター。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Function Match(tr As PositionAdjustBytes,
                       env As ABNFEnvironment,
                       ruleTable As SortedDictionary(Of String, RuleAnalysis),
                       ruleName As String,
                       answers As List(Of ABNFAnalysisItem),
                       counter As Dictionary(Of IAnalysis, Integer)) As (sccess As Boolean, shift As Integer)

    End Interface

End Namespace
