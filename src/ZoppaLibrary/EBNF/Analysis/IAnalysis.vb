Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 解析パターンを表すインターフェイス。
    ''' </summary>
    Public Interface IAnalysis

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        ReadOnly Property Pattern As List(Of IAnalysis)

        ''' <summary>
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。入力ストリームの現在位置を管理します。</param>
        ''' <param name="env">解析環境。ルールテーブルとメソッドテーブルを含みます。</param>
        ''' <param name="ruleTable">ルール名とその解析ロジックのマッピング。</param>
        ''' <param name="specialMethods">特殊構文を処理するカスタムメソッドのテーブル。</param>
        ''' <param name="ruleName">現在評価中のルール名。デバッグ用。</param>
        ''' <param name="answers">解析結果を格納するリスト。マッチした要素が追加されます。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Function Match(tr As IPositionAdjustReader,
                       env As EBNFEnvironment,
                       ruleTable As SortedDictionary(Of String, OldRuleAnalysis),
                       specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                       ruleName As String,
                       answers As List(Of EBNFAnalysisItem)) As (sccess As Boolean, shift As Integer)

    End Interface

End Namespace
