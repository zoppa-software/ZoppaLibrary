Option Explicit On
Option Strict On

Imports System.Runtime.CompilerServices
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 解析モジュール。
    ''' </summary>
    Module AnalysisModule

        ''' <summary>
        ''' 解析パターンに紐付くパターンを順次評価する。
        ''' </summary>
        ''' <param name="analysis">解析パターン。</param>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <param name="ruleTable">ルール解析テーブル。</param>
        ''' <param name="specialMethods">特殊メソッドテーブル。</param>
        ''' <param name="ruleName">現在のルール名。</param>
        ''' <param name="answers">解析結果のリスト。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        <Extension()>
        Public Function AnalysisNextPattern(analysis As IAnalysis,
                                            tr As IPositionAdjustReader,
                                            env As ABNFEnvironment,
                                            ruleTable As SortedDictionary(Of String, RuleAnalysis),
                                            ruleName As String,
                                            answers As List(Of ABNFAnalysisItem)) As (sccess As Boolean, shift As Integer)
            Dim shift As Integer = Integer.MaxValue

            ' パターンを順次評価する
            For Each evalExpr In analysis.Pattern
                ' パターンを評価する
                Dim evalResult = evalExpr.ToAnalysis.Match(tr, env, ruleTable, ruleName, answers)

                ' 解析が成功した場合は真を返す
                If evalResult.sccess Then
                    Return (True, 0)
                ElseIf evalResult.shift < shift Then
                    shift = evalResult.shift
                End If
            Next

            ' どれもマッチしなかった場合は偽を返す
            Return (False, shift)
        End Function

    End Module

End Namespace
