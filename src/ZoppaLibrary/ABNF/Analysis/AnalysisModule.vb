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
        ''' <param name="counter">訪問回数カウンター。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        <Extension()>
        Public Function AnalysisNextPattern(analysis As IAnalysis,
                                            tr As PositionAdjustBytes,
                                            env As ABNFEnvironment,
                                            ruleTable As SortedDictionary(Of String, RuleAnalysis),
                                            ruleName As String,
                                            answers As List(Of ABNFAnalysisItem),
                                            counter As Dictionary(Of IAnalysis, Integer)) As (sccess As Boolean, shift As Integer)
            Dim shift As Integer = Integer.MaxValue

            '' パターンを順次評価する
            'For Each evalExpr In analysis.Pattern
            '    ' 訪問回数を取得する
            '    Dim arrival As Integer = 0
            '    If counter.ContainsKey(evalExpr.ToAnalysis) Then
            '        arrival = counter(evalExpr.ToAnalysis)
            '    Else
            '        counter.Add(evalExpr.ToAnalysis, 0)
            '    End If

            '    If arrival >= evalExpr.MinLimit AndAlso arrival < evalExpr.MaxLimit Then
            '        ' パターンを評価する
            '        counter(evalExpr.ToAnalysis) = arrival + 1
            '        Dim evalResult = evalExpr.ToAnalysis.Match(tr, env, ruleTable, ruleName, answers, counter)
            '        counter(evalExpr.ToAnalysis) = arrival

            '        ' 解析が成功した場合は真を返す
            '        If evalResult.sccess Then
            '            Return (True, 0)
            '        ElseIf evalResult.shift < shift Then
            '            shift = evalResult.shift
            '        End If
            '    End If
            'Next

            ' どれもマッチしなかった場合は偽を返す
            Return (False, 0)
        End Function

    End Module

End Namespace
