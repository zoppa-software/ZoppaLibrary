Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 分析処理中に発生する例外を表すクラスです。
    ''' このクラスは、分析処理に関連するエラーを示すために使用されます。
    ''' </summary>
    ''' <remarks>
    ''' この例外は、分析処理中に特定の条件が満たされない場合や、無効なデータが検出された場合にスローされます。
    ''' </remarks>
    Public NotInheritable Class AnalysisException
        Inherits Exception

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

    End Class

End Namespace