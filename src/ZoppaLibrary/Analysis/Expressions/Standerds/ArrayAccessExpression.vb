Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 配列アクセス式を表す構造体です。
    ''' この構造体は、配列の要素にアクセスするための式を表現します。
    ''' </summary>
    ''' <remarks>
    ''' この式は、変数名とインデックスを使用して配列の要素にアクセスします。
    ''' </remarks>
    NotInheritable Class ArrayAccessExpression
        Implements IExpression

        ' 変数
        Private ReadOnly _target As IExpression

        ' インデックス
        Private ReadOnly _index As IExpression

        ''' <summary>配列アクセス式のコンストラクタ。</summary>
        ''' <param name="target">アクセスする変数名。</param>
        ''' <param name="index">インデックス式。</param>
        ''' <remarks>
        ''' このコンストラクタは、配列アクセスのためのターゲット変数とインデックスを初期化します。
        ''' </remarks>
        Public Sub New(target As IExpression, index As IExpression)
            If target Is Nothing Then
                Throw New ArgumentNullException(NameOf(target))
            End If
            If index Is Nothing Then
                Throw New ArgumentNullException(NameOf(index))
            End If
            Me._target = target
            Me._index = index
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.ArrayAccessExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>配列アクセスの結果としての値。</returns>
        ''' <remarks>
        ''' このメソッドは、変数名とインデックスを使用して配列の要素にアクセスし、その値を返します。
        ''' </remarks>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            ' 変数とインデックスの値を取得
            Dim arr = _target.GetValue(venv)
            Dim idx = _index.GetValue(venv)

            If arr.Type = ValueType.Array AndAlso idx.Type = ValueType.Number Then
                ' 配列とインデックスが数値であることを確認
                Dim arrayValue As ArrayValue = DirectCast(arr, ArrayValue)
                Dim index As Integer = CInt(Math.Floor(idx.Number))

                ' インデックスが配列の範囲内であることを確認して、要素を取得
                If index >= 0 AndAlso index < arrayValue.Array.Length Then
                    Return arrayValue.Array(index)
                Else
                    Throw New IndexOutOfRangeException("配列のインデックスが範囲外です。")
                End If
            Else
                Throw New InvalidOperationException("配列アクセスは、配列と数値の型でなければなりません。")
            End If
        End Function

    End Class

End Namespace
