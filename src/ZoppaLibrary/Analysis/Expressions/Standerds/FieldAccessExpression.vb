Option Strict On
Option Explicit On

Imports System.Reflection
Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' フィールドアクセスを表す式の構造体です。
    ''' この構造体は、オブジェクトのプロパティにアクセスするために使用されます。
    ''' </summary>
    ''' <remarks>
    ''' この式は、オブジェクトのプロパティにアクセスし、その値を取得します。
    ''' </remarks>
    NotInheritable Class FieldAccessExpression
        Implements IExpression

        ' 変数
        Private ReadOnly _target As IExpression

        ' プロパティ名リスト
        Private ReadOnly _propertyName As String

        ''' <summary>フィールドアクセス式のコンストラクタ。</summary>
        ''' <param name="target">アクセスするインスタンス。</param>
        ''' <param name="propertyName">プロパティ名リスト。</param>
        Public Sub New(target As IExpression, propertyName As U8String)
            If target Is Nothing Then
                Throw New ArgumentNullException(NameOf(target))
            End If
            _target = target
            _propertyName = propertyName.ToString()
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.FieldAccessExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>フィールドアクセスの結果としての値。</returns>
        ''' <remarks>
        ''' このメソッドは、変数名とインデックスを使用して配列の要素にアクセスし、その値を返します。
        ''' </remarks>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            ' 変数を取得
            Dim target = _target.GetValue(venv).Obj
            If target Is Nothing Then
                Return NullValue.Value
            End If

            If TypeOf target Is DynamicObject Then
                ' DynamicObjectの場合、プロパティ名を使用して値を取得
                Dim obj = DirectCast(target, DynamicObject)(_propertyName)

                ' 最終的な値をIValueに変換して返す
                If TypeOf obj Is IValue Then
                    Return DirectCast(obj, IValue)
                End If
                Return ConvertToValue(obj)
            Else
                ' プロパティ名を使用してオブジェクトのプロパティにアクセス
                Dim propInfo As PropertyInfo = target.GetType().GetProperty(_propertyName)
                If propInfo Is Nothing Then
                    Throw New InvalidOperationException($"プロパティ '{_propertyName}' が見つかりません。")
                End If

                ' プロパティの値を取得
                Dim obj = propInfo.GetValue(target, Nothing)

                ' 最終的な値をIValueに変換して返す
                If TypeOf obj Is IValue Then
                    Return DirectCast(obj, IValue)
                End If
                Return ConvertToValue(obj)
            End If
        End Function

    End Class

End Namespace
