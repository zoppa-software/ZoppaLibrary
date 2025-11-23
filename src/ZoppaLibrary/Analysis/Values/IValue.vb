Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 値を表すインターフェイスです。
    ''' このインターフェイスは、値の型と値自体を取得するためのプロパティを定義します。
    ''' </summary>
    ''' <remarks>
    ''' このインターフェイスは、数値、文字列、真偽値などの異なる型の値を表現するために使用されます。
    ''' </remarks>
    Public Interface IValue

        ''' <summary>
        ''' 値の型を取得します。
        ''' </summary>
        ''' <returns>値の型。</returns>
        ''' <remarks>
        ''' このプロパティは、値の型を示すValueType列挙体を返します。
        ''' </remarks>
        ReadOnly Property Type As ValueType

    End Interface

    ''' <summary>
    ''' 未定義値を表す構造体です。
    ''' この構造体は、値が未定義であることを示し、IValueインターフェイスを実装します。
    ''' </summary>
    ''' <remarks>
    ''' この構造体は、値が存在しないことを示すために使用されます。
    ''' </remarks>
    NotInheritable Class NullValue
        Implements IValue

        ' NullValueのインスタンスをLazyに生成するためのフィールド
        Private Shared ReadOnly _instanse As New Lazy(Of NullValue)(Function() New NullValue())

        ''' <summary>
        ''' 値を取得します。
        ''' このプロパティは、NullValueのインスタンスを返します。
        ''' NullValueは、値が未定義であることを示すために使用されます。
        ''' </summary>
        ''' <returns>Null値。</returns>
        Public Shared ReadOnly Property Value As NullValue
            Get
                Return _instanse.Value
            End Get
        End Property

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Public ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.Null
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 数値を表す構造体です。
    ''' この構造体は、数値の値を保持し、IValueインターフェイスを実装します。
    ''' 数値の型はValueType.Numberとして定義されます。
    ''' </summary>
    NotInheritable Class NumberValue
        Implements IValue

        ''' <summary>値を取得します。</summary>
        ''' <returns>値。</returns>
        Public ReadOnly Property Value As Double

        ''' <summary>数値のコンストラクタ。</summary>
        ''' <param name="value">数値の値。</param>
        Public Sub New(value As Double)
            Me.Value = value
        End Sub

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Public ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.Number
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 文字列を表す構造体です。
    ''' この構造体は、文字列の値を保持し、IValueインターフェイスを実装します。
    ''' 文字列の型はValueType.Stringとして定義されます。
    ''' </summary>
    NotInheritable Class StringValue
        Implements IValue

        ' 空の文字列値を表すLazyなインスタンス
        Private Shared ReadOnly _emptyValue As New Lazy(Of StringValue)(Function() New StringValue(U8String.Empty))

        ' 空の文字列値を表すLazyなインスタンス
        Private Shared ReadOnly _brValue As New Lazy(Of StringValue)(Function() New StringValue(U8String.NewString(Environment.NewLine)))

        ''' <summary>空の文字列値を取得します。</summary>
        ''' <returns>空の文字列値。</returns>
        ''' <remarks>
        ''' このプロパティは、空の文字列値を表すStringValueのインスタンスを返します。
        ''' </remarks>
        Shared ReadOnly Property Empty As StringValue
            Get
                Return _emptyValue.Value
            End Get
        End Property

        ''' <summary>改行文字列値を取得します。</summary>
        ''' <returns>改行文字列値。</returns>
        ''' <remarks>
        ''' このプロパティは、改行を表すStringValueのインスタンスを返します。
        ''' </remarks>
        Shared ReadOnly Property Br As StringValue
            Get
                Return _brValue.Value
            End Get
        End Property

        ''' <summary>値を取得します。</summary>
        ''' <returns>値。</returns>
        Public ReadOnly Property Value As U8String

        ''' <summary>文字列のコンストラクタ。</summary>
        ''' <param name="value">文字列。</param>
        Public Sub New(value As U8String)
            Me.Value = value
        End Sub

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Public ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.Str
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 真偽値を表す構造体です。
    ''' この構造体は、真偽値の値を保持し、IValueインターフェイスを実装します。
    ''' 真偽値の型はValueType.Boolとして定義されます。
    ''' </summary>
    NotInheritable Class BooleanValue
        Implements IValue

        ' 真の値を表すLazyなインスタンス
        Private Shared ReadOnly _trueValue As New Lazy(Of BooleanValue)(Function() New BooleanValue(True))

        ' 偽の値を表すLazyなインスタンス
        Private Shared ReadOnly _falseValue As New Lazy(Of BooleanValue)(Function() New BooleanValue(False))

        ''' <summary>真の値を取得します。</summary>
        ''' <returns>真の値。</returns>
        Public Shared ReadOnly Property TrueValue As BooleanValue
            Get
                Return _trueValue.Value
            End Get
        End Property

        ''' <summary>偽の値を取得します。</summary>
        ''' <returns>偽の値。</returns>
        Public Shared ReadOnly Property FalseValue As BooleanValue
            Get
                Return _falseValue.Value
            End Get
        End Property

        ''' <summary>値を取得します。</summary>
        ''' <returns>値。</returns>
        Public ReadOnly Property Value As Boolean

        ''' <summary>真偽値のコンストラクタ。</summary>
        ''' <param name="value">数値の値。</param>
        Private Sub New(value As Boolean)
            Me.Value = value
        End Sub

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Public ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.Bool
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 配列を表す構造体です。
    ''' </summary>
    NotInheritable Class ArrayValue
        Implements IValue

        ''' <summary>値を取得します。</summary>
        ''' <returns>値。</returns>
        Public ReadOnly Property Value As IValue()

        ''' <summary>配列値のコンストラクタ。</summary>
        ''' <param name="value">配列の値リスト。</param>
        Public Sub New(value As IValue())
            Me.Value = value
        End Sub

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Public ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.Array
            End Get
        End Property

    End Class

    ''' <summary>
    ''' オブジェクト値を表す構造体です。
    ''' この構造体は、オブジェクトの値を表現し、式として評価するためのメソッドを提供します。
    ''' </summary>
    ''' <remarks>
    ''' この構造体は、オブジェクトのプロパティやメソッドを含む値を表現するために使用されます。
    ''' </remarks>
    NotInheritable Class ObjectValue
        Implements IValue

        ''' <summary>値を取得します。</summary>
        ''' <returns>値。</returns>
        Public ReadOnly Property Value As Object

        ''' <summary>オブジェクト値のコンストラクタ。</summary>
        ''' <param name="obj">対象となるオブジェクト。</param>
        ''' <remarks>
        ''' このコンストラクタは、オブジェクト値を初期化します。
        ''' </remarks>
        Public Sub New(obj As Object)
            Me.Value = obj
        End Sub

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Private ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.Obj
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 日付値を表す構造体です。
    ''' この構造体は、日付の値を保持し、IValueインターフェイスを実装します。
    ''' 日付の型はValueType.DateTimeとして定義されます。
    ''' </summary>
    NotInheritable Class DateTimeValue
        Implements IValue

        ''' <summary>値を取得します。</summary>
        ''' <returns>値。</returns>
        Public ReadOnly Property Value As DateTime

        ''' <summary>日付値のコンストラクタ。</summary>
        ''' <param name="value">日付の値。</param>
        Public Sub New(value As DateTime)
            Me.Value = value
        End Sub

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Public ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.DateTime
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 時間値を表す構造体です。
    ''' この構造体は、時間の値を保持し、IValueインターフェイスを実装します。
    ''' 時間の型はValueType.TimeSpanとして定義されます。
    ''' </summary>
    NotInheritable Class TimeSpanValue
        Implements IValue

        ''' <summary>値を取得します。</summary>
        ''' <returns>値。</returns>
        Public ReadOnly Property Value As TimeSpan

        ''' <summary>時間値のコンストラクタ。</summary>
        ''' <param name="value">時間の値。</param>
        Public Sub New(value As TimeSpan)
            Me.Value = value
        End Sub

        ''' <summary>値の型を取得します。</summary>
        ''' <returns>値の型。</returns>
        Public ReadOnly Property Type As ValueType Implements IValue.Type
            Get
                Return ValueType.TimeSpan
            End Get
        End Property

    End Class

End Namespace
