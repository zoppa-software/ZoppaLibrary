Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' IVariableインターフェイスは、変数の基本的なプロパティを定義します。
    ''' 変数は数値、文字列、または真偽値のいずれかを表すことができます。
    ''' </summary>
    Public Interface IVariable

        ''' <summary>変数の型を取得します。</summary>
        ''' <returns>変数の型。</returns>
        ReadOnly Property Type As VariableType

    End Interface

    ''' <summary>
    ''' 変数の式を表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、変数の式を表現します。
    ''' </summary>
    NotInheritable Class ExprVariable
        Implements IVariable

        ''' <summary>変数の式を取得します。</summary>
        Public ReadOnly Value As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="expr">式。</param>
        Public Sub New(value As IExpression)
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Expr
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 変数の数値を表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、数値型の変数を表現します。
    ''' </summary>
    NotInheritable Class NumberVariable
        Implements IVariable

        ''' <summary>変数の数値を取得します。</summary>
        Public ReadOnly Value As Double

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Public Sub New(value As Double)
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        ''' <remarks>この構造体は数値型の変数を表します。</remarks>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Number
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 変数の文字列を表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、文字列型の変数を表現します。
    ''' </summary>
    NotInheritable Class StringVariable
        Implements IVariable

        ''' <summary>変数の文字列を取得します。</summary>
        Public ReadOnly Property Value As U8String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Public Sub New(value As U8String)
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        ''' <remarks>この構造体は文字列型の変数を表します。</remarks>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Str
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 変数の真偽値を表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、真偽値型の変数を表現します。
    ''' </summary>
    NotInheritable Class BooleanVariable
        Implements IVariable

        ' 真の値を表すLazyなインスタンス
        Private Shared ReadOnly _trueValue As New Lazy(Of BooleanVariable)(Function() New BooleanVariable(True))

        ' 偽の値を表すLazyなインスタンス
        Private Shared ReadOnly _falseValue As New Lazy(Of BooleanVariable)(Function() New BooleanVariable(False))

        ''' <summary>真の値を取得します。</summary>
        ''' <returns>真の値。</returns>
        Public Shared ReadOnly Property TrueValue As BooleanVariable
            Get
                Return _trueValue.Value
            End Get
        End Property

        ''' <summary>偽の値を取得します。</summary>
        ''' <returns>偽の値。</returns>
        Public Shared ReadOnly Property FalseValue As BooleanVariable
            Get
                Return _falseValue.Value
            End Get
        End Property

        ''' <summary>変数の真偽値を取得します。</summary>
        Public ReadOnly Property Value As Boolean

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Private Sub New(value As Boolean)
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        ''' <remarks>この構造体は真偽値型の変数を表します。</remarks>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Bool
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 変数の配列を表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、複数の変数を格納することができます。
    ''' </summary>
    NotInheritable Class ArrayVariable
        Implements IVariable

        ''' <summary>変数の配列値を取得します。</summary>
        Public ReadOnly Value As IExpression()

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値の配列。</param>
        Public Sub New(value As IExpression())
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        ''' <remarks>この構造体は真偽値型の変数を表します。</remarks>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Array
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 変数のオブジェクトを表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、オブジェクト型の変数を表現します。
    ''' </summary>
    NotInheritable Class ObjectVariable
        Implements IVariable

        ''' <summary>変数のオブジェクト値を取得します。</summary>
        Public ReadOnly Property Value As Object

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">オブジェクト。</param>
        Public Sub New(value As Object)
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        ''' <remarks>この構造体はオブジェクト型の変数を表します。</remarks>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Obj
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 変数の日付を表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、日付型の変数を表現します。
    ''' </summary>
    NotInheritable Class DateTimeVariable
        Implements IVariable

        ''' <summary>変数の日付値を取得します。</summary>
        Public ReadOnly Property Value As DateTime

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Public Sub New(value As DateTime)
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        ''' <remarks>この構造体は日付値型の変数を表します。</remarks>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Date
            End Get
        End Property

    End Class

    ''' <summary>
    ''' 変数の時間範囲を表す構造体です。
    ''' この構造体は、IVariableインターフェイスを実装し、時間範囲型の変数を表現します。
    ''' </summary>
    NotInheritable Class TimeSpanVariable
        Implements IVariable

        ''' <summary>変数の時間範囲値を取得します。</summary>
        Public ReadOnly Property Value As TimeSpan

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Public Sub New(value As TimeSpan)
            Me.Value = value
        End Sub

        ''' <summary>
        ''' 変数の型を取得します。
        ''' </summary>
        ''' <returns>変数の型。</returns>
        ''' <remarks>この構造体は時間範囲値型の変数を表します。</remarks>
        Public ReadOnly Property Type As VariableType Implements IVariable.Type
            Get
                Return VariableType.Time
            End Get
        End Property

    End Class


End Namespace
