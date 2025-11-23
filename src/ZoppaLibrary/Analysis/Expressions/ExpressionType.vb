Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 式の型を定義する列挙型です。
    ''' 式は、リスト、変数、条件式、演算子など、さまざまな形式を持つことができます。
    ''' </summary>
    Public Enum ExpressionType

        ''' <summary>空式。</summary>
        EmptyExpression

        ''' <summary>リスト式。</summary>
        ListExpression

        ''' <summary>非展開式式。</summary>
        PlainTextExpression

        ''' <summary>展開式式（エスケープ）</summary>
        UnfoldExpression

        ''' <summary>非展開式式（アンエスケープ）</summary>
        NoEscapeUnfoldExpression

        ''' <summary>変数式。</summary>
        VariableExpression

        ''' <summary>変数宣言式リスト。</summary>
        VariablesDefineListExpression

        ''' <summary>変数宣言式。</summary>
        VariableDefineExpression

        ''' <summary>ifブロック。</summary>
        IfStatementExpression

        ''' <summary>IF式。</summary>
        IfExpression

        ''' <summary>IF Else式。</summary>
        ElseExpression

        ''' <summary>三項演算式。</summary>
        TernaryExpression

        ''' <summary>括弧式。</summary>
        ParenExpression

        ''' <summary>二項演算式。</summary>
        BinaryExpression

        ''' <summary>単項演算式（前置き）</summary>
        UnaryExpression

        ''' <summary>数値式。</summary>
        NumberExpression

        ''' <summary>文字列式。</summary>
        StringExpression

        ''' <summary>非エスケープ文字列式。</summary>
        NoEscapeStringExpression

        ''' <summary>真偽値式。</summary>
        BooleanExpression

        ''' <summary>配列フィールド式。</summary>
        ArrayFieldExpression

        ''' <summary>配列アクセス式。</summary>
        ArrayAccessExpression

        ''' <summary>識別子式。</summary>
        IdentifierExpression

        ''' <summary>For式。</summary>
        ForExpression

        ''' <summary>Selectブロック。</summary>
        SelectStatementExpression

        ''' <summary>Select式。</summary>
        SelectExpression

        ''' <summary>Select Case式。</summary>
        SelectCaseExpression

        ''' <summary>Select Default式。</summary>
        SelectDefaultExpression

        ''' <summary>関数引数式。</summary>
        FunctionArgsExpression

        ''' <summary>関数式。</summary>
        FunctionCallExpression

        ''' <summary>オブジェクト式。</summary>
        ObjectExpression

        ''' <summary>オブジェクトフィールドアクセス式。</summary>
        FieldAccessExpression

        ''' <summary>Null値式。</summary>
        NullExpression

        ''' <summary>時間式。</summary>
        DateTimeExpression

        ''' <summary>時間式。</summary>
        TimeSpanExpression

        ''' <summary>代入式。</summary>
        SetExpression

        ''' <summary>変数代入式リスト。</summary>
        SetVariableListExpression

        ''' <summary>変数代入式。</summary>
        SetVariableExpression

        ''' <summary>改行式。</summary>
        BrExpression

        ''' <summary>仮想改行式。</summary>
        VlBrExpression

        ''' <summary>Trim式。</summary>
        TrimExpression

        ''' <summary>Rem式。</summary>
        RemoveExpression

    End Enum

End Namespace
