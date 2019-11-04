Public Class Result
    Private result_name As String 'имя/название аспекта
    Private result_weight As String 'вес аспекта
    Private result_tolerance As Double 'допустимое отклонение аспекта
    Private result_standard_value As String 'значение аспекта из эталонной сборки
    Private result_checked_value As String 'значение аспекта из проверяемой эталонной сборки
    Private result_delta As Double 'текущее отклонение (расчетное поле)
    'невидимые поля
    '0-нет, 1-только в пределах диапазона, 2-да, полностью 
    Private result_is_correct As Integer 'совпадает ли правильное значение с эталонным? (расчетное поле)

    Public Property name() As String
        Get
            Return result_name
        End Get
        Set(ByVal value As String)
            result_name = value
        End Set
    End Property

    Public Property weight() As String
        Get
            Return result_weight
        End Get
        Set(ByVal value As String)
            result_weight = value
        End Set
    End Property

    Public Property tolerance() As Double
        Get
            Return result_tolerance
        End Get
        Set(ByVal value As Double)
            result_tolerance = value
        End Set
    End Property

    Public Property standard_value() As String
        Get
            Return result_standard_value
        End Get
        Set(ByVal value As String)
            result_standard_value = value
        End Set
    End Property

    Public Property checked_value() As String
        Get
            Return result_checked_value
        End Get
        Set(ByVal value As String)
            result_checked_value = value
        End Set
    End Property

    Public Property delta() As Double
        Get
            Return result_delta
        End Get
        Set(ByVal value As Double)
            result_delta = value
        End Set
    End Property

    Public Property is_correct() As Integer
        Get
            Return result_is_correct
        End Get
        Set(ByVal value As Integer)
            result_is_correct = value
        End Set
    End Property
End Class
