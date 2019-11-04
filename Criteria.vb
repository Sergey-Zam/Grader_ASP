Public Class Criteria
    Private criteria_id As String 'id аспекта, для его идентификации (имя может быть не уникальным/длинным)
    Private criteria_name As String 'имя/название аспекта
    Private criteria_weight As String 'вес аспекта
    Private criteria_tolerance As Double 'допустимое отклонение аспекта
    Private criteria_value As String 'значение аспекта (будет получено из документов сборок)

    Public Property id() As String
        Get
            Return criteria_id
        End Get
        Set(ByVal value As String)
            criteria_id = value
        End Set
    End Property

    Public Property name() As String
        Get
            Return criteria_name
        End Get
        Set(ByVal value As String)
            criteria_name = value
        End Set
    End Property

    Public Property weight() As String
        Get
            Return criteria_weight
        End Get
        Set(ByVal value As String)
            criteria_weight = value
        End Set
    End Property

    Public Property tolerance() As Double
        Get
            Return criteria_tolerance
        End Get
        Set(ByVal value As Double)
            criteria_tolerance = value
        End Set
    End Property

    Public Property value() As String
        Get
            Return criteria_value
        End Get
        Set(ByVal value As String)
            criteria_value = value
        End Set
    End Property

End Class
