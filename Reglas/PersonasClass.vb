Public Class PersonasClass
    Dim id_, codpostal_, idprovincia_ As Integer
    Dim nrodocumento_ As Decimal
    Dim nombre_, direccion_, tipodocumento_ As String
   
   
   
    Public Property id() As Integer
        Get
            Return id_
        End Get
        Set(ByVal value As Integer)
            id_ = value
        End Set
    End Property


    Public Property nombre() As String
        Get
            Return nombre_
        End Get
        Set(ByVal value As String)
            nombre_ = value
        End Set
    End Property



    Public Property nrodocumento() As Decimal
        Get
            Return nrodocumento_
        End Get
        Set(ByVal value As Decimal)
            nrodocumento_ = value
        End Set
    End Property


    Public Property tipodocumento() As String
        Get
            Return tipodocumento_
        End Get
        Set(ByVal value As String)
            tipodocumento_ = value
        End Set
    End Property
    Public Property idprovincia() As Integer
        Get
            Return idprovincia_
        End Get
        Set(ByVal value As Integer)
            idprovincia_ = value
        End Set
    End Property


    Public Property direccion() As String
        Get
            Return direccion_
        End Get
        Set(ByVal value As String)
            direccion_ = value
        End Set
    End Property


    Public Property codpostal() As Integer
        Get
            Return codpostal_
        End Get
        Set(ByVal value As Integer)
            codpostal_ = value
        End Set
    End Property

    Dim documentos_() As String = {"DNI", "LE", "LC", "CI", "PAS"}

    Dim provincias_() As String = _
    {"Ciudad Autónoma de Buenos Aires", _
    "Buenos Aires", _
    "Catamara", _
    "Córdoba", _
    "Corrientes", _
    "Entre Ríos", _
    "Jujuy", _
    "Mendoza", _
    "La Rioja", _
    "Salta", _
    "San Juan", _
    "San Luis", _
    "Santa Fe", _
    "Santiago del Estero", _
    "Tucumán", _
    "Chaco", _
    "Chubut", _
    "Formosa", _
    "Misiones", _
    "Neuquén", _
    "La Pampa", _
    "Río Negro", _
    "Santa Cruz", _
    "Tierra del Fuego"}

    Public ReadOnly Property documentos() As Array
        Get
            Return documentos_
        End Get
    End Property


    Public ReadOnly Property nombreDocumentos() As String
        Get

            Return documentos_(CInt(tipodocumento_))
        End Get
    End Property


    Public ReadOnly Property provincias() As Array
        Get
            Return provincias_
        End Get
    End Property



    Public ReadOnly Property nombreProvincia() As String
        Get
           
            Return provincias_(CInt(idprovincia_))
        End Get
    End Property

End Class