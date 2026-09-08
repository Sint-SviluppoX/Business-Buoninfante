
Public Class DTO_HHID_TestataImp
    Public strFonte As String = ""

    Public nConto As Integer = 0
    Public nCoddest As Integer = 0

    Public strTipoOrd As String = "R"
    Public nAnnoOrd As Integer = 0
    Public strSerieOrd As String = " "
    Public lNumTmpOrd As Integer = 0

    Public dDatDoc As DateTime
    Public strRiferim As String = ""

    Public nScoPag As Integer = 0
    Public strScorporo As String = "N" ' S or N

    Public nTipoBf As Integer = 0
    Public nCodMaga As Integer = 0
    Public nListino As Integer = 0
    Public nCodPaga As Integer = 0

    Public nSpeAcc As Decimal = 0 'Trasporto

    Public dDtaImpo As DateTime = DateTime.Now
    Public dOrImpo As Decimal = CDec(DateTime.Now.TimeOfDay.Hours + DateTime.Now.TimeOfDay.Minutes / 100)

    Public oCorpo As New List(Of DTO_HHID_CorpoImp)
    Public oCliente As New DTO_HHID_ClienteImp
End Class

Public Class DTO_HHID_CorpoImp
    Public nRiga As Integer = 0
    Public strCodArt As String = ""
    Public strCodAlt As String = ""
    Public strDescr As String = ""

    Public dQuant As Decimal = 0
    Public dPrezzoIva As Decimal = 0
    Public dPrezzo As Decimal = 0

    Public nControp As Decimal = 0
    Public nCodIva As Decimal = 0

    Public dColPre As Decimal = 0
    Public dQuaPre As Decimal = 0

    Public _dPrzPromo As Decimal = 0
    Public nContropPromo As Decimal = 0
    Public nCodIvaPromo As Decimal = 0
    Public strTipoPromo As String = "ART"
    Public Property dPrzPromo As Decimal
        Get
            If _dPrzPromo > 0 Then
                ' Abbrevia il nome della città
                Return (_dPrzPromo * -1)
            Else
                Return _dPrzPromo
            End If
        End Get
        Set(value As Decimal)
            _dPrzPromo = value
        End Set
    End Property

End Class

Public Class DTO_HHID_ClienteImp
    Public strNome As String = ""
    Public strEmail As String = ""
    Public strPhone As String = ""

    Public strPrivato As String = "S" ' S or N

    Public _strIndir As String = ""
    Public _strCitta As String = ""
    Public strCap As String = ""
    Public strCodComune As String = ""
    Public strCodProv As String = ""
    Public strCodStato As String = ""
    Public strCodNazione As String = ""

    Public Property strIndir As String
        Get
            If _strIndir.Length > 50 Then
                ' Abbrevia il nome della città
                Return _strIndir.Substring(0, 40)
            Else
                Return _strIndir
            End If
        End Get
        Set(value As String)
            _strIndir = value
        End Set
    End Property
    Public Property strCitta As String
        Get
            If _strCitta.Length > 50 Then
                ' Abbrevia il nome della città
                Return _strCitta.Substring(0, 50)
            Else
                Return _strCitta
            End If
        End Get
        Set(value As String)
            _strCitta = value
        End Set
    End Property
End Class




