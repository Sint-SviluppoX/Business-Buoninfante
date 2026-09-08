Imports System.Data
Imports NTSInformatica.CLN__STD

Imports System
Public Class FROHHFESI
    Inherits FRMHHFESI

    Public _oClfCons As CLFHHFESI
    Public Property oClfCons() As CLFHHFESI
        Get
            If _oClfCons Is Nothing Then _oClfCons = CType(oCleCons, CLFHHFESI)
            Return _oClfCons
        End Get
        Set(ByVal value As CLFHHFESI)
            _oClfCons = value
        End Set
    End Property

End Class

