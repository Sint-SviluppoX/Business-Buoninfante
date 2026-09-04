Imports System.Data
Imports NTSInformatica.CLN__STD

Imports System
Public Class FRO__CLIE
  Inherits FRM__CLIE

  Public _oClfClie As CLF__CLIE
  Public Property oClfClie() As CLF__CLIE
    Get
      If _oClfClie Is Nothing Then _oClfClie = CType(ocleClie, CLF__CLIE)
      Return _oClfClie
    End Get
    Set(ByVal value As CLF__CLIE)
      _oClfClie = value
    End Set
  End Property

End Class

