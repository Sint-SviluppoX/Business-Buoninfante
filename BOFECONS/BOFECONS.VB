Imports System.Data
Imports NTSInformatica.CLN__STD

Imports System
Public Class FROFECONS
  Inherits FRMFECONS

  Public _oClfCons As CLFFECONS
  Public Property oClfCons() As CLFFECONS
    Get
      If _oClfCons Is Nothing Then _oClfCons = CType(ocleCons, CLFFECONS)
      Return _oClfCons
    End Get
    Set(ByVal value As CLFFECONS)
      _oClfCons = value
    End Set
  End Property

End Class

