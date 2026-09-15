Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization


Imports System
Public Class CLFMGARTI
  Inherits CLEMGARTI

  Public _oClhArti As CLHMGARTI
  Public Property oClhArti() As CLHMGARTI
    Get
      If _oClhArti Is Nothing Then _oClhArti = CType(ocldArti, CLHMGARTI)
      Return _oClhArti
    End Get
    Set(ByVal value As CLHMGARTI)
      _oClhArti = value
    End Set
  End Property

End Class