Imports EnvDTE
Imports EnvDTE80

Public Class DigPro
    Private Shared _companyname As String = "Digital Production"
    Private Shared _addinname As String = "VisualStudioTools"
    Private Shared _connectstring As String = _addinname + ".Connect"
    Private Shared _application As DTE2

    Public Shared ReadOnly Property CompanyName() As String
        Get
            Return _companyname
        End Get
    End Property

    Public Shared ReadOnly Property AddInName() As String
        Get
            Return _addinname
        End Get
    End Property

    Public Shared ReadOnly Property ConnectString() As String
        Get
            Return _connectstring
        End Get
    End Property

    Public Shared Property Application() As DTE2
        Get
            Return _application
        End Get
        Set(ByVal value As DTE2)
            _application = value
        End Set
    End Property
End Class
