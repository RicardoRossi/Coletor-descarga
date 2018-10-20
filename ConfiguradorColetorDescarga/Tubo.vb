Imports System.Globalization
Imports SldWorks
Imports SwConst

Public Class Tubo
    Public Property Codigo As String
    Public Property Template As String
    Public Property diamBSoldaTubo As String


    Private _diaExterno As String
    Public Property DiaExterno() As String
        Get
            Return _DiaExterno
        End Get
        Set(ByVal value As String)
            _diaExterno = value.Replace(",", ".")
        End Set
    End Property

    Private _espParede As String
    Public Property EspParede() As String
        Get
            Return _EspParede
        End Get
        Set(ByVal value As String)
            _EspParede = value.Replace(",", ".")
        End Set
    End Property

    Private _diaBSolda As String
    Public Property DiaBSolda() As String
        Get
            Return _diaBSolda
        End Get
        Set(ByVal value As String)
            _diaBSolda = value.Replace(",", ".")
        End Set
    End Property

    Private _profBSolda As String
    Public Property ProfBSolda() As String
        Get
            Return _profBSolda
        End Get
        Set(ByVal value As String)
            _profBSolda = value.Replace(",", ".")
        End Set
    End Property

    Private _diaEncaixeRamal As String
    Public Property DiaEncaixeRamal() As String
        Get
            Return _DiaEncaixeRamal
        End Get
        Set(ByVal value As String)
            _DiaEncaixeRamal = value.Replace(",", ".")
        End Set
    End Property

    'Dim swApp As SldWorks.SldWorks
    'Dim swModel As ModelDoc2
    Dim swExt As ModelDocExtension
    Dim retVal As Boolean
    'Dim dimensao As Object = New Object

    Sub RedimensionarTubo(swApp As SldWorks.SldWorks, swModel As ModelDoc2)

        Dim d As Object
        Dim dimensao As Dimension
        swModel = swApp.ActiveDoc

        Dim diametroTubo = Double.Parse(DiaExterno, CultureInfo.InvariantCulture)
        Dim espessuraParede = Double.Parse(EspParede, CultureInfo.InvariantCulture)
        Dim diametroBSolda = Double.Parse(DiaBSolda, CultureInfo.InvariantCulture)
        Dim profundidadeBSolda = Double.Parse(ProfBSolda, CultureInfo.InvariantCulture)
        Dim diametroEncaixeRamal = Double.Parse(DiaEncaixeRamal, CultureInfo.InvariantCulture)


        d = swModel.Parameter("diam_tb@Esboço1") 'Retorna um object
        dimensao = CType(d, Dimension) 'Cast para Dimension
        dimensao.SystemValue = diametroTubo / 1000.0 'Medida em metros

        swModel.EditRebuild3()

        d = swModel.Parameter("esp_parede_tb@Esboço1") 'Retorna um object
        dimensao = CType(d, Dimension) 'Cast para Dimension
        dimensao.SystemValue = espessuraParede / 1000.0 'Medida em metros

        swModel.EditRebuild3()

        d = swModel.Parameter("diam_ramal@Esboço6") 'Retorna um object
        dimensao = CType(d, Dimension) 'Cast para Dimension
        dimensao.SystemValue = diametroEncaixeRamal / 1000.0 'Medida em metros

        swModel.EditRebuild3()

        d = swModel.Parameter("diam_bs@Esboço7") 'Retorna um object
        dimensao = CType(d, Dimension) 'Cast para Dimension
        dimensao.SystemValue = diametroBSolda / 1000.0 'Medida em metros

        swModel.EditRebuild3()

        d = swModel.Parameter("prof_bs@Corte-extrusão3") 'Retorna um object
        dimensao = CType(d, Dimension) 'Cast para Dimension
        dimensao.SystemValue = profundidadeBSolda / 1000.0 'Medida em metros

    End Sub

End Class
