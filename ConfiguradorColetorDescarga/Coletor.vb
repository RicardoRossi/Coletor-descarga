Imports System.IO
Imports SldWorks
Imports SwConst

Public Class Coletor

    Public Property Codigo As String
    Public Property Template As String
    Public Property Tubo As Tubo
    Public Property BSolda As BSolda
    Public Property qtRamal As Integer
    Public Property Cap As Cap

    Public Sub New()
        Tubo = New Tubo
        BSolda = New BSolda
        Cap = New Cap
    End Sub

    Public Function AbrirArquivo(_swApp As SldWorks.SldWorks, _swModel As ModelDoc2, extensao As String, dir As String) As ModelDoc2

        Dim extensaoArquivo = extensao
        Dim docType = ""

        Select Case extensaoArquivo.ToUpper
            Case ".SLDASM"
                docType = swDocumentTypes_e.swDocASSEMBLY

            Case ".SLDPRT"
                docType = swDocumentTypes_e.swDocPART

            Case ".SLDDRW"
                docType = swDocumentTypes_e.swDocDRAWING
        End Select

        Dim fullPath = Path.Combine(dir, Codigo) + extensao

        Dim swModel = _swModel
        Dim swApp = _swApp
        Dim errors, warnings As Integer

        swModel = swApp.OpenDoc6(fullPath, docType, swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", errors, warnings)
        swApp.ActivateDoc(fullPath)
        swModel = swApp.ActiveDoc
        Return swModel

    End Function

End Class
