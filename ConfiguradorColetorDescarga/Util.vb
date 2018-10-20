Imports System.IO
Imports SldWorks
Imports SwConst

Module Util

    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2


    Public Function AbrirArquivo(_swApp As SldWorks.SldWorks, _swModel As ModelDoc2, dir As String, extensao As String, codigo As String) As ModelDoc2

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

        Dim fullPath = Path.Combine(dir, codigo) + extensao

        swModel = _swModel
        swApp = _swApp
        Dim errors, warnings As Integer

        swModel = swApp.OpenDoc6(fullPath, docType, swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", errors, warnings)
        swApp.ActivateDoc(fullPath)
        swModel = swApp.ActiveDoc
        Return swModel

    End Function

    Public Sub ReplacePeca(swModel As ModelDoc2, dir As String, codigoVelho As String, codigoNovo As String, extensao As String)
        Dim swAssembly As AssemblyDoc = swModel

        'Pega todos os componentes top level da montagem
        Dim components = swAssembly.GetComponents(True)

        'Percorre todos os componetes
        For Each comp In components
            'Cast de object components para Component2
            Dim component As Component2 = comp
            swModel = component.GetModelDoc2
            Dim fullPathName = swModel.GetPathName
            Dim nomeSemExtensao = Path.GetFileNameWithoutExtension(fullPathName)

            If nomeSemExtensao = codigoVelho Then
                component.Select(False)
                Exit For
            End If
        Next

        Dim fullPathToReplace = Path.Combine(dir, codigoNovo) + extensao
        Dim retVal = swAssembly.ReplaceComponents(fullPathToReplace, "", True, True)
        'swModel = swApp.ActivateDoc()
        swModel.Save()

    End Sub


    Sub SetPropriedade(swApp As SldWorks.SldWorks, swModel As ModelDoc2, nomeDaProp As String, valorDaProp As String)

        swModel = swApp.ActiveDoc
        Dim swExt As ModelDocExtension
        swExt = swModel.Extension
        Dim configMgr As ConfigurationManager
        Dim swCustProp As CustomPropertyManager
        configMgr = swModel.ConfigurationManager
        Dim config As Configuration
        config = configMgr.ActiveConfiguration
        Dim nomeConfig = config.Name
        swCustProp = swExt.CustomPropertyManager(nomeConfig)
        swCustProp.Add3(nomeDaProp, swCustomInfoType_e.swCustomInfoText, valorDaProp, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)

    End Sub

    Public Enum Extensao
        SLDPRT
        SLDASM
        SLDDRW
        PDF
        DWG
        DXF
    End Enum

End Module
