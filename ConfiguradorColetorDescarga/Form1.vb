Imports SldWorks
Imports SwConst
Imports System.IO
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1

    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim dirPathTemplate = "C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR DESCARGA"
    Dim dirConexoes = "C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\CONFIGURADOR\TEMPLATES\CONEXOES"
    Dim errors, warnings As Integer
    Dim coletores As List(Of Coletor)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Try
            swApp = GetObject("", "SldWorks.Application")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        If swApp Is Nothing Then
            MsgBox("Erro ao conectar")
            Exit Sub
        End If

        'swApp.SendMsgToUser("Conectado")

        coletores = GetDadosColetor() 'Recebe o retorno com coletores da planilha

        For Each coletor As Coletor In coletores


            'Nome do template e o Codigo vem da planilha
            'Faz a copia do 2d e abre para troca de ref. por saveAs
            Dim nomeNovo = CopiarTemplate(coletor.Template, coletor.Codigo, ".SLDDRW")
            nomeNovo = CopiarTemplate(coletor.Tubo.Template, coletor.Tubo.Codigo, ".SLDPRT")
            'AbrirArquivo(coletor.Codigo, dirPathTemplate, ".SLDDRW")
            'AbrirArquivo(coletor.Template, dirPathTemplate, ".SLDASM")

            'coletor.AbrirArquivo(swApp, swModel, ".SLDDRW", dirPathTemplate)
            'coletor.AbrirArquivo(swApp, swModel, ".SLDASM", dirPathTemplate)


            AbrirArquivo(swApp, swModel, dirPathTemplate, ".SLDDRW", coletor.Codigo)
            AbrirArquivo(swApp, swModel, dirPathTemplate, ".SLDASM", coletor.Template)
            SalvarComo(coletor.Codigo, ".SLDASM")
            ReplacePeca(swModel, dirPathTemplate, coletor.BSolda.Template, coletor.BSolda.Codigo, ".SLDPRT")
            ReplacePeca(swModel, dirPathTemplate, coletor.Tubo.Template, coletor.Tubo.Codigo, ".SLDPRT")
            ReplacePeca(swModel, dirConexoes, coletor.Cap.Template, coletor.Cap.Codigo, ".SLDPRT")

            AbrirArquivo(swApp, swModel, dirPathTemplate, ".SLDPRT", coletor.Tubo.Codigo)
            coletor.Tubo.RedimensionarTubo(swApp, swModel)
            swModel.Save()

            Salvar(coletor.Codigo, ".SLDDRW")

            swApp.CloseAllDocuments(False)

            'Salvar(coletor.Codigo, ".SLDDRW")

            'Abre o template que está vinculado no 2d para salvar como e trocar ref.
            'Dim nomeTemp = Path.Combine(dirPathTemplate, Path.ChangeExtension(nomeDoTemplate, "SLDASM"))
            'swModel = AbrirNovo(nomeTemp)
            'swApp.ActivateDoc(nomeTemp)
            'swModel.SaveAs(Path.ChangeExtension(nomeNovo, "SLDASM"))

        Next

        'Replace(swModel, novaPeca)

    End Sub

    Private Sub Salvar(codigo As String, extensao As String)
        swApp.ActivateDoc(codigo + extensao)
        swModel = swApp.ActiveDoc
        swModel.Save()
    End Sub

    Private Sub SalvarComo(codigo As String, extensao As String)
        swApp.ActivateDoc(codigo + extensao)
        swModel = swApp.ActiveDoc
        swModel.SaveAs3(codigo + extensao, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent)
    End Sub

    Private Function CopiarTemplate(nomeDoTemplate As String, codigo As String, extensao As String) As String
        'Monta as informações para fazer a cópia.
        Dim enderecoNomeTemplate = Path.Combine(dirPathTemplate, nomeDoTemplate) + extensao
        Dim enderecoNomeNovo = Path.Combine(dirPathTemplate, codigo) + extensao
        File.Copy(enderecoNomeTemplate, enderecoNomeNovo, True)

        'return path o novo arquivo
        Return enderecoNomeNovo
    End Function


    Private Function GetDadosColetor() As List(Of Coletor)

        Dim coletores As List(Of Coletor) = New List(Of Coletor)

        Dim appXL As Excel.Application = New Excel.Application
        Dim wbXL As Excel.Workbook = appXL.Workbooks.Open("C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\col_descarga.xlsx")
        Dim shXL As Excel.Worksheet = wbXL.Sheets(1)
        Dim raXL As Excel.Range = shXL.UsedRange
        shXL = wbXL.ActiveSheet

        Dim rowCount = raXL.Rows.Count
        Dim colCount = raXL.Columns.Count

        Try

            For i = 2 To 11

                Dim coletor = New Coletor 'Cada iteração cria um novo

                Dim cellValue = CType(raXL.Cells(i, 1), Excel.Range) 'Cells retorna oject que e convertido para Range
                coletor.Template = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 2), Excel.Range)
                coletor.Codigo = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 3), Excel.Range)
                coletor.qtRamal = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 5), Excel.Range)
                coletor.BSolda.Codigo = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 8), Excel.Range)
                coletor.Tubo.Codigo = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 9), Excel.Range)
                coletor.Tubo.Template = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 10), Excel.Range)
                coletor.Tubo.DiaExterno = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 11), Excel.Range)
                coletor.Tubo.EspParede = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 12), Excel.Range)
                coletor.Tubo.DiaBSolda = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 13), Excel.Range)
                coletor.Tubo.ProfBSolda = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 14), Excel.Range)
                coletor.Tubo.DiaEncaixeRamal = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 15), Excel.Range)
                coletor.BSolda.Template = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 16), Excel.Range)
                coletor.Cap.Template = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 17), Excel.Range)
                coletor.Cap.Codigo = cellValue.Value.ToString

                'MsgBox($"código do tubo -> {coletor.Tubo.Codigo.ToString}")
                coletores.Add(coletor) 'Adiciona na lista
            Next

        Finally

            GC.Collect()
            GC.WaitForPendingFinalizers()
            Marshal.ReleaseComObject(raXL)
            Marshal.ReleaseComObject(shXL)
            wbXL.Close()
            Marshal.ReleaseComObject(wbXL)
            appXL.Quit()
            Marshal.ReleaseComObject(appXL)

        End Try
        Return coletores
    End Function

End Class
