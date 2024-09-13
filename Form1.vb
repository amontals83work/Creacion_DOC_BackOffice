Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports System.Drawing
Imports System.Data.SqlClient
Imports Microsoft.Win32

'acuerdo de pago
'ajustar quita y fecha
'descargar en escritorio
Public Class Form1

    Dim connectionString As String = "Data Source=192.168.50.48;Initial Catalog=DespachoMc;Persist Security Info=True;User ID=sa;Password=Binabiq2018_;MultipleActiveResultSets=True;"

    Dim nif As String = ""
    Dim nombre As String = String.Empty
    Dim expediente As String = String.Empty
    Dim refCliente As String = String.Empty
    Dim importe As Double = 0.0
    Dim idCliente As String = String.Empty
    Dim contrato As String = String.Empty
    Dim cliente As String = String.Empty
    Dim fechaHoy As String = DateTime.Now.ToString("yyyyMMdd")
    Dim portfolio As String = String.Empty
    Dim portfolioText As String = String.Empty
    Dim tipo As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Size = New Size(Me.Width, 145)

        cBox.SelectedIndex = -1
        cBox.DisplayMember = ""
        cBox.Items.Add("Carta de acuerdo AK")
        cBox.Items.Add("Carta de acuerdo plazos AK")
        cBox.Items.Add("Carta de cancelación")
        cBox.Items.Add("Helloletter")

        cBoxCancelacion.SelectedIndex = -1
        cBoxCancelacion.DisplayMember = ""
        cBoxCancelacion.Items.Add("Crisalida")
        cBoxCancelacion.Items.Add("Orange")
        cBoxCancelacion.Items.Add("Pagantis")
        cBoxCancelacion.Visible = False

        txtDNI.Visible = False
        pnlDatos.Visible = False
        pnlBotones.Visible = False

        btnDescargar.Visible = False
        btnBorrar.Visible = False

    End Sub

    Private Sub cBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cBox.SelectedIndexChanged

        txtDNI.Visible = False
        Borrar()

        If cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo AK" Then

            txtDNI.Visible = True
            cBoxCancelacion.Visible = False
            pnlDatos.Visible = False
            lblPlazos.Text = "Quita"
            lblPlazos.Visible = True
            lblPlazoAPT.Text = "Fecha de Pago"
            lblPlazoAPT.Visible = True
            txtContrato.Visible = False
            txtImporte.Visible = False
            txtImporteQuita.Visible = False
            txtImportePlazo1.Visible = True
            txtImportePlazo2.Visible = False
            txtImportePlazo3.Visible = False
            txtImportePlazo4.Visible = False
            txtImportePlazo5.Visible = False
            txtImportePlazo6.Visible = False
            txtFechaPlazo1.Visible = True
            txtFechaPlazo2.Visible = False
            txtFechaPlazo3.Visible = False
            txtFechaPlazo4.Visible = False
            txtFechaPlazo5.Visible = False
            txtFechaPlazo6.Visible = False
            pnlAPT.Location = New Drawing.Point(0, 50)
            pnlBotones.Location = New Drawing.Point(10, 115)
            pnlBotones.Visible = True
            btnDescargar.Visible = False
            btnMostrar.Visible = True
            btnBorrar.Visible = False
            Me.Size = New Size(Me.Width, 195)
            tipo = 1

        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo plazos AK" Then

            txtDNI.Visible = True
            cBoxCancelacion.Visible = False
            pnlDatos.Visible = False
            lblPlazos.Text = "Plazos"
            lblPlazos.Visible = True
            lblPlazoAPT.Text = "Fechas de Pagos"
            lblPlazoAPT.Visible = True
            txtContrato.Visible = True
            txtImporte.Visible = True
            txtImporteQuita.Visible = True
            txtImportePlazo1.Visible = True
            txtImportePlazo2.Visible = True
            txtImportePlazo3.Visible = True
            txtImportePlazo4.Visible = True
            txtImportePlazo5.Visible = True
            txtImportePlazo6.Visible = True
            txtFechaPlazo1.Visible = True
            txtFechaPlazo2.Visible = True
            txtFechaPlazo3.Visible = True
            txtFechaPlazo4.Visible = True
            txtFechaPlazo5.Visible = True
            txtFechaPlazo6.Visible = True
            pnlAPT.Location = New Drawing.Point(0, 122)
            pnlBotones.Location = New Drawing.Point(10, 115)
            pnlBotones.Visible = True
            btnDescargar.Visible = False
            btnMostrar.Visible = True
            btnBorrar.Visible = False
            Me.Size = New Size(Me.Width, 195)
            tipo = 2

        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de cancelación" Then

            txtDNI.Visible = False
            cBoxCancelacion.Visible = True
            pnlDatos.Visible = False
            lblPlazos.Visible = False
            lblPlazoAPT.Visible = False
            txtContrato.Visible = False
            txtImporte.Visible = False
            txtImporteQuita.Visible = False
            txtImportePlazo1.Visible = False
            txtImportePlazo2.Visible = False
            txtImportePlazo3.Visible = False
            txtImportePlazo4.Visible = False
            txtImportePlazo5.Visible = False
            txtImportePlazo6.Visible = False
            txtFechaPlazo1.Visible = False
            txtFechaPlazo2.Visible = False
            txtFechaPlazo3.Visible = False
            txtFechaPlazo4.Visible = False
            txtFechaPlazo5.Visible = False
            txtFechaPlazo6.Visible = False
            pnlAPT.Location = New Drawing.Point(0, 122)
            pnlBotones.Location = New Drawing.Point(10, 115)
            pnlBotones.Visible = True
            btnDescargar.Visible = False
            btnMostrar.Visible = True
            btnBorrar.Visible = False
            Me.Size = New Size(Me.Width, 195)
            tipo = 3

        End If

    End Sub

    Private Sub cBoxCancelacion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cBoxCancelacion.SelectedIndexChanged

        txtDNI.Visible = True
        pnlDatos.Visible = False
        lblPlazos.Visible = False
        pnlBotones.Location = New Drawing.Point(10, 115)
        pnlBotones.Visible = True
        btnDescargar.Visible = False
        btnMostrar.Visible = True
        btnBorrar.Visible = False
        Me.Size = New Size(Me.Width, 195)
        cBoxCancelacion.Visible = False

        If cBoxCancelacion.SelectedItem IsNot Nothing AndAlso cBoxCancelacion.SelectedItem.ToString() = "Crisalida" Then
            portfolio = "Crisalida"
            portfolioText = "AXACTOR PORTFOLIO HOLDING AB"
        ElseIf cBoxCancelacion.SelectedItem IsNot Nothing AndAlso cBoxCancelacion.SelectedItem.ToString() = "Orange" Then
            portfolio = "Orange"
            portfolioText = "QUARTZ CAPITAL FUND II, cuyo origen es Orange"
        ElseIf cBoxCancelacion.SelectedItem IsNot Nothing AndAlso cBoxCancelacion.SelectedItem.ToString() = "Pagantis" Then
            portfolio = "Pagantis"
            portfolioText = "Pagamastarde, S.L. y cuyo origen es Pagantis"
        End If

    End Sub

    Private Sub btnMostrar_Click(sender As Object, e As EventArgs) Handles btnMostrar.Click

        txtDNI.Visible = False
        pnlDatos.Visible = False
        pnlBotones.Visible = False

        If cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo AK" Then
            txtDNI.Visible = True
            pnlBotones.Location = New Drawing.Point(10, 210)
            pnlDatos.Location = New Drawing.Point(10, 114)
            pnlDatos.Visible = True
            pnlBotones.Visible = True
            Me.Size = New Size(Me.Width, 320)
        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo plazos AK" Then
            txtDNI.Visible = True
            pnlBotones.Location = New Drawing.Point(10, 405)
            pnlDatos.Location = New Drawing.Point(10, 114)
            pnlDatos.Visible = True
            pnlBotones.Visible = True
            Me.Size = New Size(Me.Width, 515)
        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de cancelación" Then
            txtDNI.Visible = True
            pnlBotones.Location = New Drawing.Point(10, 165)
            pnlDatos.Location = New Drawing.Point(10, 114)
            pnlDatos.Visible = True
            pnlBotones.Visible = True
            Me.Size = New Size(Me.Width, 275)
        End If

        btnMostrar.Visible = False
        btnDescargar.Visible = True
        btnBorrar.Visible = True

        Select Case tipo
            Case 1
                DatosAcuerdoTotal()
            Case 2
                DatosAcuerdoPlazos()
            Case 3
                DatosCancelacion()
            Case 4
                'DatosHelloLetter()
        End Select

    End Sub

    Private Sub btnDescargar_Click(sender As Object, e As EventArgs) Handles btnDescargar.Click

        Select Case tipo
            Case 1
                CartaAcuerdoTotal()
            Case 2
                CartaAcuerdoPlazos()
            Case 3
                CartaCancelacion()
            Case 4
                'HelloLetter()

        End Select
    End Sub

    Private Sub DatosAcuerdoTotal()
        Dim query As String = "SELECT D.TITULAR, ED.IdExpediente, E.Expediente, E.RefCliente, E.DeudaTotal, E.IdCliente, R.CodFactura, C.Descripcion FROM Deudores AS D 
                                JOIN ExpedientesDeudores AS ED ON ED.idDeudor = D.idDeudor 
                                JOIN Expedientes AS E ON E.IdExpediente = ED.IdExpediente 
                                JOIN Recibos AS R ON R.IdExpediente = E.idExpediente 
                                JOIN Clientes AS C ON C.idCliente = E.IdCliente
                                WHERE NIF = @nif"

        Dim listaDatos As New List(Of Dictionary(Of String, String))
        nif = txtDNI.Text.Trim

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@nif", nif)
                    connection.Open()
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim registro As New Dictionary(Of String, String)
                            registro("Nombre") = reader("TITULAR").ToString()
                            registro("Expediente") = reader("Expediente").ToString()
                            registro("RefCliente") = reader("RefCliente").ToString()
                            registro("Importe") = reader("DeudaTotal").ToString()
                            registro("IdCliente") = reader("IdCliente").ToString()
                            registro("Contrato") = reader("CodFactura").ToString()
                            registro("Cliente") = reader("Descripcion").ToString()

                            listaDatos.Add(registro)
                        End While
                        If listaDatos.Count = 0 Then
                            MessageBox.Show("No se encontraron datos para el NIF proporcionado.")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error: " & ex.Message)
        End Try

        If listaDatos.Count > 1 Then
            For Each dato In listaDatos
                nombre = dato("Nombre").Trim
                expediente &= dato("RefCliente").Trim & " - "
                contrato &= dato("Contrato").Trim & " - "
                idCliente = dato("IdCliente").Trim
                importe += Convert.ToDouble(dato("Importe"))
                cliente = dato("Cliente").Trim
            Next
            expediente = expediente.TrimEnd(" "c, "-"c)
            contrato = contrato.TrimEnd(" "c, "-"c)
        ElseIf listaDatos.Count = 1 Then
            Dim dato As Dictionary(Of String, String) = listaDatos(0)
            nombre = dato("Nombre").Trim
            expediente = dato("RefCliente").Trim
            contrato = dato("Contrato").Trim
            idCliente = dato("IdCliente").Trim
            importe = Convert.ToDouble(dato("Importe"))
            cliente = dato("Cliente").Trim
        End If

        txtNombre.Text = nombre
        txtExpediente.Text = expediente
        txtContrato.Text = contrato
        txtImporte.Text = importe.ToString.Trim
    End Sub

    Private Sub CartaAcuerdoTotal()

        Dim wordApp As New Microsoft.Office.Interop.Word.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()

        Dim logo As String = "\\192.168.50.46\e\PBO\LOGO_AKCP.jpg"
        Dim logoRange As Microsoft.Office.Interop.Word.Range = doc.Range(0, 0)
        Dim imageLogo As Microsoft.Office.Interop.Word.InlineShape = doc.InlineShapes.AddPicture(logo, False, True, logoRange)
        imageLogo.Width = 75
        imageLogo.Height = 75

        ' Encabezado de AKCP
        Dim encabezadoAKCP As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        encabezadoAKCP.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        encabezadoAKCP.Text = vbCrLf & "AKCP EUROPE SCSP" & vbCrLf & "1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo" & vbCrLf
        encabezadoAKCP.Font.Size = 9
        encabezadoAKCP.Font.Bold = True

        ' Fecha
        Dim culture As New System.Globalization.CultureInfo("es-ES")
        Dim formatoFecha As String = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", culture)
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim encabezadoFecha As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        encabezadoFecha.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
        encabezadoFecha.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdYellow
        encabezadoFecha.Text = vbCrLf & "Madrid, " & formatoFecha & vbCrLf
        encabezadoFecha.Font.Size = 9
        encabezadoFecha.Font.Bold = True
        encabezadoFecha.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight

        ' Párrafo 1
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf1 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf1.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf1.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight
        prf1.Text = vbCrLf & """AKCP EUROPE SCSP."" Certifica que " & nombre.ToUpper() & " con NIF.- " & nif & ", adeuda hoy en día el importe derivado de un producto bancario en situación contenciosa que se cita a continuación." & vbCrLf
        prf1.Font.Size = 11
        prf1.Font.Bold = False

        ' Párrafo 2 DATOS
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf2 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf2.Text = vbCrLf & "Expediente: " & expediente & vbCr & "Contrato: " & contrato & vbCr & "Importe Pendiente: " & importe & "€" & vbCrLf


        ' Párrafo 3 
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf3 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf3.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf3.Text = vbCrLf & "AKCP EUROPE SCSP manifiesta y se compromete a solicitar el ARCHIVO definitivo del procedimiento, una vez que se realice y se perciba el pago de la cantidad de " & txtImportePlazo1.Text.Trim & "€, que aplicaremos a la cancelación total del contrato, siempre y cuando se realice el abono de la indicada cantidad antes del día " & txtFechaPlazo1.Text.Trim & "." & vbCrLf


        ' Párrafo 4
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf4 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf4.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf4.Text = vbCrLf & "Con dicha entrega, una vez se produzca la misma y siempre que se realice dentro del plazo concedido, AKCP EUROPE SCSP dará por pagadas las cantidades totales reclamadas."


        ' Párrafo 5 FECHA
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf5 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf5.Text = vbCrLf & "En Madrid, " & formatoFecha & vbCrLf


        ' Párrafo 6
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf6 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf6.Text = vbCrLf & "FDO. en representación de AKCP S.à.r.l. como sociedad gestora de AKCP Europe SCSp" & vbCrLf & "D. Guilherme Carvalho" & vbCrLf
        prf6.Font.Size = 9
        prf6.Font.Bold = True

        ' Firma
        Dim firma As String = "\\192.168.50.46\e\PBO\FIRMA_GUILLERME_AKCP.jpg"
        Dim firmaRange As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        Dim imageFirma As Microsoft.Office.Interop.Word.InlineShape = doc.InlineShapes.AddPicture(firma, False, True, firmaRange)
        imageFirma.Width = 100
        imageFirma.Height = 40

        ' Párrafo 7
        Dim prf7 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf7.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        prf7.Text = vbCrLf & "(FIRMANTE)"
        prf7.Font.Size = 11
        prf7.Font.Bold = True

        ' Guardar cambios y cerrar
        Dim destinationFolderPath As String = "\\192.168.50.46\e\PBO\Acuerdo Pago\" & cliente.ToString()
        If Not Directory.Exists(destinationFolderPath) Then
            Directory.CreateDirectory(destinationFolderPath)
        End If
        Dim desktopFolderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

        doc.SaveAs(Path.Combine(destinationFolderPath, "Acuerdo_Pago_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)
        doc.SaveAs(Path.Combine(desktopFolderPath, "Acuerdo_Pago_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)


        doc.Close()
        wordApp.Quit()

        MessageBox.Show("Guardado en el ESCRITORIO'")

    End Sub

    Private Sub DatosAcuerdoPlazos()
        Dim query As String = "SELECT D.TITULAR, ED.IdExpediente, E.Expediente, E.RefCliente, E.DeudaTotal, E.IdCliente, R.CodFactura, C.Descripcion FROM Deudores AS D 
                                JOIN ExpedientesDeudores AS ED ON ED.idDeudor = D.idDeudor 
                                JOIN Expedientes AS E ON E.IdExpediente = ED.IdExpediente 
                                JOIN Recibos AS R ON R.IdExpediente = E.idExpediente 
                                JOIN Clientes AS C ON C.idCliente = E.IdCliente
                                WHERE NIF = @nif"

        Dim listaDatos As New List(Of Dictionary(Of String, String))
        nif = txtDNI.Text.Trim

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@nif", nif)
                    connection.Open()
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim registro As New Dictionary(Of String, String)
                            registro("Nombre") = reader("TITULAR").ToString().Trim
                            registro("Expediente") = reader("Expediente").ToString().Trim
                            registro("RefCliente") = reader("RefCliente").ToString().Trim
                            registro("Importe") = reader("DeudaTotal").ToString().Trim
                            registro("IdCliente") = reader("IdCliente").ToString().Trim
                            registro("Contrato") = reader("CodFactura").ToString().Trim
                            registro("Cliente") = reader("Descripcion").ToString().Trim

                            listaDatos.Add(registro)
                        End While
                        If listaDatos.Count = 0 Then
                            MessageBox.Show("No se encontraron datos para el NIF proporcionado.")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error: " & ex.Message)
        End Try

        If listaDatos.Count > 1 Then
            For Each dato In listaDatos
                nombre = dato("Nombre").Trim
                expediente &= dato("RefCliente").Trim & " - "
                contrato &= dato("Contrato").Trim & " - "
                idCliente = dato("IdCliente").Trim
                importe += Convert.ToDouble(dato("Importe"))
                cliente = dato("Cliente").Trim
            Next
            expediente = expediente.TrimEnd(" "c, "-"c)
            contrato = contrato.TrimEnd(" "c, "-"c)
        ElseIf listaDatos.Count = 1 Then
            Dim dato As Dictionary(Of String, String) = listaDatos(0)
            nombre = dato("Nombre").Trim
            expediente = dato("RefCliente").Trim
            contrato = dato("Contrato").Trim
            idCliente = dato("IdCliente").Trim
            importe = Convert.ToDouble(dato("Importe"))
            cliente = dato("Cliente").Trim
        End If

        txtNombre.Text = nombre
        txtExpediente.Text = expediente
        txtContrato.Text = contrato
        txtImporte.Text = importe.ToString.Trim
    End Sub

    Private Sub CartaAcuerdoPlazos()

        Dim wordApp As New Microsoft.Office.Interop.Word.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()

        Dim logo As String = "\\192.168.50.46\e\PBO\LOGO_AKCP.jpg"
        Dim logoRange As Microsoft.Office.Interop.Word.Range = doc.Range(0, 0)
        Dim imageLogo As Microsoft.Office.Interop.Word.InlineShape = doc.InlineShapes.AddPicture(logo, False, True, logoRange)
        imageLogo.Width = 75
        imageLogo.Height = 75

        ' Encabezado de AKCP
        Dim encabezadoAKCP As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        encabezadoAKCP.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        encabezadoAKCP.Text = vbCrLf & "AKCP EUROPE SCSP" & vbCrLf & "1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo" & vbCrLf
        encabezadoAKCP.Font.Size = 9
        encabezadoAKCP.Font.Bold = True

        ' Fecha
        Dim culture As New System.Globalization.CultureInfo("es-ES")
        Dim formatoFecha As String = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", culture)
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim encabezadoFecha As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        encabezadoFecha.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
        encabezadoFecha.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdYellow
        encabezadoFecha.Text = vbCrLf & "Madrid, " & formatoFecha & vbCrLf
        encabezadoFecha.Font.Size = 9
        encabezadoFecha.Font.Bold = True
        encabezadoFecha.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight

        ' Párrafo 1
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf1 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf1.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf1.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight
        prf1.Text = vbCrLf & """AKCP EUROPE SCSP."" Certifica que " & nombre.ToUpper() & " con NIF.- " & nif & ", adeuda hoy en día el importe derivado de un producto bancario en situación contenciosa que se cita a continuación." & vbCrLf
        prf1.Font.Size = 11
        prf1.Font.Bold = False

        ' Párrafo 2 DATOS
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf2 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf2.Text = vbCrLf & "Expediente: " & expediente & vbCr & "Contrato: " & contrato & vbCr & "Importe Pendiente Total: " & importe & "€" & vbCrLf


        ' Párrafo 3 
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf3 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf3.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf3.Text = vbCrLf & "AKCP EUROPE SCSP manifiesta y se compromete a solicitar el ARCHIVO definitivo del procedimiento, una vez que se realice y se perciba el pago de la cantidad de " & txtImporteQuita.Text.Trim & "€, que aplicaremos a la cancelación total del contrato, siempre y cuando se realice el abono de la indicada cantidad en los plazos acordados que se citan a continuación:" & vbCrLf

        'Listado de Plazos
        Dim listadoImportes As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        listadoImportes.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        listadoImportes.ListFormat.ApplyBulletDefault()
        Dim listadoCompleto As String = ""
        listadoCompleto &= txtImportePlazo1.Text.Trim & "€ antes del " & txtFechaPlazo1.Text.Trim & vbCrLf & txtImportePlazo2.Text.Trim & "€ antes del " & txtFechaPlazo2.Text.Trim & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtImportePlazo3.Text) AndAlso Not String.IsNullOrWhiteSpace(txtImportePlazo3.Text) Then
            listadoCompleto &= txtImportePlazo3.Text.Trim & "€ antes del " & txtFechaPlazo3.Text.Trim & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtImportePlazo4.Text) AndAlso Not String.IsNullOrWhiteSpace(txtImportePlazo4.Text) Then
            listadoCompleto &= txtImportePlazo4.Text.Trim & "€ antes del " & txtFechaPlazo4.Text.Trim & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtImportePlazo5.Text) AndAlso Not String.IsNullOrWhiteSpace(txtImportePlazo5.Text) Then
            listadoCompleto &= txtImportePlazo5.Text.Trim & "€ antes del " & txtFechaPlazo5.Text.Trim & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtImportePlazo6.Text) AndAlso Not String.IsNullOrWhiteSpace(txtImportePlazo6.Text) Then
            listadoCompleto &= txtImportePlazo6.Text.Trim & "€ antes del " & txtFechaPlazo6.Text.Trim & vbCrLf
        End If
        listadoImportes.Text = listadoCompleto.Trim

        ' Párrafo 4
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf4 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf4.ListFormat.RemoveNumbers()
        prf4.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf4.Text = "Con dicha entrega, una vez se produzca la misma y siempre que se realice dentro del plazo concedido, AKCP EUROPE SCSP dará por pagadas las cantidades totales reclamadas."


        ' Párrafo 5 FECHA
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf5 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf5.Text = vbCrLf & "En Madrid, " & formatoFecha & vbCrLf


        ' Párrafo 6
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf6 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf6.Text = vbCrLf & "FDO. en representación de AKCP S.à.r.l. como sociedad gestora de AKCP Europe SCSp" & vbCrLf & "D. Guilherme Carvalho" & vbCrLf
        prf6.Font.Size = 9
        prf6.Font.Bold = True

        ' Firma
        Dim firma As String = "\\192.168.50.46\e\PBO\FIRMA_GUILLERME_AKCP.jpg"
        Dim firmaRange As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        Dim imageFirma As Microsoft.Office.Interop.Word.InlineShape = doc.InlineShapes.AddPicture(firma, False, True, firmaRange)
        imageFirma.Width = 100
        imageFirma.Height = 40

        ' Párrafo 7
        Dim prf7 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf7.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        prf7.Text = vbCrLf & "(FIRMANTE)"
        prf7.Font.Size = 11
        prf7.Font.Bold = True

        ' Guardar cambios y cerrar
        Dim destinationFolderPath As String = "\\192.168.50.46\e\PBO\Acuerdo Pago Plazos\" & cliente.ToString()
        If Not Directory.Exists(destinationFolderPath) Then
            Directory.CreateDirectory(destinationFolderPath)
        End If
        Dim desktopFolderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

        doc.SaveAs(Path.Combine(destinationFolderPath, "Acuerdo_Plazos_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)
        doc.SaveAs(Path.Combine(desktopFolderPath, "Acuerdo_Plazos_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)

        doc.Close()
        wordApp.Quit()

        MessageBox.Show("Guardado en el ESCRITORIO'")

    End Sub

    Private Sub DatosCancelacion()
        Dim query As String = "SELECT D.TITULAR, ED.IdExpediente, E.Expediente, E.RefCliente, E.DeudaTotal, E.IdCliente FROM Deudores AS D 
                                JOIN ExpedientesDeudores AS ED ON ED.idDeudor = D.idDeudor 
                                JOIN Expedientes AS E ON E.IdExpediente = ED.IdExpediente 
                                WHERE NIF = @nif"

        Dim listaDatos As New List(Of Dictionary(Of String, String))
        nif = txtDNI.Text.Trim

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@nif", nif)
                    connection.Open()
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim registro As New Dictionary(Of String, String)
                            registro("Nombre") = reader("TITULAR").ToString().Trim
                            registro("Expediente") = reader("Expediente").ToString().Trim
                            registro("RefCliente") = reader("RefCliente").ToString().Trim
                            registro("IdCliente") = reader("IdCliente").ToString().Trim

                            listaDatos.Add(registro)
                        End While
                        If listaDatos.Count = 0 Then
                            MessageBox.Show("No se encontraron datos para el NIF proporcionado.")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error: " & ex.Message)
        End Try

        If listaDatos.Count > 1 Then
            For Each dato In listaDatos
                nombre = dato("Nombre").Trim
                expediente &= dato("RefCliente").Trim & " - "
                idCliente = dato("IdCliente").Trim
            Next
            expediente = expediente.TrimEnd(" "c, "-"c)
            contrato = contrato.TrimEnd(" "c, "-"c)
        ElseIf listaDatos.Count = 1 Then
            Dim dato As Dictionary(Of String, String) = listaDatos(0)
            nombre = dato("Nombre").Trim
            expediente = dato("RefCliente").Trim
            idCliente = dato("IdCliente").Trim
        End If

        txtNombre.Text = nombre
        txtExpediente.Text = expediente
    End Sub

    Private Sub CartaCancelacion()

        Dim wordApp As New Microsoft.Office.Interop.Word.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()

        Dim logo As String = "\\192.168.50.46\e\PBO\LOGO_AKCP.jpg"
        Dim logoRange As Microsoft.Office.Interop.Word.Range = doc.Range(0, 0)
        Dim imageLogo As Microsoft.Office.Interop.Word.InlineShape = doc.InlineShapes.AddPicture(logo, False, True, logoRange)
        imageLogo.Width = 75
        imageLogo.Height = 75

        ' Encabezado de AKCP
        Dim encabezadoAKCP As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        encabezadoAKCP.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        encabezadoAKCP.Text = vbCrLf & "AKCP EUROPE SCSP" & vbCrLf & "1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo" & vbCrLf
        encabezadoAKCP.Font.Size = 9
        encabezadoAKCP.Font.Bold = True

        ' Fecha
        Dim culture As New System.Globalization.CultureInfo("es-ES")
        Dim formatoFecha As String = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", culture)
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim encabezadoFecha As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        encabezadoFecha.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
        encabezadoFecha.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdYellow
        encabezadoFecha.Text = vbCrLf & "Madrid, " & formatoFecha & vbCrLf
        encabezadoFecha.Font.Size = 9
        encabezadoFecha.Font.Bold = True
        encabezadoFecha.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight

        ' Referencia y Nombre
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf1 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf1.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf1.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight
        prf1.Text = vbCrLf & "Referencia del crédito: " & expediente & vbCrLf & vbCrLf & vbCrLf & "Estimado/a D/Dña.: " & nombre.ToUpper() & vbCrLf
        prf1.Font.Size = 11
        prf1.Font.Bold = False

        ' Párrafo 1
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf2 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf2.Text = vbCrLf & "Nos dirigimos a usted para recordarle que " & portfolio & ", cedió a AKCP EUROPE SCSP el crédito que tenía pendiente con usted, bajo la referencia: " & expediente & "." & vbCrLf


        ' Párrafo 2 
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf3 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf3.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf3.Text = vbCrLf & "En virtud de dicha cesión, AKCP EUROPE SCSP certifica que " & nombre & ", identificado con el NIF " & nif & " y en calidad de titular, mantenía una deuda pendiente derivada del contrato anteriormente indicado." & vbCrLf


        ' Párrafo 3
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf4 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf4.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf4.Text = vbCrLf & "Teniendo en cuenta lo anterior, y en respuesta a su solicitud por escrito sobre el estado actual de la deuda mencionada, nos complace informarle que:"

        ' Párrafo 4
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf5 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf5.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf5.Text = vbCrLf & "SU DEUDA HA SIDO TOTALMENTE CANCELADA. Por consiguiente, en relación con el contrato mencionado anteriormente, no se adeuda cantidad alguna."
        prf5.Font.Bold = True

        ' Párrafo 5 FECHA
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf6 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf6.Text = vbCrLf & "En Madrid, " & formatoFecha & vbCrLf
        prf6.Font.Bold = False

        ' Párrafo 6
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf7 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf7.Text = vbCrLf & "FDO. en representación de AKCP S.à.r.l. como sociedad gestora de AKCP Europe SCSp" & vbCrLf & "D. Guilherme Carvalho" & vbCrLf
        prf7.Font.Size = 9
        prf7.Font.Bold = True

        ' Firma
        Dim firma As String = "\\192.168.50.46\e\PBO\FIRMA_GUILLERME_AKCP.jpg"
        Dim firmaRange As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        Dim imageFirma As Microsoft.Office.Interop.Word.InlineShape = doc.InlineShapes.AddPicture(firma, False, True, firmaRange)
        imageFirma.Width = 100
        imageFirma.Height = 40

        ' Párrafo 7
        Dim prf8 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf8.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        prf8.Text = vbCrLf & "(FIRMANTE)"
        prf8.Font.Size = 11
        prf8.Font.Bold = True

        ' Guardar cambios y cerrar
        Dim destinationFolderPath As String = "\\192.168.50.46\e\PBO\Cancelacion\" & portfolio.ToString()
        If Not Directory.Exists(destinationFolderPath) Then
            Directory.CreateDirectory(destinationFolderPath)
        End If
        Dim desktopFolderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

        doc.SaveAs(Path.Combine(destinationFolderPath, "Cancelacion_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)
        doc.SaveAs(Path.Combine(desktopFolderPath, "Cancelacion_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)

        doc.Close()
        wordApp.Quit()

        MessageBox.Show("Guardado en el ESCRITORIO'")

    End Sub

    Private Sub btnBorrar_Click(sender As Object, e As EventArgs) Handles btnBorrar.Click

        Borrar()

    End Sub

    Function Borrar()

        nif = ""
        nombre = String.Empty
        expediente = String.Empty
        refCliente = String.Empty
        importe = 0.0
        idCliente = String.Empty
        contrato = String.Empty
        cliente = String.Empty
        portfolio = String.Empty
        tipo = 0

        txtDNI.Text = "DNI"
        txtNombre.Text = "Nombre"
        txtExpediente.Text = "Expediente"
        txtContrato.Text = "Contrato"
        txtImporte.Text = "Deuda total"
        txtImporteQuita.Text = "Deuda quita"

        txtImportePlazo1.Text = "0.00"
        txtImportePlazo2.Text = ""
        txtImportePlazo3.Text = ""
        txtImportePlazo4.Text = ""
        txtImportePlazo5.Text = ""
        txtImportePlazo6.Text = ""

        txtFechaPlazo1.Text = "Ejemplo: 1 de enero del 2000"

    End Function


End Class