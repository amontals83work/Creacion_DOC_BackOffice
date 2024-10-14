Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports System.Drawing
Imports System.Data.SqlClient
Imports Microsoft.Win32
Imports System.Threading

Public Class Form1

    Dim connectionString As String = "Data Source=192.168.50.48;Initial Catalog=DespachoMc;Persist Security Info=True;User ID=sa;Password=Binabiq2018_;MultipleActiveResultSets=True;"

    Dim nif As String = String.Empty
    Dim exp As String = String.Empty
    Dim nombre As String = String.Empty
    Dim expediente As String = String.Empty
    Dim aExpediente As New List(Of String)
    Dim selecExp As New List(Of String)
    Dim expSQL As New List(Of String)
    Dim selecSQL As String = String.Empty
    Dim refCliente As String = String.Empty
    Dim importe As Double = 0.0
    Dim idCliente As String = String.Empty
    Dim contrato As String = String.Empty
    Dim cliente As String = String.Empty
    Dim descripcion As String = String.Empty
    Dim fechaHoy As String = DateTime.Now.ToString("yyyyMMdd")
    Dim portfolio As String = String.Empty
    Dim portfolioText As String = String.Empty
    Dim tipo As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cBox.Items.Clear()
        cBox.SelectedIndex = -1
        cBox.DisplayMember = ""
        cBox.Items.Add("Carta de acuerdo")
        cBox.Items.Add("Carta de acuerdo a plazos")
        cBox.Items.Add("Carta de cancelación")
        cBox.Items.Add("Helloletter")

        cBoxCancelacion.Items.Clear()
        cBoxCancelacion.SelectedIndex = -1
        cBoxCancelacion.DisplayMember = ""
        cBoxCancelacion.Items.Add("Crisalidas")
        cBoxCancelacion.Items.Add("Orange")
        cBoxCancelacion.Items.Add("Pagantis")
        cBoxCancelacion.Visible = False

        cBoxHelloletter.Items.Clear()
        cBoxHelloletter.SelectedIndex = -1
        cBoxHelloletter.DisplayMember = ""
        cBoxHelloletter.Visible = False

        txtDNI.Visible = False
        txtExpHelloLetter.Visible = False
        cBoxCancelacion.Visible = False
        cBoxHelloletter.Visible = False
        pnlDatos.Visible = False
        pnlBotones.Visible = False

        chLbExpedientes.Items.Clear()

        btnDescargar.Visible = False
        btnBorrar.Visible = False

        Me.Size = New Size(258, 160)

    End Sub

    Private Sub cBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cBox.SelectedIndexChanged

        txtDNI.Visible = False
        cBoxHelloletter.Visible = False
        cBoxHelloletter.Items.Clear()
        Borrar()

        If cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo" Then

            txtDNI.Visible = True
            txtDNI.ReadOnly = False
            txtExpHelloLetter.Visible = False
            txtExpHelloLetter.ReadOnly = False
            cBoxCancelacion.Visible = False
            pnlDatos.Visible = False
            lblImportes.Text = "Importe"
            lblImportes.Visible = True
            lblPlazoAPT.Text = "Fecha de Pago"
            lblPlazoAPT.Visible = True
            txtContrato.Visible = True
            txtContrato.ReadOnly = False
            txtImporte.Visible = True
            txtImporte.ReadOnly = False
            txtImporteQuita.Visible = False
            txtImporteQuita.ReadOnly = False
            txtImportePlazo1.Visible = True
            txtImportePlazo1.ReadOnly = False
            txtImportePlazo2.Visible = False
            txtImportePlazo2.ReadOnly = False
            txtImportePlazo3.Visible = False
            txtImportePlazo3.ReadOnly = False
            txtImportePlazo4.Visible = False
            txtImportePlazo4.ReadOnly = False
            txtImportePlazo5.Visible = False
            txtFechaPlazo5.ReadOnly = False
            txtImportePlazo6.Visible = False
            txtImportePlazo6.ReadOnly = False
            txtFechaPlazo1.Visible = True
            txtFechaPlazo1.ReadOnly = False
            txtFechaPlazo2.Visible = False
            txtFechaPlazo2.ReadOnly = False
            txtFechaPlazo3.Visible = False
            txtFechaPlazo3.ReadOnly = False
            txtFechaPlazo4.Visible = False
            txtFechaPlazo4.ReadOnly = False
            txtFechaPlazo5.Visible = False
            txtFechaPlazo5.ReadOnly = False
            txtFechaPlazo6.Visible = False
            txtFechaPlazo6.ReadOnly = False
            pnlAPT.Location = New Drawing.Point(0, 100)
            pnlBotones.Location = New Drawing.Point(10, 115)
            pnlBotones.Visible = True
            btnDescargar.Visible = False
            btnMostrarExps.Visible = True
            btnMostrar.Visible = False
            btnMostrarExps.Visible = True
            btnBorrar.Visible = False
            Me.Size = New Size(258, 195)
            tipo = 1

        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo a plazos" Then

            txtDNI.Visible = True
            txtExpHelloLetter.Visible = False
            cBoxCancelacion.Visible = False
            pnlDatos.Visible = False
            lblImportes.Text = "Importes"
            lblImportes.Visible = True
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
            pnlAPT.Location = New Drawing.Point(0, 123)
            pnlBotones.Location = New Drawing.Point(10, 115)
            pnlBotones.Visible = True
            btnDescargar.Visible = False
            btnMostrarExps.Visible = True
            btnMostrar.Visible = True
            btnBorrar.Visible = False
            Me.Size = New Size(258, 195)
            tipo = 2

        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de cancelación" Then

            txtDNI.Visible = False
            txtExpHelloLetter.Visible = False
            cBoxCancelacion.Visible = True
            pnlDatos.Visible = False
            lblImportes.Visible = False
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
            btnMostrarExps.Visible = True
            btnMostrar.Visible = True
            btnBorrar.Visible = False
            Me.Size = New Size(258, 195)
            tipo = 3
            cBoxCancelacion.Font = New Drawing.Font(cBoxCancelacion.Font.FontFamily, 8.25F)

        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Helloletter" Then

            txtDNI.Visible = False
            txtExpHelloLetter.Visible = True
            cBoxCancelacion.Visible = False
            pnlDatos.Visible = False
            lblImportes.Visible = False
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
            btnMostrarExps.Visible = False
            btnMostrar.Visible = True
            btnBorrar.Visible = False
            Me.Size = New Size(258, 195)
            tipo = 4

        End If

    End Sub

    Private Sub cBoxCancelacion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cBoxCancelacion.SelectedIndexChanged

        txtDNI.Visible = True
        pnlDatos.Visible = False
        lblImportes.Visible = False
        pnlBotones.Location = New Drawing.Point(10, 115)
        pnlBotones.Visible = True
        btnDescargar.Visible = False
        btnMostrar.Visible = True
        btnBorrar.Visible = False
        Me.Size = New Size(Me.Width, 195)
        cBoxCancelacion.Visible = False

        If cBoxCancelacion.SelectedItem IsNot Nothing AndAlso cBoxCancelacion.SelectedItem.ToString() = "Crisalidas" Then
            portfolio = "Crisalidas"
            'portfolioText = "AXACTOR PORTFOLIO HOLDING AB"
        ElseIf cBoxCancelacion.SelectedItem IsNot Nothing AndAlso cBoxCancelacion.SelectedItem.ToString() = "Orange" Then
            portfolio = "Orange"
            'portfolioText = "QUARTZ CAPITAL FUND II, cuyo origen es Orange"
        ElseIf cBoxCancelacion.SelectedItem IsNot Nothing AndAlso cBoxCancelacion.SelectedItem.ToString() = "Pagantis" Then
            portfolio = "Pagantis"
            'portfolioText = "Pagamastarde, S.L. y cuyo origen es Pagantis"
        End If

    End Sub

    Private Sub cBoxHelloletter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cBoxHelloletter.SelectedIndexChanged

        pnlDatos.Visible = False
        lblImportes.Visible = False
        pnlBotones.Location = New Drawing.Point(10, 115)
        pnlBotones.Visible = True
        btnDescargar.Visible = True
        btnMostrar.Visible = False
        btnBorrar.Visible = True
        Me.Size = New Size(Me.Width, 225)

    End Sub

    Private Sub btnMostrarExps_Click(sender As Object, e As EventArgs) Handles btnMostrarExps.Click

        txtDNI.Visible = True
        txtDNI.ReadOnly = True
        btnMostrar.Visible = True
        btnMostrarExps.Visible = False
        pnlBotones.Visible = True
        pnlDatos.Visible = False
        Me.Size = New Size(355, 195)

        DatosExpedientes()

    End Sub

    Private Sub btnMostrar_Click(sender As Object, e As EventArgs) Handles btnMostrar.Click

        txtDNI.Visible = False
        pnlAPT.Visible = False
        pnlDatos.Visible = False
        pnlBotones.Visible = False
        btnMostrar.Visible = False
        btnBorrar.Visible = True

        If cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo" Then
            txtDNI.Visible = True
            txtDNI.ReadOnly = True
            txtNombre.ReadOnly = True
            txtContrato.ReadOnly = True
            txtExpediente.ReadOnly = True
            txtImporte.ReadOnly = True
            btnDescargar.Visible = True
            pnlBotones.Location = New Drawing.Point(10, 282)
            pnlDatos.Location = New Drawing.Point(10, 114)
            pnlDatos.Visible = True
            pnlAPT.Visible = True
            pnlBotones.Visible = True
            Me.Size = New Size(355, 392)
        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de acuerdo a plazos" Then
            txtDNI.Visible = True
            txtDNI.ReadOnly = True
            txtNombre.ReadOnly = True
            txtContrato.ReadOnly = True
            txtExpediente.ReadOnly = True
            txtImporte.ReadOnly = True
            btnDescargar.Visible = True
            pnlBotones.Location = New Drawing.Point(10, 425)
            pnlDatos.Location = New Drawing.Point(10, 114)
            pnlDatos.Visible = True
            pnlAPT.Visible = True
            pnlBotones.Visible = True
            Me.Size = New Size(355, 535)
        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Carta de cancelación" Then
            txtDNI.Visible = True
            txtDNI.ReadOnly = True
            txtNombre.ReadOnly = True
            txtContrato.ReadOnly = True
            txtExpediente.ReadOnly = True
            txtImporte.ReadOnly = True
            btnDescargar.Visible = True
            pnlBotones.Location = New Drawing.Point(10, 165)
            pnlDatos.Location = New Drawing.Point(10, 114)
            pnlDatos.Visible = True
            pnlBotones.Visible = True
            Me.Size = New Size(355, 275)
        ElseIf cBox.SelectedItem IsNot Nothing AndAlso cBox.SelectedItem.ToString() = "Helloletter" Then
            txtExpHelloLetter.Visible = True
            txtExpHelloLetter.ReadOnly = True
            txtNombre.ReadOnly = True
            txtContrato.ReadOnly = True
            txtExpediente.ReadOnly = True
            txtImporte.ReadOnly = True
            cBoxHelloletter.Visible = True
            btnDescargar.Visible = False
            pnlBotones.Location = New Drawing.Point(10, 85)
            pnlBotones.Visible = True
            Me.Size = New Size(258, 195)

        End If
        selecExp.Clear()
        expSQL.Clear()

        Select Case tipo
            Case 1, 2, 3
                For Each item As String In chLbExpedientes.CheckedItems
                    selecExp.Add(item)
                    expSQL.Add("'" & item & "'")
                Next
                selecSQL = String.Join(",", expSQL)
                DatosAcuerdoTotal()
            Case 4
                DescripcionCarteraHelloLetter()
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
                Helloletter()
        End Select
        Borrar()
        'System.Windows.Forms.Application.Restart()

    End Sub

    Private Sub DatosExpedientes()
        Dim query As String = "SELECT E.RefCliente FROM Deudores AS D 
                                JOIN ExpedientesDeudores AS ED ON ED.idDeudor = D.idDeudor 
                                JOIN Expedientes AS E ON E.IdExpediente = ED.IdExpediente 
                                WHERE NIF = @nif"

        Dim listaExp As New List(Of String)
        nif = txtDNI.Text.Trim
        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@nif", nif)
                    connection.Open()
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim refCliente As String = reader("RefCliente").ToString()
                            listaExp.Add(refCliente)
                        End While
                        If listaExp.Count = 0 Then
                            MessageBox.Show("No se encontraron expedientes para el NIF proporcionado.")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error: " & ex.Message)
        End Try

        If listaExp.Count > 1 Then
            For Each refCliente As String In listaExp
                aExpediente.Add(refCliente.Trim)
            Next
        ElseIf listaExp.Count = 1 Then
            aExpediente.Add(listaExp(0).Trim)
        End If

        chLbExpedientes.Items.Clear()

        For Each item As String In aExpediente
            chLbExpedientes.Items.Add(item)
        Next

    End Sub

    Private Sub DatosAcuerdoTotal()
        Dim query As String = "SELECT D.TITULAR, ED.IdExpediente, E.DeudaTotal, E.IdCliente, R.CodFactura, C.Descripcion FROM Deudores AS D " &
                              "JOIN ExpedientesDeudores AS ED ON ED.idDeudor = D.idDeudor " &
                              "JOIN Expedientes AS E ON E.IdExpediente = ED.IdExpediente " &
                              "JOIN Recibos AS R ON R.IdExpediente = E.idExpediente " &
                              "JOIN Clientes AS C ON C.idCliente = E.IdCliente " &
                              "WHERE NIF = @nif AND E.RefCliente IN (" & selecSQL & ")"

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
                contrato &= dato("Contrato").Trim & " - "
                idCliente = dato("IdCliente").Trim
                importe += Convert.ToDouble(dato("Importe"))
                cliente = dato("Cliente").Trim
            Next
            contrato = contrato.TrimEnd(" "c, "-"c)
        ElseIf listaDatos.Count = 1 Then
            Dim dato As Dictionary(Of String, String) = listaDatos(0)
            nombre = dato("Nombre").Trim
            contrato = dato("Contrato").Trim
            idCliente = dato("IdCliente").Trim
            importe = Convert.ToDouble(dato("Importe"))
            cliente = dato("Cliente").Trim
        End If

        txtNombre.Text = nombre
        txtExpediente.Text = String.Join(" - ", selecExp)
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
        prf1.Text = vbCrLf & """AKCP EUROPE SCSP."" Certifica que " & nombre.ToUpper() & " con NIF.- " & nif.ToUpper & ", adeuda hoy en día el importe derivado de un producto bancario en situación contenciosa que se cita a continuación." & vbCrLf
        prf1.Font.Size = 11
        prf1.Font.Bold = False

        ' Párrafo 2 DATOS
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf2 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf2.Text = vbCrLf & "Expediente: " & txtExpediente.Text & vbCr & "Contrato: " & contrato & vbCr & "Importe Pendiente: " & importe & "€" & vbCrLf


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

        doc.SaveAs(Path.Combine(destinationFolderPath, "Acuerdo_Pago_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim.ToUpper & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)
        doc.SaveAs(Path.Combine(desktopFolderPath, "Acuerdo_Pago_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim.ToUpper & ".pdf"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF)


        doc.Close()
        wordApp.Quit()

        MessageBox.Show("Guardado en el ESCRITORIO")

    End Sub

    'Private Sub DatosAcuerdoPlazos()
    '    Dim query As String = "SELECT D.TITULAR, ED.IdExpediente, E.DeudaTotal, E.IdCliente, R.CodFactura, C.Descripcion FROM Deudores AS D " &
    '                          "JOIN ExpedientesDeudores AS ED ON ED.idDeudor = D.idDeudor " &
    '                          "JOIN Expedientes AS E ON E.IdExpediente = ED.IdExpediente " &
    '                          "JOIN Recibos AS R ON R.IdExpediente = E.idExpediente " &
    '                          "JOIN Clientes AS C ON C.idCliente = E.IdCliente " &
    '                          "WHERE NIF = @nif AND E.RefCliente IN (" & selecSQL & ")"

    '    Dim listaDatos As New List(Of Dictionary(Of String, String))
    '    nif = txtDNI.Text.Trim

    '    Try
    '        Using connection As New SqlConnection(connectionString)
    '            Using command As New SqlCommand(query, connection)
    '                command.Parameters.AddWithValue("@nif", nif)
    '                connection.Open()
    '                Using reader As SqlDataReader = command.ExecuteReader()
    '                    While reader.Read()
    '                        Dim registro As New Dictionary(Of String, String)
    '                        registro("Nombre") = reader("TITULAR").ToString()
    '                        registro("Importe") = reader("DeudaTotal").ToString()
    '                        registro("IdCliente") = reader("IdCliente").ToString()
    '                        registro("Contrato") = reader("CodFactura").ToString()
    '                        registro("Cliente") = reader("Descripcion").ToString()

    '                        listaDatos.Add(registro)
    '                    End While
    '                    If listaDatos.Count = 0 Then
    '                        MessageBox.Show("No se encontraron datos para el NIF proporcionado.")
    '                    End If
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        MessageBox.Show("Ocurrió un error: " & ex.Message)
    '    End Try

    '    If listaDatos.Count > 1 Then
    '        For Each dato In listaDatos
    '            nombre = dato("Nombre").Trim
    '            contrato &= dato("Contrato").Trim & " - "
    '            idCliente = dato("IdCliente").Trim
    '            importe += Convert.ToDouble(dato("Importe"))
    '            cliente = dato("Cliente").Trim
    '        Next
    '        contrato = contrato.TrimEnd(" "c, "-"c)
    '    ElseIf listaDatos.Count = 1 Then
    '        Dim dato As Dictionary(Of String, String) = listaDatos(0)
    '        nombre = dato("Nombre").Trim
    '        contrato = dato("Contrato").Trim
    '        idCliente = dato("IdCliente").Trim
    '        importe = Convert.ToDouble(dato("Importe"))
    '        cliente = dato("Cliente").Trim
    '    End If

    '    txtNombre.Text = nombre
    '    txtExpediente.Text = String.Join(" - ", selecExp)
    '    txtContrato.Text = contrato
    '    txtImporte.Text = importe.ToString.Trim

    'End Sub

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
        prf1.Text = vbCrLf & """AKCP EUROPE SCSP."" Certifica que " & nombre.ToUpper() & " con NIF.- " & nif.ToUpper & ", adeuda hoy en día el importe derivado de un producto bancario en situación contenciosa que se cita a continuación." & vbCrLf
        prf1.Font.Size = 11
        prf1.Font.Bold = False

        ' Párrafo 2 DATOS
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf2 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf2.Text = vbCrLf & "Expediente: " & txtExpediente.Text & vbCr & "Contrato: " & contrato & vbCr & "Importe Pendiente Total: " & importe & "€" & vbCrLf


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
        If Not String.IsNullOrWhiteSpace(txtImportePlazo3.Text) AndAlso Not String.IsNullOrWhiteSpace(txtFechaPlazo3.Text) Then
            listadoCompleto &= txtImportePlazo3.Text.Trim & "€ antes del " & txtFechaPlazo3.Text.Trim & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtImportePlazo4.Text) AndAlso Not String.IsNullOrWhiteSpace(txtFechaPlazo4.Text) Then
            listadoCompleto &= txtImportePlazo4.Text.Trim & "€ antes del " & txtFechaPlazo4.Text.Trim & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtImportePlazo5.Text) AndAlso Not String.IsNullOrWhiteSpace(txtFechaPlazo5.Text) Then
            listadoCompleto &= txtImportePlazo5.Text.Trim & "€ antes del " & txtFechaPlazo5.Text.Trim & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtImportePlazo6.Text) AndAlso Not String.IsNullOrWhiteSpace(txtFechaPlazo6.Text) Then
            listadoCompleto &= txtImportePlazo6.Text.Trim & "€ antes del " & txtFechaPlazo6.Text.Trim & vbCrLf
        End If
        listadoImportes.Text = listadoCompleto

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

        doc.SaveAs(Path.Combine(destinationFolderPath, "Acuerdo_Plazos_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim.ToUpper & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)
        doc.SaveAs(Path.Combine(desktopFolderPath, "Acuerdo_Plazos_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim.ToUpper & ".pdf"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF)

        doc.Close()
        wordApp.Quit()

        MessageBox.Show("Guardado en el ESCRITORIO")

    End Sub

    'Private Sub DatosCancelacion()
    '    Dim query As String = "SELECT D.TITULAR, ED.IdExpediente, E.Expediente, E.RefCliente, E.DeudaTotal, E.IdCliente, R.CodFactura FROM Deudores AS D 
    '                            JOIN ExpedientesDeudores AS ED ON ED.idDeudor = D.idDeudor 
    '                            JOIN Expedientes AS E ON E.IdExpediente = ED.IdExpediente 
    '                            JOIN Recibos AS R ON R.IdExpediente = ED.idExpediente
    '                            WHERE NIF = @nif"

    '    Dim listaDatos As New List(Of Dictionary(Of String, String))
    '    nif = txtDNI.Text.Trim

    '    Try
    '        Using connection As New SqlConnection(connectionString)
    '            Using command As New SqlCommand(query, connection)
    '                command.Parameters.AddWithValue("@nif", nif)
    '                connection.Open()
    '                Using reader As SqlDataReader = command.ExecuteReader()
    '                    While reader.Read()
    '                        Dim registro As New Dictionary(Of String, String)
    '                        registro("Nombre") = reader("TITULAR").ToString().Trim
    '                        registro("Expediente") = reader("Expediente").ToString().Trim
    '                        registro("RefCliente") = reader("RefCliente").ToString().Trim
    '                        registro("IdCliente") = reader("IdCliente").ToString().Trim
    '                        registro("Contrato") = reader("CodFactura").ToString().Trim

    '                        listaDatos.Add(registro)
    '                    End While
    '                    If listaDatos.Count = 0 Then
    '                        MessageBox.Show("No se encontraron datos para el NIF proporcionado.")
    '                    End If
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        MessageBox.Show("Ocurrió un error: " & ex.Message)
    '    End Try

    '    If listaDatos.Count > 1 Then
    '        For Each dato In listaDatos
    '            nombre = dato("Nombre").Trim
    '            expediente &= dato("RefCliente").Trim & " - "
    '            aExpediente.Add(dato("RefCliente").Trim)
    '            idCliente = dato("IdCliente").Trim
    '            contrato &= dato("Contrato").Trim & " - "
    '        Next
    '        expediente = expediente.TrimEnd(" "c, "-"c)
    '        contrato = contrato.TrimEnd(" "c, "-"c)
    '    ElseIf listaDatos.Count = 1 Then
    '        Dim dato As Dictionary(Of String, String) = listaDatos(0)
    '        nombre = dato("Nombre").Trim
    '        expediente = dato("RefCliente").Trim
    '        aExpediente.Add(dato("RefCliente").Trim)
    '        idCliente = dato("IdCliente").Trim
    '        contrato = dato("Contrato").Trim
    '    End If

    '    txtNombre.Text = nombre
    '    txtExpediente.Text = expediente

    '    chLbExpedientes.Items.Clear()

    '    For Each item As String In aExpediente
    '        chLbExpedientes.Items.Add(item)
    '    Next

    'End Sub

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
        If portfolio = "Crisalidas" Then
            prf1.Text = vbCrLf & "Expediente: " & txtExpediente.Text & vbCrLf & vbCrLf & "Estimado/a D/Dña.: " & nombre.ToUpper() & vbCrLf
        Else
            prf1.Text = vbCrLf & "Referencia del crédito: " & txtExpediente.Text & vbCrLf & vbCrLf & "Estimado/a D/Dña.: " & nombre.ToUpper() & vbCrLf
        End If
        prf1.Font.Size = 10
        prf1.Font.Bold = False

        ' Párrafo 1
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf2 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf2.Text = vbCrLf & "Nos dirigimos a usted para recordarle que el crédito bajo referencia " & txtExpediente.Text & " fue cedido a AKCP EUROPE SCSP" & vbCrLf
        'If portfolio = "Crisalida" Then
        'Else
        '    prf2.Text = vbCrLf & "Nos dirigimos a usted para recordarle que " & portfolioText & ", cedió a AKCP EUROPE SCSP el crédito que tenía pendiente con usted, bajo la referencia: " & expediente & "." & vbCrLf
        'End If


        ' Párrafo 2 
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf3 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf3.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf3.Text = vbCrLf & "En virtud de dicha cesión, AKCP EUROPE SCSP certifica que " & nombre & ", identificado con el NIF " & nif.ToUpper & " y en calidad de titular, mantenía una deuda pendiente derivada del contrato anteriormente indicado." & vbCrLf


        ' Párrafo 3
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf4 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf4.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf4.Text = vbCrLf & "Teniendo en cuenta lo anterior, y en respuesta a su solicitud por escrito sobre el estado actual de la deuda mencionada, nos complace informarle que:" & vbCrLf

        ' Párrafo 4
        doc.Content.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
        Dim prf5 As Microsoft.Office.Interop.Word.Range = doc.Range(doc.Content.End - 1, doc.Content.End)
        prf5.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
        prf5.Text = vbCrLf & "SU DEUDA HA SIDO TOTALMENTE CANCELADA. Por consiguiente, en relación con el contrato mencionado anteriormente, no se adeuda cantidad alguna." & vbCrLf
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

        doc.SaveAs(Path.Combine(destinationFolderPath, "Cancelacion_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim.ToUpper & ".docx"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault)
        doc.SaveAs(Path.Combine(desktopFolderPath, "Cancelacion_" & fechaHoy & "_" & idCliente.ToString() & "_" & txtDNI.Text.Trim.ToUpper & ".pdf"), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF)

        doc.Close()
        wordApp.Quit()

        MessageBox.Show("Guardado en el ESCRITORIO")

    End Sub

    Private Sub DescripcionCarteraHelloLetter()

        Dim query As String = String.Empty
        Dim listaDatos As New List(Of Dictionary(Of String, String))
        exp = txtExpHelloLetter.Text.Trim
        If IsNumeric(exp) Then
            query = "SELECT Descripcion FROM Clientes AS C JOIN Expedientes AS E ON E.IdCliente = C.idCliente WHERE RefCliente=@exp OR Expediente=@exp"
        Else
            query = "SELECT Descripcion FROM Clientes AS C JOIN Expedientes AS E ON E.IdCliente = C.idCliente WHERE RefCliente=@exp OR RefInterna=@exp"
        End If
        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.Add(New SqlParameter("@exp", exp))
                    connection.Open()
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim registro As New Dictionary(Of String, String)
                            registro("descripcion") = reader("Descripcion").ToString()
                            listaDatos.Add(registro)
                        End While
                        If listaDatos.Count = 0 Then
                            MessageBox.Show("No se encontraron datos para el expediente proporcionado.")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error: " & ex.Message)
        End Try

        If listaDatos.Count > 1 Then
            For Each dato In listaDatos
                cBoxHelloletter.Items.Add(dato("descripcion"))
            Next
        ElseIf listaDatos.Count = 1 Then
            Dim dato As Dictionary(Of String, String) = listaDatos(0)
            cBoxHelloletter.Items.Add(dato("descripcion"))
        End If

    End Sub

    Private Function IdClienteHelloLetter(descripcion As String) As String

        Dim query As String = "SELECT idCliente FROM Clientes WHERE Descripcion=@descripcion"
        Dim idCliente As String = Nothing

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.Add("@descripcion", SqlDbType.VarChar).Value = descripcion
                    connection.Open()
                    Try
                        Using reader As SqlDataReader = command.ExecuteReader()
                            If reader.Read() Then
                                idCliente = reader("idCliente").ToString()
                            End If
                            If String.IsNullOrEmpty(idCliente) Then
                                MessageBox.Show("No se encontró el IdCliente para la cartera proporcionada.")
                            End If
                        End Using
                    Catch ex As Exception
                        MessageBox.Show("Ocurrió un error: " & ex.Message)
                    End Try
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Ocurrió un error: " & ex.Message)
        End Try

        Return If(String.IsNullOrEmpty(idCliente), Nothing, idCliente)

    End Function

    Private Function BuscarArchivoHelloLetter(idCliente As String, expediente As String) As String

        Dim rutaCarpetaPrincipal As String = "\\192.168.50.46\e\PBO\HelloLetters"

        Try
            If Not Directory.Exists(rutaCarpetaPrincipal) Then
                MessageBox.Show("La carpeta principal especificada no existe.")
                Return Nothing
            End If

            'Dim carpetasCliente As String() = Directory.GetDirectories(rutaCarpetaPrincipal, idCliente, SearchOption.AllDirectories)
            Dim carpetasCliente As String() = Directory.GetDirectories(rutaCarpetaPrincipal, idCliente)

            If carpetasCliente.Length = 0 Then
                MessageBox.Show("No se encontró una carpeta para el idCliente proporcionado.")
                Return Nothing
            Else
                Dim thread As New Thread(AddressOf Mensaje)
                thread.Start()
            End If

            Dim carpetaCliente As String = carpetasCliente(0)

            Dim archivos As String() = Directory.GetFiles(carpetaCliente, "*.pdf", SearchOption.AllDirectories)

            For Each archivo In archivos
                Dim nombreArchivo As String = Path.GetFileNameWithoutExtension(archivo)

                If nombreArchivo.Contains($"_{expediente}") Then
                    Return archivo
                End If
            Next

            MessageBox.Show("No se encontró ningún archivo con el expediente proporcionado.")
            Return Nothing

        Catch ex As Exception
            MessageBox.Show("Ocurrió un error durante la búsqueda del archivo: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Sub Helloletter()

        If cBoxHelloletter.SelectedItem IsNot Nothing Then
            idCliente = Nothing
            idCliente = IdClienteHelloLetter(cBoxHelloletter.SelectedItem.ToString)

            If idCliente IsNot Nothing Then
                Dim rutaArchivo As String = BuscarArchivoHelloLetter(idCliente, txtExpHelloLetter.Text)

                If File.Exists(rutaArchivo) Then
                    Dim rutaEscritorio As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    Dim rutaArchivoDestino As String = Path.Combine(rutaEscritorio, Path.GetFileName(rutaArchivo))
                    File.Copy(rutaArchivo, rutaArchivoDestino, True) ' El último parámetro es para sobrescribir si ya existe
                    MessageBox.Show("Archivo descargado en el ESCRITORIO")
                Else
                    MessageBox.Show("Revise primero los datos o informe de este error.")
                End If
            End If
        End If

    End Sub

    Private Sub Mensaje()
        MessageBox.Show("Buscando el archivo...")
    End Sub

    Private Sub Borrar()

        nif = ""
        nombre = String.Empty
        expediente = String.Empty
        refCliente = String.Empty
        importe = 0.0
        idCliente = String.Empty
        contrato = String.Empty
        cliente = String.Empty
        descripcion = String.Empty
        portfolio = String.Empty
        tipo = 0

        txtDNI.Text = "Introduce un DNI"
        txtDNI.ReadOnly = False
        txtExpHelloLetter.Text = "Introduce un Expediente"
        txtExpHelloLetter.ReadOnly = False
        txtNombre.Text = "Nombre"
        txtNombre.ReadOnly = False
        txtExpediente.Text = "Expediente"
        txtExpediente.ReadOnly = False
        txtContrato.Text = "Contrato"
        txtContrato.ReadOnly = False
        txtImporte.Text = "Deuda total"
        txtImporte.ReadOnly = False
        txtImporteQuita.Text = "Introduce la deuda final a pagar"

        txtImportePlazo1.Text = "0.00"
        txtImportePlazo2.Text = "0.00"
        txtImportePlazo3.Text = ""
        txtImportePlazo4.Text = ""
        txtImportePlazo5.Text = ""
        txtImportePlazo6.Text = ""

        txtFechaPlazo1.Text = "1 de enero del 2024"
        txtFechaPlazo2.Text = "1 de febrero del 2024"

        btnMostrar.Visible = True
        btnMostrarExps.Visible = True
        btnBorrar.Visible = False

        pnlAPT.Location = New Drawing.Point(0, 100)
        pnlBotones.Location = New Drawing.Point(10, 115)
        pnlDatos.Visible = False
        pnlBotones.Visible = True
        pnlAPT.Visible = False

        aExpediente.Clear()
        chLbExpedientes.Items.Clear()

        Me.Size = New Size(258, 195)

    End Sub

    Private Sub btnBorrar_Click(sender As Object, e As EventArgs) Handles btnBorrar.Click

        Borrar()
        'System.Windows.Forms.Application.Restart()

    End Sub


End Class