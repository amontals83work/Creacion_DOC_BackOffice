<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.cBox = New System.Windows.Forms.ComboBox()
        Me.txtDNI = New System.Windows.Forms.TextBox()
        Me.pnlDatos = New System.Windows.Forms.Panel()
        Me.txtImporteQuita = New System.Windows.Forms.TextBox()
        Me.pnlAPT = New System.Windows.Forms.Panel()
        Me.txtFechaPlazo6 = New System.Windows.Forms.TextBox()
        Me.txtImportePlazo6 = New System.Windows.Forms.TextBox()
        Me.txtFechaPlazo5 = New System.Windows.Forms.TextBox()
        Me.txtImportePlazo5 = New System.Windows.Forms.TextBox()
        Me.txtFechaPlazo4 = New System.Windows.Forms.TextBox()
        Me.txtImportePlazo4 = New System.Windows.Forms.TextBox()
        Me.txtFechaPlazo3 = New System.Windows.Forms.TextBox()
        Me.txtImportePlazo3 = New System.Windows.Forms.TextBox()
        Me.txtFechaPlazo2 = New System.Windows.Forms.TextBox()
        Me.txtImportePlazo2 = New System.Windows.Forms.TextBox()
        Me.txtExpediente = New System.Windows.Forms.TextBox()
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.txtContrato = New System.Windows.Forms.TextBox()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.pnlBotones = New System.Windows.Forms.Panel()
        Me.btnMostrarExps = New System.Windows.Forms.Button()
        Me.btnBorrar = New System.Windows.Forms.Button()
        Me.btnMostrar = New System.Windows.Forms.Button()
        Me.btnDescargar = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cBoxCancelacion = New System.Windows.Forms.ComboBox()
        Me.txtExpHelloLetter = New System.Windows.Forms.TextBox()
        Me.cBoxHelloletter = New System.Windows.Forms.ComboBox()
        Me.chLbExpedientes = New System.Windows.Forms.CheckedListBox()
        Me.lblPlazos = New System.Windows.Forms.Label()
        Me.lblPlazoAPT = New System.Windows.Forms.Label()
        Me.lblImportes = New System.Windows.Forms.Label()
        Me.txtFechaPlazo1 = New System.Windows.Forms.TextBox()
        Me.txtImportePlazo1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pnlDatos.SuspendLayout()
        Me.pnlAPT.SuspendLayout()
        Me.pnlBotones.SuspendLayout()
        Me.SuspendLayout()
        '
        'cBox
        '
        Me.cBox.FormattingEnabled = True
        Me.cBox.Location = New System.Drawing.Point(18, 63)
        Me.cBox.Name = "cBox"
        Me.cBox.Size = New System.Drawing.Size(203, 21)
        Me.cBox.TabIndex = 1
        Me.cBox.Text = "Elige una opción"
        '
        'txtDNI
        '
        Me.txtDNI.Location = New System.Drawing.Point(18, 89)
        Me.txtDNI.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDNI.Name = "txtDNI"
        Me.txtDNI.Size = New System.Drawing.Size(203, 20)
        Me.txtDNI.TabIndex = 2
        Me.txtDNI.Text = "Introduce un DNI"
        '
        'pnlDatos
        '
        Me.pnlDatos.Controls.Add(Me.txtImporteQuita)
        Me.pnlDatos.Controls.Add(Me.pnlAPT)
        Me.pnlDatos.Controls.Add(Me.txtFechaPlazo6)
        Me.pnlDatos.Controls.Add(Me.txtImportePlazo6)
        Me.pnlDatos.Controls.Add(Me.txtFechaPlazo5)
        Me.pnlDatos.Controls.Add(Me.txtImportePlazo5)
        Me.pnlDatos.Controls.Add(Me.txtFechaPlazo4)
        Me.pnlDatos.Controls.Add(Me.txtImportePlazo4)
        Me.pnlDatos.Controls.Add(Me.txtFechaPlazo3)
        Me.pnlDatos.Controls.Add(Me.txtImportePlazo3)
        Me.pnlDatos.Controls.Add(Me.txtFechaPlazo2)
        Me.pnlDatos.Controls.Add(Me.txtImportePlazo2)
        Me.pnlDatos.Controls.Add(Me.txtExpediente)
        Me.pnlDatos.Controls.Add(Me.txtImporte)
        Me.pnlDatos.Controls.Add(Me.txtContrato)
        Me.pnlDatos.Controls.Add(Me.txtNombre)
        Me.pnlDatos.Location = New System.Drawing.Point(9, 179)
        Me.pnlDatos.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlDatos.Name = "pnlDatos"
        Me.pnlDatos.Size = New System.Drawing.Size(213, 306)
        Me.pnlDatos.TabIndex = 3
        '
        'txtImporteQuita
        '
        Me.txtImporteQuita.Location = New System.Drawing.Point(8, 97)
        Me.txtImporteQuita.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporteQuita.Name = "txtImporteQuita"
        Me.txtImporteQuita.Size = New System.Drawing.Size(203, 20)
        Me.txtImporteQuita.TabIndex = 8
        Me.txtImporteQuita.Text = "Introduce la deuda final a pagar"
        '
        'pnlAPT
        '
        Me.pnlAPT.Controls.Add(Me.lblPlazoAPT)
        Me.pnlAPT.Controls.Add(Me.lblImportes)
        Me.pnlAPT.Controls.Add(Me.txtFechaPlazo1)
        Me.pnlAPT.Controls.Add(Me.txtImportePlazo1)
        Me.pnlAPT.Controls.Add(Me.lblPlazos)
        Me.pnlAPT.Location = New System.Drawing.Point(0, 122)
        Me.pnlAPT.Name = "pnlAPT"
        Me.pnlAPT.Size = New System.Drawing.Size(213, 65)
        Me.pnlAPT.TabIndex = 9
        '
        'txtFechaPlazo6
        '
        Me.txtFechaPlazo6.Location = New System.Drawing.Point(63, 284)
        Me.txtFechaPlazo6.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFechaPlazo6.Name = "txtFechaPlazo6"
        Me.txtFechaPlazo6.Size = New System.Drawing.Size(148, 20)
        Me.txtFechaPlazo6.TabIndex = 21
        '
        'txtImportePlazo6
        '
        Me.txtImportePlazo6.Location = New System.Drawing.Point(8, 284)
        Me.txtImportePlazo6.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImportePlazo6.Name = "txtImportePlazo6"
        Me.txtImportePlazo6.Size = New System.Drawing.Size(51, 20)
        Me.txtImportePlazo6.TabIndex = 20
        '
        'txtFechaPlazo5
        '
        Me.txtFechaPlazo5.Location = New System.Drawing.Point(63, 260)
        Me.txtFechaPlazo5.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFechaPlazo5.Name = "txtFechaPlazo5"
        Me.txtFechaPlazo5.Size = New System.Drawing.Size(148, 20)
        Me.txtFechaPlazo5.TabIndex = 19
        '
        'txtImportePlazo5
        '
        Me.txtImportePlazo5.Location = New System.Drawing.Point(8, 260)
        Me.txtImportePlazo5.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImportePlazo5.Name = "txtImportePlazo5"
        Me.txtImportePlazo5.Size = New System.Drawing.Size(51, 20)
        Me.txtImportePlazo5.TabIndex = 18
        '
        'txtFechaPlazo4
        '
        Me.txtFechaPlazo4.Location = New System.Drawing.Point(63, 236)
        Me.txtFechaPlazo4.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFechaPlazo4.Name = "txtFechaPlazo4"
        Me.txtFechaPlazo4.Size = New System.Drawing.Size(148, 20)
        Me.txtFechaPlazo4.TabIndex = 17
        '
        'txtImportePlazo4
        '
        Me.txtImportePlazo4.Location = New System.Drawing.Point(8, 236)
        Me.txtImportePlazo4.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImportePlazo4.Name = "txtImportePlazo4"
        Me.txtImportePlazo4.Size = New System.Drawing.Size(51, 20)
        Me.txtImportePlazo4.TabIndex = 16
        '
        'txtFechaPlazo3
        '
        Me.txtFechaPlazo3.Location = New System.Drawing.Point(63, 212)
        Me.txtFechaPlazo3.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFechaPlazo3.Name = "txtFechaPlazo3"
        Me.txtFechaPlazo3.Size = New System.Drawing.Size(148, 20)
        Me.txtFechaPlazo3.TabIndex = 15
        '
        'txtImportePlazo3
        '
        Me.txtImportePlazo3.Location = New System.Drawing.Point(8, 212)
        Me.txtImportePlazo3.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImportePlazo3.Name = "txtImportePlazo3"
        Me.txtImportePlazo3.Size = New System.Drawing.Size(51, 20)
        Me.txtImportePlazo3.TabIndex = 14
        '
        'txtFechaPlazo2
        '
        Me.txtFechaPlazo2.Location = New System.Drawing.Point(63, 188)
        Me.txtFechaPlazo2.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFechaPlazo2.Name = "txtFechaPlazo2"
        Me.txtFechaPlazo2.Size = New System.Drawing.Size(148, 20)
        Me.txtFechaPlazo2.TabIndex = 13
        Me.txtFechaPlazo2.Text = "1 de febrero del 2024"
        '
        'txtImportePlazo2
        '
        Me.txtImportePlazo2.Location = New System.Drawing.Point(8, 188)
        Me.txtImportePlazo2.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImportePlazo2.Name = "txtImportePlazo2"
        Me.txtImportePlazo2.Size = New System.Drawing.Size(51, 20)
        Me.txtImportePlazo2.TabIndex = 12
        Me.txtImportePlazo2.Text = "0,00"
        '
        'txtExpediente
        '
        Me.txtExpediente.Location = New System.Drawing.Point(8, 26)
        Me.txtExpediente.Margin = New System.Windows.Forms.Padding(2)
        Me.txtExpediente.Name = "txtExpediente"
        Me.txtExpediente.ReadOnly = True
        Me.txtExpediente.Size = New System.Drawing.Size(203, 20)
        Me.txtExpediente.TabIndex = 5
        Me.txtExpediente.Text = "Expediente"
        '
        'txtImporte
        '
        Me.txtImporte.Location = New System.Drawing.Point(8, 73)
        Me.txtImporte.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.ReadOnly = True
        Me.txtImporte.Size = New System.Drawing.Size(203, 20)
        Me.txtImporte.TabIndex = 7
        Me.txtImporte.Text = "Deuda total"
        '
        'txtContrato
        '
        Me.txtContrato.Location = New System.Drawing.Point(8, 50)
        Me.txtContrato.Margin = New System.Windows.Forms.Padding(2)
        Me.txtContrato.Name = "txtContrato"
        Me.txtContrato.ReadOnly = True
        Me.txtContrato.Size = New System.Drawing.Size(203, 20)
        Me.txtContrato.TabIndex = 6
        Me.txtContrato.Text = "Contrato"
        '
        'txtNombre
        '
        Me.txtNombre.Location = New System.Drawing.Point(8, 2)
        Me.txtNombre.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.ReadOnly = True
        Me.txtNombre.Size = New System.Drawing.Size(203, 20)
        Me.txtNombre.TabIndex = 4
        Me.txtNombre.Text = "Nombre"
        '
        'pnlBotones
        '
        Me.pnlBotones.Controls.Add(Me.btnBorrar)
        Me.pnlBotones.Controls.Add(Me.btnMostrarExps)
        Me.pnlBotones.Controls.Add(Me.btnMostrar)
        Me.pnlBotones.Controls.Add(Me.btnDescargar)
        Me.pnlBotones.Location = New System.Drawing.Point(9, 114)
        Me.pnlBotones.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlBotones.Name = "pnlBotones"
        Me.pnlBotones.Size = New System.Drawing.Size(213, 59)
        Me.pnlBotones.TabIndex = 1
        '
        'btnMostrarExps
        '
        Me.btnMostrarExps.BackColor = System.Drawing.Color.MistyRose
        Me.btnMostrarExps.FlatAppearance.BorderSize = 0
        Me.btnMostrarExps.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMostrarExps.Location = New System.Drawing.Point(8, 3)
        Me.btnMostrarExps.Name = "btnMostrarExps"
        Me.btnMostrarExps.Size = New System.Drawing.Size(203, 23)
        Me.btnMostrarExps.TabIndex = 34
        Me.btnMostrarExps.Text = "Mostrar Expedientes"
        Me.btnMostrarExps.UseVisualStyleBackColor = False
        '
        'btnBorrar
        '
        Me.btnBorrar.Location = New System.Drawing.Point(8, 33)
        Me.btnBorrar.Name = "btnBorrar"
        Me.btnBorrar.Size = New System.Drawing.Size(203, 23)
        Me.btnBorrar.TabIndex = 32
        Me.btnBorrar.Text = "Borrar"
        Me.btnBorrar.UseVisualStyleBackColor = True
        '
        'btnMostrar
        '
        Me.btnMostrar.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnMostrar.FlatAppearance.BorderSize = 0
        Me.btnMostrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMostrar.Location = New System.Drawing.Point(8, 3)
        Me.btnMostrar.Name = "btnMostrar"
        Me.btnMostrar.Size = New System.Drawing.Size(203, 23)
        Me.btnMostrar.TabIndex = 4
        Me.btnMostrar.Text = "Mostrar Datos"
        Me.btnMostrar.UseVisualStyleBackColor = False
        '
        'btnDescargar
        '
        Me.btnDescargar.BackColor = System.Drawing.Color.FromArgb(CType(CType(212, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.btnDescargar.FlatAppearance.BorderSize = 0
        Me.btnDescargar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDescargar.Location = New System.Drawing.Point(8, 3)
        Me.btnDescargar.Name = "btnDescargar"
        Me.btnDescargar.Size = New System.Drawing.Size(203, 23)
        Me.btnDescargar.TabIndex = 31
        Me.btnDescargar.Text = "Descargar Documento"
        Me.btnDescargar.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(58, 25)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 20)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "BACK OFFICE"
        '
        'cBoxCancelacion
        '
        Me.cBoxCancelacion.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cBoxCancelacion.FormattingEnabled = True
        Me.cBoxCancelacion.Location = New System.Drawing.Point(18, 89)
        Me.cBoxCancelacion.Name = "cBoxCancelacion"
        Me.cBoxCancelacion.Size = New System.Drawing.Size(203, 21)
        Me.cBoxCancelacion.TabIndex = 8
        Me.cBoxCancelacion.Text = "Elige una cartera"
        '
        'txtExpHelloLetter
        '
        Me.txtExpHelloLetter.Location = New System.Drawing.Point(18, 89)
        Me.txtExpHelloLetter.Margin = New System.Windows.Forms.Padding(2)
        Me.txtExpHelloLetter.Name = "txtExpHelloLetter"
        Me.txtExpHelloLetter.Size = New System.Drawing.Size(203, 20)
        Me.txtExpHelloLetter.TabIndex = 1
        Me.txtExpHelloLetter.Text = "Introduce un Expediente"
        '
        'cBoxHelloletter
        '
        Me.cBoxHelloletter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cBoxHelloletter.FormattingEnabled = True
        Me.cBoxHelloletter.ItemHeight = 13
        Me.cBoxHelloletter.Location = New System.Drawing.Point(18, 89)
        Me.cBoxHelloletter.Name = "cBoxHelloletter"
        Me.cBoxHelloletter.Size = New System.Drawing.Size(203, 21)
        Me.cBoxHelloletter.TabIndex = 10
        Me.cBoxHelloletter.Text = "Elige una cartera"
        '
        'chLbExpedientes
        '
        Me.chLbExpedientes.FormattingEnabled = True
        Me.chLbExpedientes.HorizontalScrollbar = True
        Me.chLbExpedientes.Location = New System.Drawing.Point(243, 77)
        Me.chLbExpedientes.Name = "chLbExpedientes"
        Me.chLbExpedientes.Size = New System.Drawing.Size(75, 64)
        Me.chLbExpedientes.TabIndex = 11
        '
        'lblPlazos
        '
        Me.lblPlazos.AutoSize = True
        Me.lblPlazos.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlazos.Location = New System.Drawing.Point(7, 6)
        Me.lblPlazos.Name = "lblPlazos"
        Me.lblPlazos.Size = New System.Drawing.Size(207, 13)
        Me.lblPlazos.TabIndex = 26
        Me.lblPlazos.Text = "PLAZOS DE ACUERDOS DE PAGO"
        '
        'lblPlazoAPT
        '
        Me.lblPlazoAPT.AutoSize = True
        Me.lblPlazoAPT.Location = New System.Drawing.Point(60, 26)
        Me.lblPlazoAPT.Name = "lblPlazoAPT"
        Me.lblPlazoAPT.Size = New System.Drawing.Size(85, 13)
        Me.lblPlazoAPT.TabIndex = 34
        Me.lblPlazoAPT.Text = "Fechas de Pago"
        '
        'lblImportes
        '
        Me.lblImportes.AutoSize = True
        Me.lblImportes.Location = New System.Drawing.Point(8, 26)
        Me.lblImportes.Name = "lblImportes"
        Me.lblImportes.Size = New System.Drawing.Size(47, 13)
        Me.lblImportes.TabIndex = 33
        Me.lblImportes.Text = "Importes"
        '
        'txtFechaPlazo1
        '
        Me.txtFechaPlazo1.Location = New System.Drawing.Point(63, 42)
        Me.txtFechaPlazo1.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFechaPlazo1.Name = "txtFechaPlazo1"
        Me.txtFechaPlazo1.Size = New System.Drawing.Size(148, 20)
        Me.txtFechaPlazo1.TabIndex = 32
        Me.txtFechaPlazo1.Text = "1 de enero del 2024"
        '
        'txtImportePlazo1
        '
        Me.txtImportePlazo1.Location = New System.Drawing.Point(8, 42)
        Me.txtImportePlazo1.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImportePlazo1.Name = "txtImportePlazo1"
        Me.txtImportePlazo1.Size = New System.Drawing.Size(51, 20)
        Me.txtImportePlazo1.TabIndex = 31
        Me.txtImportePlazo1.Text = "0,00"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(243, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Expedientes"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(339, 494)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.chLbExpedientes)
        Me.Controls.Add(Me.cBoxHelloletter)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pnlBotones)
        Me.Controls.Add(Me.pnlDatos)
        Me.Controls.Add(Me.cBox)
        Me.Controls.Add(Me.cBoxCancelacion)
        Me.Controls.Add(Me.txtExpHelloLetter)
        Me.Controls.Add(Me.txtDNI)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Peticiones"
        Me.pnlDatos.ResumeLayout(False)
        Me.pnlDatos.PerformLayout()
        Me.pnlAPT.ResumeLayout(False)
        Me.pnlAPT.PerformLayout()
        Me.pnlBotones.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cBox As ComboBox
    Friend WithEvents txtDNI As TextBox
    Friend WithEvents pnlDatos As Panel
    Friend WithEvents txtNombre As TextBox
    Friend WithEvents pnlBotones As Panel
    Friend WithEvents btnMostrar As Button
    Friend WithEvents btnDescargar As Button
    Friend WithEvents txtContrato As TextBox
    Friend WithEvents txtImporte As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtExpediente As TextBox
    Friend WithEvents txtFechaPlazo6 As TextBox
    Friend WithEvents txtImportePlazo6 As TextBox
    Friend WithEvents txtFechaPlazo5 As TextBox
    Friend WithEvents txtImportePlazo5 As TextBox
    Friend WithEvents txtFechaPlazo4 As TextBox
    Friend WithEvents txtImportePlazo4 As TextBox
    Friend WithEvents txtFechaPlazo3 As TextBox
    Friend WithEvents txtImportePlazo3 As TextBox
    Friend WithEvents txtFechaPlazo2 As TextBox
    Friend WithEvents txtImportePlazo2 As TextBox
    Friend WithEvents btnBorrar As Button
    Friend WithEvents cBoxCancelacion As ComboBox
    Friend WithEvents pnlAPT As Panel
    Friend WithEvents txtImporteQuita As TextBox
    Friend WithEvents txtExpHelloLetter As TextBox
    Friend WithEvents cBoxHelloletter As ComboBox
    Friend WithEvents chLbExpedientes As CheckedListBox
    Friend WithEvents btnMostrarExps As Button
    Friend WithEvents lblPlazos As Label
    Friend WithEvents lblPlazoAPT As Label
    Friend WithEvents lblImportes As Label
    Friend WithEvents txtFechaPlazo1 As TextBox
    Friend WithEvents txtImportePlazo1 As TextBox
    Friend WithEvents Label2 As Label
End Class
