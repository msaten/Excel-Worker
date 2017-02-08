<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWorkExcel
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tbRuta = New System.Windows.Forms.TextBox()
        Me.btnExaminar = New System.Windows.Forms.Button()
        Me.btnExplorarExcel = New System.Windows.Forms.Button()
        Me.lblContingutCela = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'tbRuta
        '
        Me.tbRuta.Location = New System.Drawing.Point(65, 45)
        Me.tbRuta.Name = "tbRuta"
        Me.tbRuta.Size = New System.Drawing.Size(260, 20)
        Me.tbRuta.TabIndex = 0
        '
        'btnExaminar
        '
        Me.btnExaminar.Location = New System.Drawing.Point(350, 45)
        Me.btnExaminar.Name = "btnExaminar"
        Me.btnExaminar.Size = New System.Drawing.Size(75, 23)
        Me.btnExaminar.TabIndex = 1
        Me.btnExaminar.Text = "Examinar"
        Me.btnExaminar.UseVisualStyleBackColor = True
        '
        'btnExplorarExcel
        '
        Me.btnExplorarExcel.Location = New System.Drawing.Point(184, 114)
        Me.btnExplorarExcel.Name = "btnExplorarExcel"
        Me.btnExplorarExcel.Size = New System.Drawing.Size(170, 35)
        Me.btnExplorarExcel.TabIndex = 2
        Me.btnExplorarExcel.Text = "Explorar Excel"
        Me.btnExplorarExcel.UseVisualStyleBackColor = True
        '
        'lblContingutCela
        '
        Me.lblContingutCela.AutoSize = True
        Me.lblContingutCela.Location = New System.Drawing.Point(184, 186)
        Me.lblContingutCela.Name = "lblContingutCela"
        Me.lblContingutCela.Size = New System.Drawing.Size(39, 13)
        Me.lblContingutCela.TabIndex = 3
        Me.lblContingutCela.Text = "Label1"
        '
        'frmWorkExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(555, 261)
        Me.Controls.Add(Me.lblContingutCela)
        Me.Controls.Add(Me.btnExplorarExcel)
        Me.Controls.Add(Me.btnExaminar)
        Me.Controls.Add(Me.tbRuta)
        Me.Name = "frmWorkExcel"
        Me.Text = "Explora Excel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbRuta As TextBox
    Friend WithEvents btnExaminar As Button
    Friend WithEvents btnExplorarExcel As Button
    Friend WithEvents lblContingutCela As Label
End Class
