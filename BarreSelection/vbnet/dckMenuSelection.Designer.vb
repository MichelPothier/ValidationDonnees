<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dckMenuSelection
    Inherits System.Windows.Forms.UserControl

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tabMenuSelection = New System.Windows.Forms.TabControl()
        Me.pgeParametres = New System.Windows.Forms.TabPage()
        Me.chkZoomGeometrieErreur = New System.Windows.Forms.CheckBox()
        Me.chkCreerClasseErreur = New System.Windows.Forms.CheckBox()
        Me.cboLayerDecoupage = New System.Windows.Forms.ComboBox()
        Me.lblLayerDecoupage = New System.Windows.Forms.Label()
        Me.chkAfficherTable = New System.Windows.Forms.CheckBox()
        Me.txtPrecision = New System.Windows.Forms.TextBox()
        Me.lblPrecision = New System.Windows.Forms.Label()
        Me.btnInitialiser = New System.Windows.Forms.Button()
        Me.tabMenuSelection.SuspendLayout()
        Me.pgeParametres.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabMenuSelection
        '
        Me.tabMenuSelection.Controls.Add(Me.pgeParametres)
        Me.tabMenuSelection.Location = New System.Drawing.Point(12, 0)
        Me.tabMenuSelection.Name = "tabMenuSelection"
        Me.tabMenuSelection.SelectedIndex = 0
        Me.tabMenuSelection.Size = New System.Drawing.Size(278, 265)
        Me.tabMenuSelection.TabIndex = 0
        '
        'pgeParametres
        '
        Me.pgeParametres.Controls.Add(Me.chkZoomGeometrieErreur)
        Me.pgeParametres.Controls.Add(Me.chkCreerClasseErreur)
        Me.pgeParametres.Controls.Add(Me.cboLayerDecoupage)
        Me.pgeParametres.Controls.Add(Me.lblLayerDecoupage)
        Me.pgeParametres.Controls.Add(Me.chkAfficherTable)
        Me.pgeParametres.Controls.Add(Me.txtPrecision)
        Me.pgeParametres.Controls.Add(Me.lblPrecision)
        Me.pgeParametres.Location = New System.Drawing.Point(4, 22)
        Me.pgeParametres.Name = "pgeParametres"
        Me.pgeParametres.Padding = New System.Windows.Forms.Padding(3)
        Me.pgeParametres.Size = New System.Drawing.Size(270, 239)
        Me.pgeParametres.TabIndex = 0
        Me.pgeParametres.Text = "Paramètres"
        Me.pgeParametres.UseVisualStyleBackColor = True
        '
        'chkZoomGeometrieErreur
        '
        Me.chkZoomGeometrieErreur.AutoSize = True
        Me.chkZoomGeometrieErreur.Location = New System.Drawing.Point(12, 90)
        Me.chkZoomGeometrieErreur.Name = "chkZoomGeometrieErreur"
        Me.chkZoomGeometrieErreur.Size = New System.Drawing.Size(196, 17)
        Me.chkZoomGeometrieErreur.TabIndex = 10
        Me.chkZoomGeometrieErreur.Text = "Zoom selon les géométries en erreur"
        Me.chkZoomGeometrieErreur.UseVisualStyleBackColor = True
        '
        'chkCreerClasseErreur
        '
        Me.chkCreerClasseErreur.AutoSize = True
        Me.chkCreerClasseErreur.Checked = True
        Me.chkCreerClasseErreur.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCreerClasseErreur.Location = New System.Drawing.Point(12, 113)
        Me.chkCreerClasseErreur.Name = "chkCreerClasseErreur"
        Me.chkCreerClasseErreur.Size = New System.Drawing.Size(205, 17)
        Me.chkCreerClasseErreur.TabIndex = 9
        Me.chkCreerClasseErreur.Text = "Créer une classe d'erreurs en mémoire"
        Me.chkCreerClasseErreur.UseVisualStyleBackColor = True
        '
        'cboLayerDecoupage
        '
        Me.cboLayerDecoupage.FormattingEnabled = True
        Me.cboLayerDecoupage.Location = New System.Drawing.Point(12, 54)
        Me.cboLayerDecoupage.Name = "cboLayerDecoupage"
        Me.cboLayerDecoupage.Size = New System.Drawing.Size(252, 21)
        Me.cboLayerDecoupage.TabIndex = 8
        '
        'lblLayerDecoupage
        '
        Me.lblLayerDecoupage.AutoSize = True
        Me.lblLayerDecoupage.Location = New System.Drawing.Point(9, 37)
        Me.lblLayerDecoupage.Name = "lblLayerDecoupage"
        Me.lblLayerDecoupage.Size = New System.Drawing.Size(111, 13)
        Me.lblLayerDecoupage.TabIndex = 7
        Me.lblLayerDecoupage.Text = "Layer de découpage :"
        '
        'chkAfficherTable
        '
        Me.chkAfficherTable.AutoSize = True
        Me.chkAfficherTable.Checked = True
        Me.chkAfficherTable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAfficherTable.Location = New System.Drawing.Point(12, 136)
        Me.chkAfficherTable.Name = "chkAfficherTable"
        Me.chkAfficherTable.Size = New System.Drawing.Size(137, 17)
        Me.chkAfficherTable.TabIndex = 6
        Me.chkAfficherTable.Text = "Afficher la table d'erreur"
        Me.chkAfficherTable.UseVisualStyleBackColor = True
        '
        'txtPrecision
        '
        Me.txtPrecision.Location = New System.Drawing.Point(136, 10)
        Me.txtPrecision.Name = "txtPrecision"
        Me.txtPrecision.Size = New System.Drawing.Size(81, 20)
        Me.txtPrecision.TabIndex = 1
        '
        'lblPrecision
        '
        Me.lblPrecision.AutoSize = True
        Me.lblPrecision.Location = New System.Drawing.Point(9, 13)
        Me.lblPrecision.Name = "lblPrecision"
        Me.lblPrecision.Size = New System.Drawing.Size(121, 13)
        Me.lblPrecision.TabIndex = 0
        Me.lblPrecision.Text = "Tolérance de précision :"
        '
        'btnInitialiser
        '
        Me.btnInitialiser.Location = New System.Drawing.Point(12, 267)
        Me.btnInitialiser.Name = "btnInitialiser"
        Me.btnInitialiser.Size = New System.Drawing.Size(82, 25)
        Me.btnInitialiser.TabIndex = 1
        Me.btnInitialiser.Text = "Initialiser"
        Me.btnInitialiser.UseVisualStyleBackColor = True
        '
        'dckMenuSelection
        '
        Me.Controls.Add(Me.btnInitialiser)
        Me.Controls.Add(Me.tabMenuSelection)
        Me.Name = "dckMenuSelection"
        Me.Size = New System.Drawing.Size(300, 300)
        Me.tabMenuSelection.ResumeLayout(False)
        Me.pgeParametres.ResumeLayout(False)
        Me.pgeParametres.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabMenuSelection As System.Windows.Forms.TabControl
    Friend WithEvents pgeParametres As System.Windows.Forms.TabPage
    Friend WithEvents txtPrecision As System.Windows.Forms.TextBox
    Friend WithEvents lblPrecision As System.Windows.Forms.Label
    Friend WithEvents btnInitialiser As System.Windows.Forms.Button
    Friend WithEvents chkAfficherTable As System.Windows.Forms.CheckBox
    Friend WithEvents cboLayerDecoupage As System.Windows.Forms.ComboBox
    Friend WithEvents lblLayerDecoupage As System.Windows.Forms.Label
    Friend WithEvents chkCreerClasseErreur As System.Windows.Forms.CheckBox
    Friend WithEvents chkZoomGeometrieErreur As System.Windows.Forms.CheckBox

End Class
