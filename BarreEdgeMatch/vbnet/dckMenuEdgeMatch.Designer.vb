<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dckMenuEdgeMatch
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
        Me.tabEdgeMatch = New System.Windows.Forms.TabControl()
        Me.pgeClasses = New System.Windows.Forms.TabPage()
        Me.lblClasses = New System.Windows.Forms.Label()
        Me.lstClasses = New System.Windows.Forms.ListBox()
        Me.cboIdentifiantDecoupage = New System.Windows.Forms.ComboBox()
        Me.lblIdentifiantDecoupage = New System.Windows.Forms.Label()
        Me.lblClasseDecoupage = New System.Windows.Forms.Label()
        Me.cboClasseDecoupage = New System.Windows.Forms.ComboBox()
        Me.pgeParametres = New System.Windows.Forms.TabPage()
        Me.chkIdentifiantPareil = New System.Windows.Forms.CheckBox()
        Me.lblAttributAdjacence = New System.Windows.Forms.Label()
        Me.lstAttributAdjacence = New System.Windows.Forms.ListBox()
        Me.chkClasseDifferente = New System.Windows.Forms.CheckBox()
        Me.chkAdjacenceUnique = New System.Windows.Forms.CheckBox()
        Me.lblPrecision = New System.Windows.Forms.Label()
        Me.lblTolRecherche = New System.Windows.Forms.Label()
        Me.lblTolAdjacence = New System.Windows.Forms.Label()
        Me.txtPrecision = New System.Windows.Forms.TextBox()
        Me.txtTolRecherche = New System.Windows.Forms.TextBox()
        Me.txtTolAdjacence = New System.Windows.Forms.TextBox()
        Me.pgePoints = New System.Windows.Forms.TabPage()
        Me.chkOuvrirPoints = New System.Windows.Forms.CheckBox()
        Me.trePoints = New System.Windows.Forms.TreeView()
        Me.pgeErreurs = New System.Windows.Forms.TabPage()
        Me.chkOuvrirErreurs = New System.Windows.Forms.CheckBox()
        Me.treErreurs = New System.Windows.Forms.TreeView()
        Me.chkAttribut = New System.Windows.Forms.CheckBox()
        Me.chkAdjacence = New System.Windows.Forms.CheckBox()
        Me.chkPrecision = New System.Windows.Forms.CheckBox()
        Me.cmdInitialiser = New System.Windows.Forms.Button()
        Me.lblNbErreur = New System.Windows.Forms.Label()
        Me.lblNbPointAdjacence = New System.Windows.Forms.Label()
        Me.tabEdgeMatch.SuspendLayout()
        Me.pgeClasses.SuspendLayout()
        Me.pgeParametres.SuspendLayout()
        Me.pgePoints.SuspendLayout()
        Me.pgeErreurs.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabEdgeMatch
        '
        Me.tabEdgeMatch.Controls.Add(Me.pgeClasses)
        Me.tabEdgeMatch.Controls.Add(Me.pgeParametres)
        Me.tabEdgeMatch.Controls.Add(Me.pgePoints)
        Me.tabEdgeMatch.Controls.Add(Me.pgeErreurs)
        Me.tabEdgeMatch.Location = New System.Drawing.Point(3, 3)
        Me.tabEdgeMatch.Name = "tabEdgeMatch"
        Me.tabEdgeMatch.SelectedIndex = 0
        Me.tabEdgeMatch.Size = New System.Drawing.Size(297, 271)
        Me.tabEdgeMatch.TabIndex = 0
        '
        'pgeClasses
        '
        Me.pgeClasses.Controls.Add(Me.lblClasses)
        Me.pgeClasses.Controls.Add(Me.lstClasses)
        Me.pgeClasses.Controls.Add(Me.cboIdentifiantDecoupage)
        Me.pgeClasses.Controls.Add(Me.lblIdentifiantDecoupage)
        Me.pgeClasses.Controls.Add(Me.lblClasseDecoupage)
        Me.pgeClasses.Controls.Add(Me.cboClasseDecoupage)
        Me.pgeClasses.Location = New System.Drawing.Point(4, 22)
        Me.pgeClasses.Name = "pgeClasses"
        Me.pgeClasses.Padding = New System.Windows.Forms.Padding(3)
        Me.pgeClasses.Size = New System.Drawing.Size(289, 245)
        Me.pgeClasses.TabIndex = 0
        Me.pgeClasses.Text = "Classes"
        Me.pgeClasses.UseVisualStyleBackColor = True
        '
        'lblClasses
        '
        Me.lblClasses.AutoSize = True
        Me.lblClasses.Location = New System.Drawing.Point(3, 102)
        Me.lblClasses.Name = "lblClasses"
        Me.lblClasses.Size = New System.Drawing.Size(140, 13)
        Me.lblClasses.TabIndex = 5
        Me.lblClasses.Text = "Liste des classes d'éléments"
        '
        'lstClasses
        '
        Me.lstClasses.FormattingEnabled = True
        Me.lstClasses.Location = New System.Drawing.Point(6, 119)
        Me.lstClasses.Name = "lstClasses"
        Me.lstClasses.Size = New System.Drawing.Size(277, 121)
        Me.lstClasses.TabIndex = 4
        '
        'cboIdentifiantDecoupage
        '
        Me.cboIdentifiantDecoupage.FormattingEnabled = True
        Me.cboIdentifiantDecoupage.Location = New System.Drawing.Point(6, 71)
        Me.cboIdentifiantDecoupage.Name = "cboIdentifiantDecoupage"
        Me.cboIdentifiantDecoupage.Size = New System.Drawing.Size(277, 21)
        Me.cboIdentifiantDecoupage.TabIndex = 3
        '
        'lblIdentifiantDecoupage
        '
        Me.lblIdentifiantDecoupage.AutoSize = True
        Me.lblIdentifiantDecoupage.Location = New System.Drawing.Point(3, 55)
        Me.lblIdentifiantDecoupage.Name = "lblIdentifiantDecoupage"
        Me.lblIdentifiantDecoupage.Size = New System.Drawing.Size(112, 13)
        Me.lblIdentifiantDecoupage.TabIndex = 2
        Me.lblIdentifiantDecoupage.Text = "Attribut de découpage"
        '
        'lblClasseDecoupage
        '
        Me.lblClasseDecoupage.AutoSize = True
        Me.lblClasseDecoupage.Location = New System.Drawing.Point(3, 8)
        Me.lblClasseDecoupage.Name = "lblClasseDecoupage"
        Me.lblClasseDecoupage.Size = New System.Drawing.Size(110, 13)
        Me.lblClasseDecoupage.TabIndex = 1
        Me.lblClasseDecoupage.Text = "Classe de découpage"
        '
        'cboClasseDecoupage
        '
        Me.cboClasseDecoupage.FormattingEnabled = True
        Me.cboClasseDecoupage.Location = New System.Drawing.Point(6, 25)
        Me.cboClasseDecoupage.Name = "cboClasseDecoupage"
        Me.cboClasseDecoupage.Size = New System.Drawing.Size(277, 21)
        Me.cboClasseDecoupage.TabIndex = 0
        '
        'pgeParametres
        '
        Me.pgeParametres.Controls.Add(Me.chkIdentifiantPareil)
        Me.pgeParametres.Controls.Add(Me.lblAttributAdjacence)
        Me.pgeParametres.Controls.Add(Me.lstAttributAdjacence)
        Me.pgeParametres.Controls.Add(Me.chkClasseDifferente)
        Me.pgeParametres.Controls.Add(Me.chkAdjacenceUnique)
        Me.pgeParametres.Controls.Add(Me.lblPrecision)
        Me.pgeParametres.Controls.Add(Me.lblTolRecherche)
        Me.pgeParametres.Controls.Add(Me.lblTolAdjacence)
        Me.pgeParametres.Controls.Add(Me.txtPrecision)
        Me.pgeParametres.Controls.Add(Me.txtTolRecherche)
        Me.pgeParametres.Controls.Add(Me.txtTolAdjacence)
        Me.pgeParametres.Location = New System.Drawing.Point(4, 22)
        Me.pgeParametres.Name = "pgeParametres"
        Me.pgeParametres.Padding = New System.Windows.Forms.Padding(3)
        Me.pgeParametres.Size = New System.Drawing.Size(289, 245)
        Me.pgeParametres.TabIndex = 1
        Me.pgeParametres.Text = "Paramètres"
        Me.pgeParametres.UseVisualStyleBackColor = True
        '
        'chkIdentifiantPareil
        '
        Me.chkIdentifiantPareil.AutoSize = True
        Me.chkIdentifiantPareil.Location = New System.Drawing.Point(6, 108)
        Me.chkIdentifiantPareil.Name = "chkIdentifiantPareil"
        Me.chkIdentifiantPareil.Size = New System.Drawing.Size(150, 17)
        Me.chkIdentifiantPareil.TabIndex = 10
        Me.chkIdentifiantPareil.Text = "Adjacence sans identifiant"
        Me.chkIdentifiantPareil.UseVisualStyleBackColor = True
        '
        'lblAttributAdjacence
        '
        Me.lblAttributAdjacence.AutoSize = True
        Me.lblAttributAdjacence.Location = New System.Drawing.Point(3, 155)
        Me.lblAttributAdjacence.Name = "lblAttributAdjacence"
        Me.lblAttributAdjacence.Size = New System.Drawing.Size(150, 13)
        Me.lblAttributAdjacence.TabIndex = 9
        Me.lblAttributAdjacence.Text = "Liste des attributs d'adjacence"
        '
        'lstAttributAdjacence
        '
        Me.lstAttributAdjacence.FormattingEnabled = True
        Me.lstAttributAdjacence.Location = New System.Drawing.Point(6, 172)
        Me.lstAttributAdjacence.Name = "lstAttributAdjacence"
        Me.lstAttributAdjacence.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstAttributAdjacence.Size = New System.Drawing.Size(277, 69)
        Me.lstAttributAdjacence.TabIndex = 8
        '
        'chkClasseDifferente
        '
        Me.chkClasseDifferente.AutoSize = True
        Me.chkClasseDifferente.Location = New System.Drawing.Point(6, 131)
        Me.chkClasseDifferente.Name = "chkClasseDifferente"
        Me.chkClasseDifferente.Size = New System.Drawing.Size(187, 17)
        Me.chkClasseDifferente.TabIndex = 7
        Me.chkClasseDifferente.Text = "Adjacence des classes différentes"
        Me.chkClasseDifferente.UseVisualStyleBackColor = True
        '
        'chkAdjacenceUnique
        '
        Me.chkAdjacenceUnique.AutoSize = True
        Me.chkAdjacenceUnique.Location = New System.Drawing.Point(6, 85)
        Me.chkAdjacenceUnique.Name = "chkAdjacenceUnique"
        Me.chkAdjacenceUnique.Size = New System.Drawing.Size(163, 17)
        Me.chkAdjacenceUnique.TabIndex = 6
        Me.chkAdjacenceUnique.Text = "Adjacence unique seulement"
        Me.chkAdjacenceUnique.UseVisualStyleBackColor = True
        '
        'lblPrecision
        '
        Me.lblPrecision.AutoSize = True
        Me.lblPrecision.Location = New System.Drawing.Point(86, 66)
        Me.lblPrecision.Name = "lblPrecision"
        Me.lblPrecision.Size = New System.Drawing.Size(114, 13)
        Me.lblPrecision.TabIndex = 5
        Me.lblPrecision.Text = "Précision des données"
        '
        'lblTolRecherche
        '
        Me.lblTolRecherche.AutoSize = True
        Me.lblTolRecherche.Location = New System.Drawing.Point(84, 40)
        Me.lblTolRecherche.Name = "lblTolRecherche"
        Me.lblTolRecherche.Size = New System.Drawing.Size(165, 13)
        Me.lblTolRecherche.TabIndex = 4
        Me.lblTolRecherche.Text = "Tolérance de recherche en mètre"
        '
        'lblTolAdjacence
        '
        Me.lblTolAdjacence.AutoSize = True
        Me.lblTolAdjacence.Location = New System.Drawing.Point(84, 14)
        Me.lblTolAdjacence.Name = "lblTolAdjacence"
        Me.lblTolAdjacence.Size = New System.Drawing.Size(160, 13)
        Me.lblTolAdjacence.TabIndex = 3
        Me.lblTolAdjacence.Text = "Tolérance d'adjacence en mètre"
        '
        'txtPrecision
        '
        Me.txtPrecision.Location = New System.Drawing.Point(6, 59)
        Me.txtPrecision.Name = "txtPrecision"
        Me.txtPrecision.Size = New System.Drawing.Size(73, 20)
        Me.txtPrecision.TabIndex = 2
        '
        'txtTolRecherche
        '
        Me.txtTolRecherche.Location = New System.Drawing.Point(6, 33)
        Me.txtTolRecherche.Name = "txtTolRecherche"
        Me.txtTolRecherche.Size = New System.Drawing.Size(73, 20)
        Me.txtTolRecherche.TabIndex = 1
        '
        'txtTolAdjacence
        '
        Me.txtTolAdjacence.Location = New System.Drawing.Point(6, 7)
        Me.txtTolAdjacence.Name = "txtTolAdjacence"
        Me.txtTolAdjacence.Size = New System.Drawing.Size(73, 20)
        Me.txtTolAdjacence.TabIndex = 0
        '
        'pgePoints
        '
        Me.pgePoints.Controls.Add(Me.lblNbPointAdjacence)
        Me.pgePoints.Controls.Add(Me.chkOuvrirPoints)
        Me.pgePoints.Controls.Add(Me.trePoints)
        Me.pgePoints.Location = New System.Drawing.Point(4, 22)
        Me.pgePoints.Name = "pgePoints"
        Me.pgePoints.Size = New System.Drawing.Size(289, 245)
        Me.pgePoints.TabIndex = 2
        Me.pgePoints.Text = "Points"
        Me.pgePoints.UseVisualStyleBackColor = True
        '
        'chkOuvrirPoints
        '
        Me.chkOuvrirPoints.AutoSize = True
        Me.chkOuvrirPoints.Location = New System.Drawing.Point(3, 5)
        Me.chkOuvrirPoints.Name = "chkOuvrirPoints"
        Me.chkOuvrirPoints.Size = New System.Drawing.Size(124, 17)
        Me.chkOuvrirPoints.TabIndex = 1
        Me.chkOuvrirPoints.Text = "Ouvrir tous les points"
        Me.chkOuvrirPoints.UseVisualStyleBackColor = True
        '
        'trePoints
        '
        Me.trePoints.Location = New System.Drawing.Point(3, 28)
        Me.trePoints.Name = "trePoints"
        Me.trePoints.Size = New System.Drawing.Size(283, 197)
        Me.trePoints.TabIndex = 0
        '
        'pgeErreurs
        '
        Me.pgeErreurs.Controls.Add(Me.lblNbErreur)
        Me.pgeErreurs.Controls.Add(Me.chkOuvrirErreurs)
        Me.pgeErreurs.Controls.Add(Me.treErreurs)
        Me.pgeErreurs.Controls.Add(Me.chkAttribut)
        Me.pgeErreurs.Controls.Add(Me.chkAdjacence)
        Me.pgeErreurs.Controls.Add(Me.chkPrecision)
        Me.pgeErreurs.Location = New System.Drawing.Point(4, 22)
        Me.pgeErreurs.Name = "pgeErreurs"
        Me.pgeErreurs.Size = New System.Drawing.Size(289, 245)
        Me.pgeErreurs.TabIndex = 3
        Me.pgeErreurs.Text = "Erreurs"
        Me.pgeErreurs.UseVisualStyleBackColor = True
        '
        'chkOuvrirErreurs
        '
        Me.chkOuvrirErreurs.AutoSize = True
        Me.chkOuvrirErreurs.Location = New System.Drawing.Point(3, 5)
        Me.chkOuvrirErreurs.Name = "chkOuvrirErreurs"
        Me.chkOuvrirErreurs.Size = New System.Drawing.Size(137, 17)
        Me.chkOuvrirErreurs.TabIndex = 4
        Me.chkOuvrirErreurs.Text = "Ouvrir toutes les erreurs"
        Me.chkOuvrirErreurs.UseVisualStyleBackColor = True
        '
        'treErreurs
        '
        Me.treErreurs.Location = New System.Drawing.Point(3, 47)
        Me.treErreurs.Name = "treErreurs"
        Me.treErreurs.Size = New System.Drawing.Size(283, 178)
        Me.treErreurs.TabIndex = 3
        '
        'chkAttribut
        '
        Me.chkAttribut.AutoSize = True
        Me.chkAttribut.Checked = True
        Me.chkAttribut.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAttribut.Location = New System.Drawing.Point(161, 24)
        Me.chkAttribut.Name = "chkAttribut"
        Me.chkAttribut.Size = New System.Drawing.Size(59, 17)
        Me.chkAttribut.TabIndex = 2
        Me.chkAttribut.Text = "Attribut"
        Me.chkAttribut.UseVisualStyleBackColor = True
        '
        'chkAdjacence
        '
        Me.chkAdjacence.AutoSize = True
        Me.chkAdjacence.Checked = True
        Me.chkAdjacence.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAdjacence.Location = New System.Drawing.Point(78, 24)
        Me.chkAdjacence.Name = "chkAdjacence"
        Me.chkAdjacence.Size = New System.Drawing.Size(77, 17)
        Me.chkAdjacence.TabIndex = 1
        Me.chkAdjacence.Text = "Adjacence"
        Me.chkAdjacence.UseVisualStyleBackColor = True
        '
        'chkPrecision
        '
        Me.chkPrecision.AutoSize = True
        Me.chkPrecision.Checked = True
        Me.chkPrecision.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPrecision.Location = New System.Drawing.Point(3, 24)
        Me.chkPrecision.Name = "chkPrecision"
        Me.chkPrecision.Size = New System.Drawing.Size(69, 17)
        Me.chkPrecision.TabIndex = 0
        Me.chkPrecision.Text = "Précision"
        Me.chkPrecision.UseVisualStyleBackColor = True
        '
        'cmdInitialiser
        '
        Me.cmdInitialiser.Location = New System.Drawing.Point(3, 274)
        Me.cmdInitialiser.Name = "cmdInitialiser"
        Me.cmdInitialiser.Size = New System.Drawing.Size(101, 23)
        Me.cmdInitialiser.TabIndex = 1
        Me.cmdInitialiser.Text = "Initialiser"
        Me.cmdInitialiser.UseVisualStyleBackColor = True
        '
        'lblNbErreur
        '
        Me.lblNbErreur.AutoSize = True
        Me.lblNbErreur.Location = New System.Drawing.Point(3, 228)
        Me.lblNbErreur.Name = "lblNbErreur"
        Me.lblNbErreur.Size = New System.Drawing.Size(122, 13)
        Me.lblNbErreur.TabIndex = 5
        Me.lblNbErreur.Text = "Nombre total d'erreurs: 0"
        '
        'lblNbPointAdjacence
        '
        Me.lblNbPointAdjacence.AutoSize = True
        Me.lblNbPointAdjacence.Location = New System.Drawing.Point(3, 228)
        Me.lblNbPointAdjacence.Name = "lblNbPointAdjacence"
        Me.lblNbPointAdjacence.Size = New System.Drawing.Size(164, 13)
        Me.lblNbPointAdjacence.TabIndex = 2
        Me.lblNbPointAdjacence.Text = "Nombre de points d'Adjacence: 0"
        '
        'dckMenuEdgeMatch
        '
        Me.Controls.Add(Me.cmdInitialiser)
        Me.Controls.Add(Me.tabEdgeMatch)
        Me.Name = "dckMenuEdgeMatch"
        Me.Size = New System.Drawing.Size(300, 300)
        Me.tabEdgeMatch.ResumeLayout(False)
        Me.pgeClasses.ResumeLayout(False)
        Me.pgeClasses.PerformLayout()
        Me.pgeParametres.ResumeLayout(False)
        Me.pgeParametres.PerformLayout()
        Me.pgePoints.ResumeLayout(False)
        Me.pgePoints.PerformLayout()
        Me.pgeErreurs.ResumeLayout(False)
        Me.pgeErreurs.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabEdgeMatch As System.Windows.Forms.TabControl
    Friend WithEvents pgeClasses As System.Windows.Forms.TabPage
    Friend WithEvents pgeParametres As System.Windows.Forms.TabPage
    Friend WithEvents lblIdentifiantDecoupage As System.Windows.Forms.Label
    Friend WithEvents lblClasseDecoupage As System.Windows.Forms.Label
    Friend WithEvents cboClasseDecoupage As System.Windows.Forms.ComboBox
    Friend WithEvents pgePoints As System.Windows.Forms.TabPage
    Friend WithEvents trePoints As System.Windows.Forms.TreeView
    Friend WithEvents pgeErreurs As System.Windows.Forms.TabPage
    Friend WithEvents treErreurs As System.Windows.Forms.TreeView
    Friend WithEvents chkAttribut As System.Windows.Forms.CheckBox
    Friend WithEvents chkAdjacence As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrecision As System.Windows.Forms.CheckBox
    Friend WithEvents lblClasses As System.Windows.Forms.Label
    Friend WithEvents lstClasses As System.Windows.Forms.ListBox
    Friend WithEvents cboIdentifiantDecoupage As System.Windows.Forms.ComboBox
    Friend WithEvents lblAttributAdjacence As System.Windows.Forms.Label
    Friend WithEvents lstAttributAdjacence As System.Windows.Forms.ListBox
    Friend WithEvents chkClasseDifferente As System.Windows.Forms.CheckBox
    Friend WithEvents chkAdjacenceUnique As System.Windows.Forms.CheckBox
    Friend WithEvents lblPrecision As System.Windows.Forms.Label
    Friend WithEvents lblTolRecherche As System.Windows.Forms.Label
    Friend WithEvents lblTolAdjacence As System.Windows.Forms.Label
    Friend WithEvents txtPrecision As System.Windows.Forms.TextBox
    Friend WithEvents txtTolRecherche As System.Windows.Forms.TextBox
    Friend WithEvents txtTolAdjacence As System.Windows.Forms.TextBox
    Friend WithEvents chkOuvrirPoints As System.Windows.Forms.CheckBox
    Friend WithEvents chkOuvrirErreurs As System.Windows.Forms.CheckBox
    Friend WithEvents cmdInitialiser As System.Windows.Forms.Button
    Friend WithEvents chkIdentifiantPareil As System.Windows.Forms.CheckBox
    Friend WithEvents lblNbErreur As System.Windows.Forms.Label
    Friend WithEvents lblNbPointAdjacence As System.Windows.Forms.Label

End Class
