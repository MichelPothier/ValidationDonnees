<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dckMenuContrainteIntegrite
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(dckMenuContrainteIntegrite))
        Me.tooBarreContraintes = New System.Windows.Forms.ToolStrip()
        Me.btnInitialiser = New System.Windows.Forms.ToolStripButton()
        Me.btnCreerTableVide = New System.Windows.Forms.ToolStripButton()
        Me.btnImporter = New System.Windows.Forms.ToolStripButton()
        Me.cboTrierPar = New System.Windows.Forms.ToolStripComboBox()
        Me.cboDecrirePar = New System.Windows.Forms.ToolStripComboBox()
        Me.btnAjouterContrainte = New System.Windows.Forms.ToolStripButton()
        Me.btnAjouterRequete = New System.Windows.Forms.ToolStripButton()
        Me.btnDeplacerRequeteBas = New System.Windows.Forms.ToolStripButton()
        Me.btnDeplacerRequeteHaut = New System.Windows.Forms.ToolStripButton()
        Me.btnEnlever = New System.Windows.Forms.ToolStripButton()
        Me.btnDetruire = New System.Windows.Forms.ToolStripButton()
        Me.btnModifier = New System.Windows.Forms.ToolStripButton()
        Me.btnExecuter = New System.Windows.Forms.ToolStripButton()
        Me.btnExecuterBatch = New System.Windows.Forms.ToolStripButton()
        Me.splContraintes = New System.Windows.Forms.SplitContainer()
        Me.lblInformation = New System.Windows.Forms.GroupBox()
        Me.treContraintes = New System.Windows.Forms.TreeView()
        Me.tabParametres = New System.Windows.Forms.TabControl()
        Me.pgeObligatoire = New System.Windows.Forms.TabPage()
        Me.cboAttributDecoupage = New System.Windows.Forms.ComboBox()
        Me.lblAttributDecoupage = New System.Windows.Forms.Label()
        Me.cboClasseDecoupage = New System.Windows.Forms.ComboBox()
        Me.lblClasseDecoupage = New System.Windows.Forms.Label()
        Me.cboTableContraintes = New System.Windows.Forms.ComboBox()
        Me.lblTableContraintes = New System.Windows.Forms.Label()
        Me.cboGeodatabaseClasses = New System.Windows.Forms.ComboBox()
        Me.lblGeodatabaseClasses = New System.Windows.Forms.Label()
        Me.pgeOptionel = New System.Windows.Forms.TabPage()
        Me.cboCourriel = New System.Windows.Forms.ComboBox()
        Me.lblCourriel = New System.Windows.Forms.Label()
        Me.cboFichierJournal = New System.Windows.Forms.ComboBox()
        Me.lblFichierJournal = New System.Windows.Forms.Label()
        Me.cboRapportErreurs = New System.Windows.Forms.ComboBox()
        Me.lblRapportErreurs = New System.Windows.Forms.Label()
        Me.cboRepertoireErreurs = New System.Windows.Forms.ComboBox()
        Me.lblRepertoireErreurs = New System.Windows.Forms.Label()
        Me.rtbMessages = New System.Windows.Forms.RichTextBox()
        Me.lblMessages = New System.Windows.Forms.Label()
        Me.tooBarreContraintes.SuspendLayout()
        CType(Me.splContraintes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splContraintes.Panel1.SuspendLayout()
        Me.splContraintes.Panel2.SuspendLayout()
        Me.splContraintes.SuspendLayout()
        Me.lblInformation.SuspendLayout()
        Me.tabParametres.SuspendLayout()
        Me.pgeObligatoire.SuspendLayout()
        Me.pgeOptionel.SuspendLayout()
        Me.SuspendLayout()
        '
        'tooBarreContraintes
        '
        Me.tooBarreContraintes.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnInitialiser, Me.btnCreerTableVide, Me.btnImporter, Me.cboTrierPar, Me.cboDecrirePar, Me.btnAjouterContrainte, Me.btnAjouterRequete, Me.btnDeplacerRequeteBas, Me.btnDeplacerRequeteHaut, Me.btnEnlever, Me.btnDetruire, Me.btnModifier, Me.btnExecuter, Me.btnExecuterBatch})
        Me.tooBarreContraintes.Location = New System.Drawing.Point(0, 0)
        Me.tooBarreContraintes.Name = "tooBarreContraintes"
        Me.tooBarreContraintes.Size = New System.Drawing.Size(550, 25)
        Me.tooBarreContraintes.TabIndex = 9
        Me.tooBarreContraintes.Text = "Barre du menu des contraintes."
        '
        'btnInitialiser
        '
        Me.btnInitialiser.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnInitialiser.Image = CType(resources.GetObject("btnInitialiser.Image"), System.Drawing.Image)
        Me.btnInitialiser.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnInitialiser.Name = "btnInitialiser"
        Me.btnInitialiser.Size = New System.Drawing.Size(23, 22)
        Me.btnInitialiser.Text = "Initialiser toutes les tables d'intégrité."
        '
        'btnCreerTableVide
        '
        Me.btnCreerTableVide.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnCreerTableVide.Image = CType(resources.GetObject("btnCreerTableVide.Image"), System.Drawing.Image)
        Me.btnCreerTableVide.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCreerTableVide.Name = "btnCreerTableVide"
        Me.btnCreerTableVide.Size = New System.Drawing.Size(23, 22)
        Me.btnCreerTableVide.Text = "Créer une table vide de contraintes d'intégrité spatiales."
        '
        'btnImporter
        '
        Me.btnImporter.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnImporter.Enabled = False
        Me.btnImporter.Image = CType(resources.GetObject("btnImporter.Image"), System.Drawing.Image)
        Me.btnImporter.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnImporter.Name = "btnImporter"
        Me.btnImporter.Size = New System.Drawing.Size(23, 22)
        Me.btnImporter.Text = "Importer les contraintes d'une table."
        '
        'cboTrierPar
        '
        Me.cboTrierPar.BackColor = System.Drawing.SystemColors.Window
        Me.cboTrierPar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTrierPar.Enabled = False
        Me.cboTrierPar.Name = "cboTrierPar"
        Me.cboTrierPar.Size = New System.Drawing.Size(100, 25)
        Me.cboTrierPar.ToolTipText = "Nom de l'attribut utilisé pour trier les contraintes d'intégrité."
        '
        'cboDecrirePar
        '
        Me.cboDecrirePar.BackColor = System.Drawing.SystemColors.Window
        Me.cboDecrirePar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDecrirePar.Enabled = False
        Me.cboDecrirePar.Name = "cboDecrirePar"
        Me.cboDecrirePar.Size = New System.Drawing.Size(100, 25)
        Me.cboDecrirePar.ToolTipText = "Nom de l'attribut utilisé pour décrire les contraintes d'intégrité."
        '
        'btnAjouterContrainte
        '
        Me.btnAjouterContrainte.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnAjouterContrainte.Enabled = False
        Me.btnAjouterContrainte.Image = CType(resources.GetObject("btnAjouterContrainte.Image"), System.Drawing.Image)
        Me.btnAjouterContrainte.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnAjouterContrainte.Name = "btnAjouterContrainte"
        Me.btnAjouterContrainte.Size = New System.Drawing.Size(23, 22)
        Me.btnAjouterContrainte.Text = "Ajouter une contrainte spatiale basée sur la requête spécifiée."
        '
        'btnAjouterRequete
        '
        Me.btnAjouterRequete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnAjouterRequete.Enabled = False
        Me.btnAjouterRequete.Image = CType(resources.GetObject("btnAjouterRequete.Image"), System.Drawing.Image)
        Me.btnAjouterRequete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnAjouterRequete.Name = "btnAjouterRequete"
        Me.btnAjouterRequete.Size = New System.Drawing.Size(23, 22)
        Me.btnAjouterRequete.Text = "Ajouter la requête spécifiée avant celle sélectionnée."
        '
        'btnDeplacerRequeteBas
        '
        Me.btnDeplacerRequeteBas.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnDeplacerRequeteBas.Enabled = False
        Me.btnDeplacerRequeteBas.Image = CType(resources.GetObject("btnDeplacerRequeteBas.Image"), System.Drawing.Image)
        Me.btnDeplacerRequeteBas.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnDeplacerRequeteBas.Name = "btnDeplacerRequeteBas"
        Me.btnDeplacerRequeteBas.Size = New System.Drawing.Size(23, 22)
        Me.btnDeplacerRequeteBas.Text = "Déplacer la requête vers le bas."
        '
        'btnDeplacerRequeteHaut
        '
        Me.btnDeplacerRequeteHaut.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnDeplacerRequeteHaut.Enabled = False
        Me.btnDeplacerRequeteHaut.Image = CType(resources.GetObject("btnDeplacerRequeteHaut.Image"), System.Drawing.Image)
        Me.btnDeplacerRequeteHaut.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnDeplacerRequeteHaut.Name = "btnDeplacerRequeteHaut"
        Me.btnDeplacerRequeteHaut.Size = New System.Drawing.Size(23, 22)
        Me.btnDeplacerRequeteHaut.Text = "Déplacer la requête vers le haut."
        '
        'btnEnlever
        '
        Me.btnEnlever.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnEnlever.Enabled = False
        Me.btnEnlever.Image = CType(resources.GetObject("btnEnlever.Image"), System.Drawing.Image)
        Me.btnEnlever.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnEnlever.Name = "btnEnlever"
        Me.btnEnlever.Size = New System.Drawing.Size(23, 22)
        Me.btnEnlever.Text = "Enlever l'item sélectionné."
        '
        'btnDetruire
        '
        Me.btnDetruire.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnDetruire.Enabled = False
        Me.btnDetruire.Image = CType(resources.GetObject("btnDetruire.Image"), System.Drawing.Image)
        Me.btnDetruire.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnDetruire.Name = "btnDetruire"
        Me.btnDetruire.Size = New System.Drawing.Size(23, 22)
        Me.btnDetruire.Text = "Détruire la contrainte ou la requête sélectionnée."
        '
        'btnModifier
        '
        Me.btnModifier.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnModifier.Enabled = False
        Me.btnModifier.Image = CType(resources.GetObject("btnModifier.Image"), System.Drawing.Image)
        Me.btnModifier.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnModifier.Name = "btnModifier"
        Me.btnModifier.Size = New System.Drawing.Size(23, 22)
        Me.btnModifier.Text = "Modifier la valeur de l'attribut sélectionné."
        '
        'btnExecuter
        '
        Me.btnExecuter.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnExecuter.Enabled = False
        Me.btnExecuter.Image = CType(resources.GetObject("btnExecuter.Image"), System.Drawing.Image)
        Me.btnExecuter.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnExecuter.Name = "btnExecuter"
        Me.btnExecuter.Size = New System.Drawing.Size(23, 22)
        Me.btnExecuter.Text = "Exécuter la requête ou la contrainte sélectionnée."
        '
        'btnExecuterBatch
        '
        Me.btnExecuterBatch.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnExecuterBatch.Enabled = False
        Me.btnExecuterBatch.Image = CType(resources.GetObject("btnExecuterBatch.Image"), System.Drawing.Image)
        Me.btnExecuterBatch.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnExecuterBatch.Name = "btnExecuterBatch"
        Me.btnExecuterBatch.Size = New System.Drawing.Size(23, 22)
        Me.btnExecuterBatch.Text = "Exécuter les contraintes sélectionnées en arrière plan."
        '
        'splContraintes
        '
        Me.splContraintes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.splContraintes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splContraintes.Location = New System.Drawing.Point(0, 25)
        Me.splContraintes.Name = "splContraintes"
        Me.splContraintes.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splContraintes.Panel1
        '
        Me.splContraintes.Panel1.Controls.Add(Me.lblInformation)
        Me.splContraintes.Panel1.Controls.Add(Me.tabParametres)
        '
        'splContraintes.Panel2
        '
        Me.splContraintes.Panel2.Controls.Add(Me.rtbMessages)
        Me.splContraintes.Panel2.Controls.Add(Me.lblMessages)
        Me.splContraintes.Size = New System.Drawing.Size(550, 375)
        Me.splContraintes.SplitterDistance = 299
        Me.splContraintes.TabIndex = 19
        '
        'lblInformation
        '
        Me.lblInformation.Controls.Add(Me.treContraintes)
        Me.lblInformation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblInformation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInformation.Location = New System.Drawing.Point(0, 185)
        Me.lblInformation.Name = "lblInformation"
        Me.lblInformation.Size = New System.Drawing.Size(546, 110)
        Me.lblInformation.TabIndex = 19
        Me.lblInformation.TabStop = False
        Me.lblInformation.Text = "0 Contrainte(s)"
        '
        'treContraintes
        '
        Me.treContraintes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treContraintes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.treContraintes.Location = New System.Drawing.Point(3, 16)
        Me.treContraintes.Name = "treContraintes"
        Me.treContraintes.Size = New System.Drawing.Size(540, 91)
        Me.treContraintes.TabIndex = 0
        '
        'tabParametres
        '
        Me.tabParametres.Controls.Add(Me.pgeObligatoire)
        Me.tabParametres.Controls.Add(Me.pgeOptionel)
        Me.tabParametres.Dock = System.Windows.Forms.DockStyle.Top
        Me.tabParametres.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabParametres.Location = New System.Drawing.Point(0, 0)
        Me.tabParametres.Name = "tabParametres"
        Me.tabParametres.SelectedIndex = 0
        Me.tabParametres.Size = New System.Drawing.Size(546, 185)
        Me.tabParametres.TabIndex = 18
        '
        'pgeObligatoire
        '
        Me.pgeObligatoire.Controls.Add(Me.cboAttributDecoupage)
        Me.pgeObligatoire.Controls.Add(Me.lblAttributDecoupage)
        Me.pgeObligatoire.Controls.Add(Me.cboClasseDecoupage)
        Me.pgeObligatoire.Controls.Add(Me.lblClasseDecoupage)
        Me.pgeObligatoire.Controls.Add(Me.cboTableContraintes)
        Me.pgeObligatoire.Controls.Add(Me.lblTableContraintes)
        Me.pgeObligatoire.Controls.Add(Me.cboGeodatabaseClasses)
        Me.pgeObligatoire.Controls.Add(Me.lblGeodatabaseClasses)
        Me.pgeObligatoire.Location = New System.Drawing.Point(4, 22)
        Me.pgeObligatoire.Name = "pgeObligatoire"
        Me.pgeObligatoire.Padding = New System.Windows.Forms.Padding(3)
        Me.pgeObligatoire.Size = New System.Drawing.Size(538, 159)
        Me.pgeObligatoire.TabIndex = 0
        Me.pgeObligatoire.Text = "Paramètres obligatoires"
        Me.pgeObligatoire.UseVisualStyleBackColor = True
        '
        'cboAttributDecoupage
        '
        Me.cboAttributDecoupage.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboAttributDecoupage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAttributDecoupage.FormattingEnabled = True
        Me.cboAttributDecoupage.Location = New System.Drawing.Point(3, 118)
        Me.cboAttributDecoupage.Name = "cboAttributDecoupage"
        Me.cboAttributDecoupage.Size = New System.Drawing.Size(532, 21)
        Me.cboAttributDecoupage.TabIndex = 21
        '
        'lblAttributDecoupage
        '
        Me.lblAttributDecoupage.AutoSize = True
        Me.lblAttributDecoupage.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblAttributDecoupage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAttributDecoupage.Location = New System.Drawing.Point(3, 105)
        Me.lblAttributDecoupage.Name = "lblAttributDecoupage"
        Me.lblAttributDecoupage.Size = New System.Drawing.Size(118, 13)
        Me.lblAttributDecoupage.TabIndex = 20
        Me.lblAttributDecoupage.Text = "Attribut de découpage :"
        '
        'cboClasseDecoupage
        '
        Me.cboClasseDecoupage.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboClasseDecoupage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboClasseDecoupage.FormattingEnabled = True
        Me.cboClasseDecoupage.Location = New System.Drawing.Point(3, 84)
        Me.cboClasseDecoupage.Name = "cboClasseDecoupage"
        Me.cboClasseDecoupage.Size = New System.Drawing.Size(532, 21)
        Me.cboClasseDecoupage.TabIndex = 19
        '
        'lblClasseDecoupage
        '
        Me.lblClasseDecoupage.AutoSize = True
        Me.lblClasseDecoupage.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblClasseDecoupage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClasseDecoupage.Location = New System.Drawing.Point(3, 71)
        Me.lblClasseDecoupage.Name = "lblClasseDecoupage"
        Me.lblClasseDecoupage.Size = New System.Drawing.Size(408, 13)
        Me.lblClasseDecoupage.TabIndex = 18
        Me.lblClasseDecoupage.Text = "Classe de découpage : (Ex. D:\ValiderContraintes\ges_Decoupage_SNRC50K_2.lyr)"
        '
        'cboTableContraintes
        '
        Me.cboTableContraintes.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboTableContraintes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTableContraintes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTableContraintes.FormattingEnabled = True
        Me.cboTableContraintes.Location = New System.Drawing.Point(3, 50)
        Me.cboTableContraintes.Name = "cboTableContraintes"
        Me.cboTableContraintes.Size = New System.Drawing.Size(532, 21)
        Me.cboTableContraintes.TabIndex = 17
        '
        'lblTableContraintes
        '
        Me.lblTableContraintes.AutoSize = True
        Me.lblTableContraintes.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblTableContraintes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTableContraintes.Location = New System.Drawing.Point(3, 37)
        Me.lblTableContraintes.Name = "lblTableContraintes"
        Me.lblTableContraintes.Size = New System.Drawing.Size(115, 13)
        Me.lblTableContraintes.TabIndex = 16
        Me.lblTableContraintes.Text = "Table des contraintes :"
        '
        'cboGeodatabaseClasses
        '
        Me.cboGeodatabaseClasses.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboGeodatabaseClasses.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGeodatabaseClasses.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboGeodatabaseClasses.FormattingEnabled = True
        Me.cboGeodatabaseClasses.Location = New System.Drawing.Point(3, 16)
        Me.cboGeodatabaseClasses.Name = "cboGeodatabaseClasses"
        Me.cboGeodatabaseClasses.Size = New System.Drawing.Size(532, 21)
        Me.cboGeodatabaseClasses.TabIndex = 15
        '
        'lblGeodatabaseClasses
        '
        Me.lblGeodatabaseClasses.AutoSize = True
        Me.lblGeodatabaseClasses.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblGeodatabaseClasses.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGeodatabaseClasses.Location = New System.Drawing.Point(3, 3)
        Me.lblGeodatabaseClasses.Name = "lblGeodatabaseClasses"
        Me.lblGeodatabaseClasses.Size = New System.Drawing.Size(178, 13)
        Me.lblGeodatabaseClasses.TabIndex = 14
        Me.lblGeodatabaseClasses.Text = "Géodatabases des classes à traiter :"
        '
        'pgeOptionel
        '
        Me.pgeOptionel.Controls.Add(Me.cboCourriel)
        Me.pgeOptionel.Controls.Add(Me.lblCourriel)
        Me.pgeOptionel.Controls.Add(Me.cboFichierJournal)
        Me.pgeOptionel.Controls.Add(Me.lblFichierJournal)
        Me.pgeOptionel.Controls.Add(Me.cboRapportErreurs)
        Me.pgeOptionel.Controls.Add(Me.lblRapportErreurs)
        Me.pgeOptionel.Controls.Add(Me.cboRepertoireErreurs)
        Me.pgeOptionel.Controls.Add(Me.lblRepertoireErreurs)
        Me.pgeOptionel.Location = New System.Drawing.Point(4, 22)
        Me.pgeOptionel.Name = "pgeOptionel"
        Me.pgeOptionel.Padding = New System.Windows.Forms.Padding(3)
        Me.pgeOptionel.Size = New System.Drawing.Size(538, 159)
        Me.pgeOptionel.TabIndex = 1
        Me.pgeOptionel.Text = "Paramètres optionels"
        Me.pgeOptionel.UseVisualStyleBackColor = True
        '
        'cboCourriel
        '
        Me.cboCourriel.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboCourriel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCourriel.FormattingEnabled = True
        Me.cboCourriel.Location = New System.Drawing.Point(3, 118)
        Me.cboCourriel.Name = "cboCourriel"
        Me.cboCourriel.Size = New System.Drawing.Size(532, 21)
        Me.cboCourriel.TabIndex = 21
        '
        'lblCourriel
        '
        Me.lblCourriel.AutoSize = True
        Me.lblCourriel.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblCourriel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCourriel.Location = New System.Drawing.Point(3, 105)
        Me.lblCourriel.Name = "lblCourriel"
        Me.lblCourriel.Size = New System.Drawing.Size(48, 13)
        Me.lblCourriel.TabIndex = 20
        Me.lblCourriel.Text = "Courriel :"
        '
        'cboFichierJournal
        '
        Me.cboFichierJournal.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboFichierJournal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFichierJournal.FormattingEnabled = True
        Me.cboFichierJournal.Location = New System.Drawing.Point(3, 84)
        Me.cboFichierJournal.Name = "cboFichierJournal"
        Me.cboFichierJournal.Size = New System.Drawing.Size(532, 21)
        Me.cboFichierJournal.TabIndex = 19
        '
        'lblFichierJournal
        '
        Me.lblFichierJournal.AutoSize = True
        Me.lblFichierJournal.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblFichierJournal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFichierJournal.Location = New System.Drawing.Point(3, 71)
        Me.lblFichierJournal.Name = "lblFichierJournal"
        Me.lblFichierJournal.Size = New System.Drawing.Size(104, 13)
        Me.lblFichierJournal.TabIndex = 18
        Me.lblFichierJournal.Text = "Journal d'exécution :"
        '
        'cboRapportErreurs
        '
        Me.cboRapportErreurs.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboRapportErreurs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRapportErreurs.FormattingEnabled = True
        Me.cboRapportErreurs.Location = New System.Drawing.Point(3, 50)
        Me.cboRapportErreurs.Name = "cboRapportErreurs"
        Me.cboRapportErreurs.Size = New System.Drawing.Size(532, 21)
        Me.cboRapportErreurs.TabIndex = 17
        '
        'lblRapportErreurs
        '
        Me.lblRapportErreurs.AutoSize = True
        Me.lblRapportErreurs.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblRapportErreurs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRapportErreurs.Location = New System.Drawing.Point(3, 37)
        Me.lblRapportErreurs.Name = "lblRapportErreurs"
        Me.lblRapportErreurs.Size = New System.Drawing.Size(94, 13)
        Me.lblRapportErreurs.TabIndex = 16
        Me.lblRapportErreurs.Text = "Rapport d'erreurs :"
        '
        'cboRepertoireErreurs
        '
        Me.cboRepertoireErreurs.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboRepertoireErreurs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRepertoireErreurs.FormattingEnabled = True
        Me.cboRepertoireErreurs.Location = New System.Drawing.Point(3, 16)
        Me.cboRepertoireErreurs.Name = "cboRepertoireErreurs"
        Me.cboRepertoireErreurs.Size = New System.Drawing.Size(532, 21)
        Me.cboRepertoireErreurs.TabIndex = 15
        '
        'lblRepertoireErreurs
        '
        Me.lblRepertoireErreurs.AutoSize = True
        Me.lblRepertoireErreurs.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblRepertoireErreurs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRepertoireErreurs.Location = New System.Drawing.Point(3, 3)
        Me.lblRepertoireErreurs.Name = "lblRepertoireErreurs"
        Me.lblRepertoireErreurs.Size = New System.Drawing.Size(187, 13)
        Me.lblRepertoireErreurs.TabIndex = 14
        Me.lblRepertoireErreurs.Text = "Répertoire ou Géodatabase d'erreurs :"
        '
        'rtbMessages
        '
        Me.rtbMessages.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbMessages.HideSelection = False
        Me.rtbMessages.Location = New System.Drawing.Point(0, 13)
        Me.rtbMessages.Name = "rtbMessages"
        Me.rtbMessages.Size = New System.Drawing.Size(546, 55)
        Me.rtbMessages.TabIndex = 21
        Me.rtbMessages.Text = ""
        '
        'lblMessages
        '
        Me.lblMessages.AutoSize = True
        Me.lblMessages.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblMessages.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessages.Location = New System.Drawing.Point(0, 0)
        Me.lblMessages.Name = "lblMessages"
        Me.lblMessages.Size = New System.Drawing.Size(132, 13)
        Me.lblMessages.TabIndex = 20
        Me.lblMessages.Text = "Messages d'exécution"
        '
        'dckMenuContrainteIntegrite
        '
        Me.Controls.Add(Me.splContraintes)
        Me.Controls.Add(Me.tooBarreContraintes)
        Me.Name = "dckMenuContrainteIntegrite"
        Me.Size = New System.Drawing.Size(550, 400)
        Me.tooBarreContraintes.ResumeLayout(False)
        Me.tooBarreContraintes.PerformLayout()
        Me.splContraintes.Panel1.ResumeLayout(False)
        Me.splContraintes.Panel2.ResumeLayout(False)
        Me.splContraintes.Panel2.PerformLayout()
        CType(Me.splContraintes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splContraintes.ResumeLayout(False)
        Me.lblInformation.ResumeLayout(False)
        Me.tabParametres.ResumeLayout(False)
        Me.pgeObligatoire.ResumeLayout(False)
        Me.pgeObligatoire.PerformLayout()
        Me.pgeOptionel.ResumeLayout(False)
        Me.pgeOptionel.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tooBarreContraintes As System.Windows.Forms.ToolStrip
    Friend WithEvents btnImporter As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnCreerTableVide As System.Windows.Forms.ToolStripButton
    Friend WithEvents cboTrierPar As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents cboDecrirePar As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents btnAjouterContrainte As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnAjouterRequete As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnDetruire As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnModifier As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnExecuter As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnInitialiser As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnExecuterBatch As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnDeplacerRequeteBas As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnDeplacerRequeteHaut As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnEnlever As System.Windows.Forms.ToolStripButton
    Friend WithEvents splContraintes As System.Windows.Forms.SplitContainer
    Friend WithEvents rtbMessages As System.Windows.Forms.RichTextBox
    Friend WithEvents lblMessages As System.Windows.Forms.Label
    Friend WithEvents lblInformation As System.Windows.Forms.GroupBox
    Friend WithEvents treContraintes As System.Windows.Forms.TreeView
    Friend WithEvents tabParametres As System.Windows.Forms.TabControl
    Friend WithEvents pgeObligatoire As System.Windows.Forms.TabPage
    Friend WithEvents cboAttributDecoupage As System.Windows.Forms.ComboBox
    Friend WithEvents lblAttributDecoupage As System.Windows.Forms.Label
    Friend WithEvents cboClasseDecoupage As System.Windows.Forms.ComboBox
    Friend WithEvents lblClasseDecoupage As System.Windows.Forms.Label
    Friend WithEvents cboTableContraintes As System.Windows.Forms.ComboBox
    Friend WithEvents lblTableContraintes As System.Windows.Forms.Label
    Friend WithEvents cboGeodatabaseClasses As System.Windows.Forms.ComboBox
    Friend WithEvents lblGeodatabaseClasses As System.Windows.Forms.Label
    Friend WithEvents pgeOptionel As System.Windows.Forms.TabPage
    Friend WithEvents cboCourriel As System.Windows.Forms.ComboBox
    Friend WithEvents lblCourriel As System.Windows.Forms.Label
    Friend WithEvents cboFichierJournal As System.Windows.Forms.ComboBox
    Friend WithEvents lblFichierJournal As System.Windows.Forms.Label
    Friend WithEvents cboRapportErreurs As System.Windows.Forms.ComboBox
    Friend WithEvents lblRapportErreurs As System.Windows.Forms.Label
    Friend WithEvents cboRepertoireErreurs As System.Windows.Forms.ComboBox
    Friend WithEvents lblRepertoireErreurs As System.Windows.Forms.Label

End Class
