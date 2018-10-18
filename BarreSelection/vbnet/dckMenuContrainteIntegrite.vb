Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Display
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks

''' <summary>
''' Designer class of the dockable window add-in. It contains user interfaces that make up the dockable window.
''' </summary>
Public Class dckMenuContrainteIntegrite
    'Variables de travail
    Private m_hook As Object    'Application appelante
    Private m_AttributTrier As String = "NOM_TABLE" 'Nom de l'attribut utilisé pour trier les contraintes.
    Private m_AttributDecrire As String = "GROUPE" 'Nom de l'attribut utilisé pour décrire les contraintes.
    Private m_AttributRequetes As String = "REQUETES" 'Nom de l'attribut utilisé pour exécuter les requêtes.
    Private m_TrackCancel As ITrackCancel = Nothing 'Interface qui permet d'annuler l'exécution du traitement en inteactif.
    Private m_RichTextBox As RichTextBox = Nothing 'RichTextBox permettant d'afficher l'exécution du tratement en interactif.

#Region "Routines et fonctions d'initialisation"
    Public Sub New(ByVal hook As Object)
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call.
        Me.Hook = hook

        'Définir l'application
        m_Application = CType(hook, IApplication)

        'Définir le document
        m_MxDocument = CType(m_Application.Document, IMxDocument)

        'Initialiser le menu
        Me.Init()

        'Conserver le menu en mémoire
        m_MenuContrainteIntegrite = Me
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        m_hook = Nothing
        m_AttributTrier = Nothing
        m_AttributDecrire = Nothing
        m_AttributRequetes = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
        'Finaliser
        MyBase.Finalize()
    End Sub

    ''' <summary>
    ''' Host object of the dockable window
    ''' </summary> 
    Public Property Hook() As Object
        Get
            Return m_hook
        End Get
        Set(ByVal value As Object)
            m_hook = value
        End Set
    End Property

    ''' <summary>
    ''' Implementation class of the dockable window add-in. It is responsible for
    ''' creating and disposing the user interface class for the dockable window.
    ''' </summary>
    Public Class AddinImpl
        Inherits ESRI.ArcGIS.Desktop.AddIns.DockableWindow

        Private m_windowUI As dckMenuContrainteIntegrite

        Protected Overrides Function OnCreateChild() As System.IntPtr
            m_windowUI = New dckMenuContrainteIntegrite(Me.Hook)
            Return m_windowUI.Handle
        End Function

        Protected Overrides Sub Dispose(ByVal Param As Boolean)
            If m_windowUI IsNot Nothing Then
                m_windowUI.Dispose(Param)
            End If

            MyBase.Dispose(Param)
        End Sub
    End Class
#End Region

#Region "Routines et fonctions d'événements de la barre"
    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnInitialiser_Click(sender As Object, e As EventArgs) Handles btnInitialiser.Click
        Try
            'Initialiser le menu des contraintes d'intégrité
            Me.Init()

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnCreerTableVide_Click(sender As Object, e As EventArgs) Handles btnCreerTableVide.Click
        'Déclarer les variables de travail
        Dim pGxDialog As IGxDialog = Nothing                    'Interface utilisé pour demander le nom de la table à créer.
        Dim pFilterTables As IGxObjectFilter = Nothing          'Interface utilisé pour filtrer seulement les tables.
        Dim pFilterColl As IGxObjectFilterCollection = Nothing  'Interface utilisé pour ajouter les filtres.
        Dim pTable As ITable = Nothing                          'Interface contenant la nouvelle table des contraintes d'intégrité.
        Dim pStandaloneTable As IStandaloneTable = Nothing      'Interface contenant la table des contraintes pouvant contenir une reqûete attributive.
        Dim pStandaloneTableColl As IStandaloneTableCollection = Nothing    'Interface utilisé pour ajouter une table dans la Map active.
        Dim iSelectedIndex As Integer = -1      'Contient le numéro de l'index à sélectionner.

        Try
            'Définir le menu pour demander le nom de la table à créer
            pGxDialog = New GxDialog
            'Définir un filtre pour les tables
            pFilterTables = New GxFilterTables
            'Interface pour ajouter les filtres
            pFilterColl = CType(pGxDialog, IGxObjectFilterCollection)
            'Ajouter le filtre des tables
            pFilterColl.AddFilter(pFilterTables, False)
            'Vérifier si la Géodatabase est valide
            If m_Geodatabase IsNot Nothing Then
                'Définir le workspace par défaut
                pGxDialog.StartingLocation = m_Geodatabase.PathName
            End If
            'Définir le titre du menu
            pGxDialog.Title = "Créer une nouvelle table des contraintes spatiales"
            'Définir le nom de la table par défaut
            pGxDialog.Name = "CONTRAINTE_INTEGRITE_SPATIALE"

            'Vérifier si le nom est spécifié
            If pGxDialog.DoModalSave(0) Then
                'Créer une nouvelle table vide des contraintes d'intégrité
                pTable = CreerNouvelleTableContraintes(pGxDialog.FinalLocation.FullName, pGxDialog.Name)

                'Vérifier si la table a été créée
                If pTable IsNot Nothing Then
                    'Créer une nouvelle StandaloneTable
                    pStandaloneTable = New StandaloneTable
                    'Ajouter la table au Standalone
                    pStandaloneTable.Table = pTable
                    'Interface pour ajouter la table dans la Map active
                    pStandaloneTableColl = CType(m_MxDocument.FocusMap, IStandaloneTableCollection)
                    'Ajouter le Standalone à la Map active
                    pStandaloneTableColl.AddStandaloneTable(pStandaloneTable)
                    'Rafraichir le TableOfContent
                    m_MxDocument.UpdateContents()

                    'Remplir le ComboBox pour définir la table d'intégrité spatiale
                    iSelectedIndex = RemplirComboBoxTable(cboTableContraintes, pGxDialog.FinalLocation.FullName & "\" & pGxDialog.Name)
                    'Sélectionner l'index
                    cboTableContraintes.SelectedIndex = iSelectedIndex

                    'Fermer tous les noeuds du TreeView
                    treContraintes.CollapseAll()

                    'Afficher les contraintes dans un TreeView
                    Call ImporterContraintes(m_TableContraintes, m_AttributTrier, m_AttributDecrire, m_AttributRequetes)
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pGxDialog = Nothing
            pFilterTables = Nothing
            pFilterColl = Nothing
            pTable = Nothing
            pStandaloneTable = Nothing
            pStandaloneTableColl = Nothing
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnImporter_Click(sender As Object, e As EventArgs) Handles btnImporter.Click
        Try
            'Afficher les contraintes dans un TreeView
            Call ImporterContraintes(m_TableContraintes, m_AttributTrier, m_AttributDecrire, m_AttributRequetes)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsque l'item sélectionné change
    Private Sub cboTrierPar_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTrierPar.SelectedIndexChanged
        Try
            'Définir les attributs
            m_AttributTrier = cboTrierPar.Text

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsque l'item sélectionné change
    Private Sub cboDecrirePar_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDecrirePar.SelectedIndexChanged
        Try
            'Définir les attributs
            m_AttributDecrire = cboDecrirePar.Text

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnAjouterContrainte_Click(sender As Object, e As EventArgs) Handles btnAjouterContrainte.Click
        'Déclarer les variables de travail
        Dim oTreeNodeTable As TreeNode = Nothing        'Objet contenant le Node du TreeView pour la table des contraintes.
        Dim oTreeNodeContrainte As TreeNode = Nothing   'Objet contenant le Node du TreeView pour une contrainte d'intégrité.
        Dim oRequete As intRequete          'Objet utilisé pour effectuer la sélection selon la requête.
        Dim pRow As IRow = Nothing                      'Interface contenant l'information d'une contrainte.
        Dim sNomClasseSel As String = ""        'Contient le nom de la classe de sélection.
        Dim sNomClasse As String = ""           'Contient le nom de la classe.
        Dim iPosAttTrier As Integer = -1        'Contient la position de l'attribut pour trier les contraintes.
        Dim iPosAttDecrire As Integer = -1      'Contient la position de l'attribut pour décrire les contraintes.

        Try
            'Définir la requête spécifiée
            oRequete = DefinirRequete()

            'Vérifier si la requête est valide
            If oRequete.EstValide Then
                'Définir le noeud de la table
                oTreeNodeTable = treContraintes.SelectedNode

                'Vérifier si la table correspond à celle active
                If oTreeNodeTable.Text = cboTableContraintes.Text Then
                    'Définir le nom de la classe
                    sNomClasseSel = oRequete.FeatureLayerSelection.FeatureClass.AliasName.ToUpper
                    'vérifier si le nom contient le nom de l'usager
                    If sNomClasseSel.Contains(".") Then
                        'Enlever le nom de l'usager
                        sNomClasseSel = sNomClasseSel.Split(CChar("."))(1)
                    End If
                    'Définir le nom de la classe
                    sNomClasse = m_FeatureLayer.FeatureClass.AliasName.ToUpper
                    'vérifier si le nom contient le nom de l'usager
                    If sNomClasse.Contains(".") Then
                        'Enlever le nom de l'usager
                        sNomClasse = sNomClasse.Split(CChar("."))(1)
                    End If
                    'Ajouter une contrainte dans la table
                    pRow = AjouterContrainte(m_TableContraintes.Table, _
                                             oRequete.Nom & "_" & oRequete.NomAttribut, _
                                             "Valider la contrainte spatiale " & oRequete.Nom & " de la classe " & sNomClasseSel, _
                                             "Corriger la contrainte spatiale " & oRequete.Nom & " de la classe " & sNomClasseSel, _
                                             oRequete.Commande, _
                                             oRequete.Nom, _
                                             sNomClasse)

                    'Définir la position de l'attribut pour regrouper et trier les contraintes
                    iPosAttTrier = m_TableContraintes.Table.FindField(m_AttributTrier)

                    'Définir la position de l'attribut pour décrire les contraintes
                    iPosAttDecrire = m_TableContraintes.Table.FindField(m_AttributDecrire)

                    'Ajouter une contrainte dans le TreeView à partir du noeud sélectionné.
                    oTreeNodeContrainte = AjouterContrainteNoeud(oTreeNodeTable, pRow, iPosAttTrier, iPosAttDecrire, m_AttributRequetes)

                    'Vérifier si le noeud a été créé
                    If oTreeNodeContrainte IsNot Nothing Then
                        'Ouvrir le noeud des contraintes
                        oTreeNodeContrainte.ExpandAll()
                        'Sélectionner le noeud de la contrainte
                        treContraintes.SelectedNode = oTreeNodeContrainte
                    End If
                Else
                    'Message d'erreur
                    MsgBox("ERREUR: Vous n'avez pas sélectionné la bonne table des contraintes!")
                End If
            Else
                'Message d'erreur
                MsgBox(oRequete.Message)
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oTreeNodeTable = Nothing
            oTreeNodeContrainte = Nothing
            oRequete = Nothing
            pRow = Nothing
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnAjouterRequete_Click(sender As Object, e As EventArgs) Handles btnAjouterRequete.Click
        'Déclarer les variables de travail
        Dim oTreeNodeTable As TreeNode = Nothing        'Objet contenant le Node du TreeView pour la table des contraintes.
        Dim oTreeNodeContrainte As TreeNode = Nothing   'Objet contenant le Node du TreeView pour la contrainte.
        Dim oTreeNodeRequetes As TreeNode = Nothing     'Objet contenant le Node du TreeView pour toutes les requêtes spatiales de la contrainte.
        Dim oTreeNodeRequete As TreeNode = Nothing      'Objet contenant le Node du TreeView pour une requête spatiale.
        Dim oTreeNode As TreeNode = Nothing             'Objet contenant le Node d'un Node.
        Dim oRequete As intRequete          'Objet utilisé pour effectuer la sélection selon la requête spécifiée.
        Dim pRow As IRow = Nothing                      'Interface contenant la contrainte d'intégrité de la table.
        Dim iPosAtt As Integer = -1     'Contient la position de l'atttribut des requêtes.

        Try
            'Définir la requête spécifiée
            oRequete = DefinirRequete()

            'Vérifier si la contrainte est valide
            If oRequete.EstValide Then
                'Vérifier si le noeud sélectionné est une requête
                If treContraintes.SelectedNode.Tag.ToString = "[REQUETE]" Then
                    'Définir le noeud de la requête
                    oTreeNodeRequete = treContraintes.SelectedNode
                    'Définir le noeud contenant toutes les requêtes
                    oTreeNodeRequetes = oTreeNodeRequete.Parent
                    'Sinon ce sont toutes les requêtes
                Else
                    'Définir le noeud de la requête
                    oTreeNodeRequetes = treContraintes.SelectedNode
                End If

                'Définir le noeud contenant la contrainte
                oTreeNodeContrainte = oTreeNodeRequetes.Parent

                'Définir le noeud contenant la table
                oTreeNodeTable = oTreeNodeContrainte.Parent.Parent

                'Vérifier si la table correspond à celle active
                If oTreeNodeTable.Text = cboTableContraintes.Text Then
                    'Extraire le noeud contenant l'attribut du OID
                    oTreeNode = oTreeNodeContrainte.Nodes.Item(m_TableContraintes.Table.OIDFieldName & "=")
                    'Extraire le noeud contenant la valeur du OID
                    oTreeNode = oTreeNode.Nodes.Item(0)

                    'Extraire la contrainte de la table
                    pRow = m_TableContraintes.Table.GetRow(CInt(oTreeNode.Text))
                    'Définir la position de l'attribut des requêtes
                    iPosAtt = m_TableContraintes.Table.FindField(m_AttributRequetes)

                    'Vérifier si la requête est absente de la contrainte
                    If Not pRow.Value(iPosAtt).ToString.Contains(oRequete.Commande) Then
                        'Vérifier si aucun noeud de requête n'a été sélectionné
                        If oTreeNodeRequete Is Nothing Then
                            'Ajouter la requête dans les requêtes à la bonne position
                            pRow.Value(iPosAtt) = pRow.Value(iPosAtt).ToString & vbCrLf & oRequete.Commande
                        Else
                            'Ajouter la requête dans les requêtes à la bonne position
                            pRow.Value(iPosAtt) = pRow.Value(iPosAtt).ToString.Replace(oTreeNodeRequete.Text, oRequete.Commande & vbCrLf & oTreeNodeRequete.Text)
                        End If
                        'Sauver la modification
                        pRow.Store()

                        'Détruire tous les noeuds des requêtes
                        oTreeNodeRequetes.Nodes.Clear()
                        'Ajouter toutes les requêtes à partir du noeud des requêtes vides.
                        oTreeNodeRequete = AjouterRequetesNoeud(oTreeNodeRequetes, m_AttributRequetes, pRow.Value(iPosAtt).ToString, oRequete.Commande)

                        'Vérifier si le noeud a été créer
                        If oTreeNodeRequete IsNot Nothing Then
                            'Ouvrir le noeud des requêtes
                            oTreeNodeRequete.ExpandAll()
                            'Sélectionner le noeud de la requête
                            treContraintes.SelectedNode = oTreeNodeRequete
                        End If
                    Else
                        'Message d'erreur
                        MsgBox("ERREUR: Vous n'avez pas sélectionné la bonne table des contraintes!")
                    End If
                Else
                    'Message d'erreur
                    MsgBox("ERREUR: Requête déjà présente : " & vbCrLf & oRequete.Commande)
                End If

            Else
                'Message d'erreur
                MsgBox(oRequete.Message)
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oTreeNodeTable = Nothing
            oTreeNodeContrainte = Nothing
            oTreeNodeRequetes = Nothing
            oTreeNodeRequete = Nothing
            oTreeNode = Nothing
            pRow = Nothing
            oRequete = Nothing
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnDeplacerRequeteBas_Click(sender As Object, e As EventArgs) Handles btnDeplacerRequeteBas.Click
        'Déclarer les variables de travail
        Dim oTreeNodeTable As TreeNode = Nothing        'Objet contenant le Node du TreeView pour la table des contraintes.
        Dim oTreeNodeContrainte As TreeNode = Nothing   'Objet contenant le Node du TreeView pour la contrainte.
        Dim oTreeNodeRequetes As TreeNode = Nothing     'Objet contenant le Node du TreeView pour toutes les requêtes spatiales de la contrainte.
        Dim oTreeNodeRequete As TreeNode = Nothing      'Objet contenant le Node du TreeView pour une requête spatiale.
        Dim oTreeNode As TreeNode = Nothing             'Objet contenant le Node d'un Node.
        Dim pRow As IRow = Nothing                      'Interface contenant la contrainte d'intégrité de la table.
        Dim sRequetes As String = ""    'Contient les requêtes de la contrainte.
        Dim iPosAtt As Integer = -1     'Contient la position de l'atttribut des requêtes.
        Dim iPosReq As Integer = -1     'Contient la position de la requête à déplacer.

        Try
            'Définir le noeud de la requête
            oTreeNodeRequete = treContraintes.SelectedNode
            'Définir le noeud contenant toutes les requêtes
            oTreeNodeRequetes = oTreeNodeRequete.Parent
            'Définir le noeud contenant la contrainte
            oTreeNodeContrainte = oTreeNodeRequetes.Parent
            'Définir le noeud contenant la table
            oTreeNodeTable = oTreeNodeContrainte.Parent.Parent

            'Vérifier si la table correspond à celle active
            If oTreeNodeTable.Text = cboTableContraintes.Text Then
                'Extraire le noeud contenant l'attribut du OID
                oTreeNode = oTreeNodeContrainte.Nodes.Item(m_TableContraintes.Table.OIDFieldName & "=")
                'Extraire le noeud contenant la valeur du OID
                oTreeNode = oTreeNode.Nodes.Item(0)

                'Extraire la contrainte de la table
                pRow = m_TableContraintes.Table.GetRow(CInt(oTreeNode.Text))
                'Définir la position de l'attribut des requêtes
                iPosAtt = m_TableContraintes.Table.FindField(m_AttributRequetes)

                'Extraire la position de la requête
                iPosReq = oTreeNodeRequetes.Nodes.IndexOf(oTreeNodeRequete)

                'Vérifier si on peut déplacer vers le bas
                If iPosReq < oTreeNodeRequetes.Nodes.Count - 1 Then
                    'Enlever la requête
                    oTreeNodeRequetes.Nodes.Remove(oTreeNodeRequete)
                    'Insérer la requête
                    oTreeNodeRequetes.Nodes.Insert(iPosReq + 1, oTreeNodeRequete)
                    'Traiter toutes les requêtes
                    For i = 0 To oTreeNodeRequetes.Nodes.Count - 1
                        'Si c'est la première requête
                        If i = 0 Then
                            'Définir la première requête
                            sRequetes = oTreeNodeRequetes.Nodes(i).Text
                        Else
                            'Ajouter les autres requêtes
                            sRequetes = sRequetes & vbCrLf & oTreeNodeRequetes.Nodes(i).Text
                        End If
                    Next
                    'Redéfinir les requêtes
                    pRow.Value(iPosAtt) = sRequetes
                    'Sauver la modification
                    pRow.Store()

                    'Vérifier si le noeud a été créer
                    If oTreeNodeRequete IsNot Nothing Then
                        'Ouvrir le noeud des requêtes
                        oTreeNodeRequete.ExpandAll()
                        'Sélectionner le noeud de la requête
                        treContraintes.SelectedNode = oTreeNodeRequete
                    End If
                End If
            Else
                'Message d'erreur
                MsgBox("ERREUR: Vous n'avez pas sélectionné la bonne table des contraintes!")
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oTreeNodeTable = Nothing
            oTreeNodeContrainte = Nothing
            oTreeNodeRequetes = Nothing
            oTreeNodeRequete = Nothing
            oTreeNode = Nothing
            pRow = Nothing
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnDeplacerRequeteHaut_Click(sender As Object, e As EventArgs) Handles btnDeplacerRequeteHaut.Click
        'Déclarer les variables de travail
        Dim oTreeNodeTable As TreeNode = Nothing        'Objet contenant le Node du TreeView pour la table des contraintes.
        Dim oTreeNodeContrainte As TreeNode = Nothing   'Objet contenant le Node du TreeView pour la contrainte.
        Dim oTreeNodeRequetes As TreeNode = Nothing     'Objet contenant le Node du TreeView pour toutes les requêtes spatiales de la contrainte.
        Dim oTreeNodeRequete As TreeNode = Nothing      'Objet contenant le Node du TreeView pour une requête spatiale.
        Dim oTreeNode As TreeNode = Nothing             'Objet contenant le Node d'un Node.
        Dim pRow As IRow = Nothing                      'Interface contenant la contrainte d'intégrité de la table.
        Dim sRequetes As String = ""    'Contient les requêtes de la contrainte.
        Dim iPosAtt As Integer = -1     'Contient la position de l'atttribut des requêtes.
        Dim iPosReq As Integer = -1     'Contient la position de la requête à déplacer.

        Try
            'Définir le noeud de la requête
            oTreeNodeRequete = treContraintes.SelectedNode
            'Définir le noeud contenant toutes les requêtes
            oTreeNodeRequetes = oTreeNodeRequete.Parent
            'Définir le noeud contenant la contrainte
            oTreeNodeContrainte = oTreeNodeRequetes.Parent
            'Définir le noeud contenant la table
            oTreeNodeTable = oTreeNodeContrainte.Parent.Parent

            'Vérifier si la table correspond à celle active
            If oTreeNodeTable.Text = cboTableContraintes.Text Then
                'Extraire le noeud contenant l'attribut du OID
                oTreeNode = oTreeNodeContrainte.Nodes.Item(m_TableContraintes.Table.OIDFieldName & "=")
                'Extraire le noeud contenant la valeur du OID
                oTreeNode = oTreeNode.Nodes.Item(0)

                'Extraire la contrainte de la table
                pRow = m_TableContraintes.Table.GetRow(CInt(oTreeNode.Text))
                'Définir la position de l'attribut des requêtes
                iPosAtt = m_TableContraintes.Table.FindField(m_AttributRequetes)

                'Extraire la position de la requête
                iPosReq = oTreeNodeRequetes.Nodes.IndexOf(oTreeNodeRequete)

                'Vérifier si on peut déplacer vers le bas
                If iPosReq > 0 Then
                    'Enlever la requête
                    oTreeNodeRequetes.Nodes.Remove(oTreeNodeRequete)
                    'Insérer la requête
                    oTreeNodeRequetes.Nodes.Insert(iPosReq - 1, oTreeNodeRequete)
                    'Traiter toutes les requêtes
                    For i = 0 To oTreeNodeRequetes.Nodes.Count - 1
                        'Si c'est la première requête
                        If i = 0 Then
                            'Définir la première requête
                            sRequetes = oTreeNodeRequetes.Nodes(i).Text
                        Else
                            'Ajouter les autres requêtes
                            sRequetes = sRequetes & vbCrLf & oTreeNodeRequetes.Nodes(i).Text
                        End If
                    Next
                    'Redéfinir les requêtes
                    pRow.Value(iPosAtt) = sRequetes
                    'Sauver la modification
                    pRow.Store()

                    'Vérifier si le noeud a été créer
                    If oTreeNodeRequete IsNot Nothing Then
                        'Ouvrir le noeud des requêtes
                        oTreeNodeRequete.ExpandAll()
                        'Sélectionner le noeud de la requête
                        treContraintes.SelectedNode = oTreeNodeRequete
                    End If
                End If
            Else
                'Message d'erreur
                MsgBox("ERREUR: Vous n'avez pas sélectionné la bonne table des contraintes!")
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oTreeNodeTable = Nothing
            oTreeNodeContrainte = Nothing
            oTreeNodeRequetes = Nothing
            oTreeNodeRequete = Nothing
            oTreeNode = Nothing
            pRow = Nothing
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnEnlever_Click(sender As Object, e As EventArgs) Handles btnEnlever.Click
        Try
            'Vérifier si le noeud est valide
            If treContraintes.SelectedNode IsNot Nothing Then
                'Vérifier si le noeud est une contrainte
                If treContraintes.SelectedNode.Tag.ToString = "CONTRAINTE" Then
                    'Vérifier s'il n'y a seulement qu'une contrainte
                    If treContraintes.SelectedNode.Parent.Nodes.Count = 1 Then
                        'Enlever le noeud de groupe car il n'y a plus de contrainte dans le groupe
                        treContraintes.SelectedNode.Parent.Remove()
                    Else
                        'Enlever le noeud de la contrainte
                        treContraintes.SelectedNode.Remove()
                    End If
                Else
                    'Enlever le noeud
                    treContraintes.SelectedNode.Remove()
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnDetruire_Click(sender As Object, e As EventArgs) Handles btnDetruire.Click
        'Déclarer les variables de travail
        Dim oTreeNodeTable As TreeNode = Nothing        'Objet contenant le Node du TreeView pour la table des contraintes.
        Dim oTreeNodeGroupe As TreeNode = Nothing       'Objet contenant le Node du TreeView pour le regroupement.
        Dim oTreeNodeContrainte As TreeNode = Nothing   'Objet contenant le Node du TreeView pour la contrainte.
        Dim oTreeNodeRequetes As TreeNode = Nothing     'Objet contenant le Node du TreeView pour toutes les requêtes spatiales de la contrainte.
        Dim oTreeNodeRequete As TreeNode = Nothing      'Objet contenant le Node du TreeView pour une requête spatiale.
        Dim oTreeNode As TreeNode = Nothing             'Objet contenant le Node d'un Node.
        Dim pRow As IRow = Nothing                      'Interface contenant la contrainte d'intégrité de la table.
        Dim iPosAtt As Integer = -1     'Contient la position de l'atttribut des requêtes.

        Try
            'Vérifier si le noeud sélectionné est une requête
            If treContraintes.SelectedNode.Tag.ToString = "[REQUETE]" Then
                'Définir le noeud de la requête
                oTreeNodeRequete = treContraintes.SelectedNode
                'Définir le noeud contenant toutes les requêtes
                oTreeNodeRequetes = oTreeNodeRequete.Parent
                'Définir le noeud contenant la contrainte
                oTreeNodeContrainte = oTreeNodeRequetes.Parent
                'Sinon on traite la requête
            Else
                'Définir le noeud de la requête
                oTreeNodeContrainte = treContraintes.SelectedNode
            End If

            'Définir le noeud contenant le groupe
            oTreeNodeGroupe = oTreeNodeContrainte.Parent

            'Définir le noeud contenant la table
            oTreeNodeTable = oTreeNodeGroupe.Parent

            'Vérifier si la table correspond à celle active
            If oTreeNodeTable.Text = cboTableContraintes.Text Then
                'Confirmer si on veut détruire
                If MsgBox("ATTENTION : L'information sera détruite directement dans la table." & vbCrLf & "Êtes-vous certains de vouloir détruire ? ", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    'Extraire le noeud contenant l'attribut du OID
                    oTreeNode = oTreeNodeContrainte.Nodes.Item(m_TableContraintes.Table.OIDFieldName & "=")
                    'Extraire le noeud contenant la valeur du OID
                    oTreeNode = oTreeNode.Nodes.Item(0)

                    'Extraire la contrainte de la table
                    pRow = m_TableContraintes.Table.GetRow(CInt(oTreeNode.Text))
                    'Définir la position de l'attribut des requêtes
                    iPosAtt = m_TableContraintes.Table.FindField(m_AttributRequetes)

                    'Vérifier si aucun noeud de requête n'a été sélectionné, on détruit la contrainte
                    If oTreeNodeRequete Is Nothing Then
                        'Détruire la contrainte de la table
                        pRow.Delete()
                        'Détruire le noeud de la contrainte du TreeView
                        oTreeNodeGroupe.Nodes.Remove(oTreeNodeContrainte)
                        'Vérifier si on doit enlever aussi détruire le noeud de regroupement
                        If oTreeNodeGroupe.Nodes.Count = 0 Then
                            'Détruire le noeud de groupe du TreeView
                            oTreeNodeTable.Nodes.Remove(oTreeNodeGroupe)
                        End If

                        'Détruire une requête de la contrainte
                    Else
                        'Détruire la requête dans les requêtes de la table
                        pRow.Value(iPosAtt) = pRow.Value(iPosAtt).ToString.Replace(oTreeNodeRequete.Text & vbCrLf, "")
                        pRow.Value(iPosAtt) = pRow.Value(iPosAtt).ToString.Replace(vbCrLf & oTreeNodeRequete.Text, "")
                        'Sauver la modification
                        pRow.Store()
                        'Détruire la requête du TreeView
                        oTreeNodeContrainte.Nodes.Remove(oTreeNodeRequete)
                    End If
                End If

            Else
                'Message d'erreur
                MsgBox("ERREUR: Vous n'avez pas sélectionné la bonne table des contraintes!")
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oTreeNodeTable = Nothing
            oTreeNodeContrainte = Nothing
            oTreeNodeRequetes = Nothing
            oTreeNodeRequete = Nothing
            oTreeNode = Nothing
            pRow = Nothing
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnModifier_Click(sender As Object, e As EventArgs) Handles btnModifier.Click
        Try
            'Modifier le noeud sélectionné.
            Call ModifierNoeud(treContraintes.SelectedNode)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnExecuter_Click(sender As Object, e As EventArgs) Handles btnExecuter.Click
        'Déclarer les variables de tavail
        Dim oTableContraintes As clsTableContraintes = Nothing  'Objet qui permet d'exécuter les contraintes
        Dim dDate As System.DateTime = System.DateTime.Now  'Contient la date de début d'exécution
        Dim vExecuter As MsgBoxResult = MsgBoxResult.Yes    'Par défaut, on exécute le traitement

        Try
            'Initialiser le message d'erreur
            rtbMessages.Text = "Initialisation du traitement ..." & vbCrLf
            rtbMessages.Refresh()

            'Initialiser le RichTextBox
            m_RichTextBox = rtbMessages

            'Permettre d'annuler la sélection avec la touche ESC
            m_TrackCancel = New CancelTracker
            m_TrackCancel.CancelOnKeyPress = True
            m_TrackCancel.CancelOnClick = False

            'Définir l'objet qui permet d'exécuter les contraintes
            oTableContraintes = New clsTableContraintes(cboGeodatabaseClasses.Text, treContraintes.SelectedNode, m_MxDocument.FocusMap, m_CreerClasseErreur)
            oTableContraintes.NomClasseDecoupage = cboClasseDecoupage.Text
            oTableContraintes.NomAttributDecoupage = cboAttributDecoupage.Text
            oTableContraintes.NomFichierJournal = cboFichierJournal.Text
            oTableContraintes.NomRepertoireErreurs = cboRepertoireErreurs.Text
            oTableContraintes.NomRapportErreurs = cboRapportErreurs.Text
            oTableContraintes.Courriel = cboCourriel.Text
            oTableContraintes.SpatialReference = m_MxDocument.FocusMap.SpatialReference
            oTableContraintes.TrackCancel = m_TrackCancel
            oTableContraintes.RichTextbox = m_RichTextBox

            'Vérifier si une requête de découpage est présente
            If Not cboClasseDecoupage.Text.Contains(":") Then
                'Démander si on veut exécuter le traitement
                vExecuter = MsgBox("Voulez-vous exécuter le traitement quand même ?", MsgBoxStyle.YesNo, "ATTENTION : Aucun découpage n'est présent ce qui peut être long à exécuter !")
            End If

            'Exécuter les contraintes
            If vExecuter = MsgBoxResult.Yes Then
                'Initialiser le message d'erreur
                rtbMessages.Text = "Traitement exécuté en interactif ..." & vbCrLf
                rtbMessages.Refresh()
                'Exécuter le traitement
                oTableContraintes.Executer()
            Else
                'Initialiser le message d'erreur
                rtbMessages.Text = "Traitement annulé !" & vbCrLf
                rtbMessages.Refresh()
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oTableContraintes = Nothing
            dDate = Nothing
        End Try
    End Sub

    'Exécuté lorsqu'un click est effectuer sur le bouton
    Private Sub btnExecuterBatch_Click(sender As Object, e As EventArgs) Handles btnExecuterBatch.Click
        'Déclarer les variables de tavail
        Dim oTableContraintes As clsTableContraintes = Nothing  'Objet qui permet d'exécuter les contraintes
        Dim oProcess As Process = Nothing                   'Contient un processus pour exécuter en batch.
        Dim dDate As System.DateTime = System.DateTime.Now  'Contient la date de début d'exécution
        Dim sProgramme As String = ""       'Contient le nom complet du programme à exécuter en batch.
        Dim sArguments As String = ""       'Contient les arguments du programme à exécuter en batch.

        Try
            'Initialiser le message
            rtbMessages.Clear()
            rtbMessages.Refresh()

            'Définir l'objet qui permet d'exécuter les contraintes
            oTableContraintes = New clsTableContraintes(cboGeodatabaseClasses.Text, treContraintes.SelectedNode, m_MxDocument.FocusMap, False)
            oTableContraintes.NomClasseDecoupage = cboClasseDecoupage.Text
            oTableContraintes.NomAttributDecoupage = cboAttributDecoupage.Text
            oTableContraintes.NomFichierJournal = cboFichierJournal.Text
            oTableContraintes.NomRepertoireErreurs = cboRepertoireErreurs.Text
            oTableContraintes.NomRapportErreurs = cboRapportErreurs.Text
            oTableContraintes.Courriel = cboCourriel.Text

            'Afficher la commande pour exécuter en batch
            sProgramme = "D:\cits\EnvCits\applications\gestion_bdg\pro\Geotraitement\exe\ValiderContrainte.exe"
            sArguments = oTableContraintes.Commande.Replace("ValiderContrainte.exe", "")
            rtbMessages.AppendText(sProgramme & " " & sArguments & vbCrLf & vbCrLf)
            rtbMessages.Refresh()

            'Exécuter le programme en batch
            oProcess = Process.Start(sProgramme, sArguments)

            'Afficher le message d'exécution en arrière plan
            rtbMessages.AppendText("Traitement exécuté en arrière plan (Batch) !")
            rtbMessages.Refresh()

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oTableContraintes = Nothing
            oProcess = Nothing
            sProgramme = Nothing
            sArguments = Nothing
            dDate = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions d'événements"
    'Exécuté lorsque le focus est mis sur le comboBox 
    Private Sub cboGeodatabaseClasses_GotFocus(sender As Object, e As EventArgs) Handles cboGeodatabaseClasses.GotFocus
        'Déclarer les variables de travail
        Dim iSelectedIndex As Integer = -1      'Contient le numéro de l'index à sélectionner.
        Dim sNom As String = ""                 'Contient le nom de la Géodatabase au départ

        Try
            'Définir le nom de la Géodatabase au départ
            sNom = cboGeodatabaseClasses.Text
            'Remplir le ComboBox pour définir la Geodatabase
            iSelectedIndex = RemplirComboBoxGeodatabase(cboGeodatabaseClasses, cboGeodatabaseClasses.Text)
            'Sélectionner l'index
            If sNom <> cboGeodatabaseClasses.Text Then cboGeodatabaseClasses.SelectedIndex = iSelectedIndex

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsque le nom de la Géodatabase change afin de définir l'objet y correspondant.
    Private Sub cboGeodatabaseClasses_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGeodatabaseClasses.SelectedIndexChanged
        'Déclarer la variables de travail
        Dim qGeodatabaseColl As Collection = Nothing        'Contient la liste des Géodatabases.
        Dim pWorkspace As IWorkspace = Nothing              'Interface contenant une Géodatabase.

        Try
            'Vérifier si le nom de la Géodatabase est un mot clé
            If cboGeodatabaseClasses.Text = "BDRS_PRO_BDG_DBA" Or cboGeodatabaseClasses.Text = "BDRS_TST_BDG_DBA" Then
                'Définir la géodatabase
                m_Geodatabase = DefinirGeodatabase(cboGeodatabaseClasses.Text)

                'Si le nom de la Géodatabase n'est pas un mot clé
            Else
                'Créer une nouvelle collection de géodatabase vide
                qGeodatabaseColl = MapGeodatabaseColl(m_MxDocument.FocusMap)

                'Traiter toutes les Géodatabases trouvées
                For i = 1 To qGeodatabaseColl.Count
                    'Définir la Géodatabase
                    pWorkspace = CType(qGeodatabaseColl.Item(i), IWorkspace)

                    'Vérifier la présence du texte recherché pour la valeur par défaut
                    If pWorkspace.PathName = cboGeodatabaseClasses.Text Then
                        'Définir la géodatabase
                        m_Geodatabase = pWorkspace

                        'Sortir
                        Exit For
                    End If
                Next
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            qGeodatabaseColl = Nothing
            pWorkspace = Nothing
        End Try
    End Sub

    'Exécuté lorsque le focus est mis sur le comboBox 
    Private Sub cboTableContraintes_GotFocus(sender As Object, e As EventArgs) Handles cboTableContraintes.GotFocus
        'Déclarer les variables de travail
        Dim iSelectedIndex As Integer = -1      'Contient le numéro de l'index à sélectionner.
        Dim sNom As String = ""                 'Contient le nom de la table au départ

        Try
            'Définir le nom de la table au départ
            sNom = cboTableContraintes.Text
            'Remplir le ComboBox pour définir la table d'intégrité spatiale
            iSelectedIndex = RemplirComboBoxTable(cboTableContraintes, cboTableContraintes.Text)
            'Sélectionner l'index
            If sNom <> cboTableContraintes.Text Then cboTableContraintes.SelectedIndex = iSelectedIndex

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    'Exécuté lorsque le nom de la table des containtes change afin de définir l'objet y correspondant 
    'et afficher les contraintes qu'elles contiennent dans le TreeView.
    Private Sub cboTableContraintes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTableContraintes.SelectedIndexChanged
        'Déclarer la variables de travail
        Dim pStandaloneTableColl As IStandaloneTableCollection = Nothing    'Interface qui permet d'extraire les StandaloneTable de la Map.
        Dim pStandaloneTable As IStandaloneTable = Nothing                  'Interface contenant une classe de données.
        Dim pDataset As IDataset = Nothing                                  'Interface contenant le nom complet de la table.
        Dim iSelectedIndex As Integer = -1      'Contient le numéro de l'index à sélectionner.
        Dim sNomTable As String = ""
        Dim sNomTableContraintes As String = ""

        Try
            'Définir le nom de la table des contraintes
            sNomTableContraintes = cboTableContraintes.Text

            'Définir la liste des StandaloneTable
            pStandaloneTableColl = CType(m_MxDocument.FocusMap, IStandaloneTableCollection)

            'Vérifier si la géodatabse correspond à un mot clé
            If cboGeodatabaseClasses.Text = "BDRS_PRO_BDG_DBA" Or cboGeodatabaseClasses.Text = "BDRS_TST_BDG_DBA" Then
                'Vérifier si la table correspond à celle par défaut
                If cboTableContraintes.Text = "CONTRAINTE_INTEGRITE_SPATIALE" Then
                    'Définir la table des contraintes d'intégrité
                    m_TableContraintes = DefinirStandaloneTable(CType(m_Geodatabase, IWorkspace2), cboTableContraintes.Text)

                    'Ajouter la table des contraintes dans la Map active
                    pStandaloneTableColl.AddStandaloneTable(m_TableContraintes)
                    'Redéfinir le nom de la table des contraintes
                    sNomTableContraintes = m_TableContraintes.Name

                    'Activer les boutons
                    btnImporter.Enabled = True
                    cboTrierPar.Enabled = True
                    cboDecrirePar.Enabled = True
                    btnAjouterContrainte.Enabled = True

                    'Sortir
                    'Exit Sub
                End If
            End If

            'Traiter tous les StandaloneTable
            For i = 0 To pStandaloneTableColl.StandaloneTableCount - 1
                'Définir la table
                pStandaloneTable = pStandaloneTableColl.StandaloneTable(i)

                'Vérifier si la table est valide
                If pStandaloneTable.Table IsNot Nothing Then
                    'Vérifier si la table contient des OID
                    If pStandaloneTable.Table.HasOID Then
                        'Interface pour extraire le nom complet de la table
                        pDataset = CType(pStandaloneTable.Table, IDataset)
                        'Vérifier la présence du Path de la Géodatabase
                        If pDataset.Workspace.PathName.Length = 0 Then
                            'Définir le nom complet de la table
                            sNomTable = pDataset.Name
                        Else
                            'Définir le nom complet de la table
                            sNomTable = pDataset.Workspace.PathName & "\" & pDataset.Name
                        End If
                        'Vérifier si la table correspond à celle sélectionnée
                        If sNomTable = sNomTableContraintes Then
                            'Définir la tables des contraintes d'intégrités spatiales
                            m_TableContraintes = pStandaloneTable

                            'Remplir le comboBox de l'attribut pour trier
                            iSelectedIndex = RemplirComboBoxAttribut(cboTrierPar, m_AttributTrier)
                            'Sélectionner l'index
                            cboTrierPar.SelectedIndex = iSelectedIndex

                            'Remplir le comboBox de l'attribut pour décrire
                            iSelectedIndex = RemplirComboBoxAttribut(cboDecrirePar, m_AttributDecrire)
                            'Sélectionner l'index
                            cboDecrirePar.SelectedIndex = iSelectedIndex

                            'Activer les boutons
                            btnImporter.Enabled = True
                            cboTrierPar.Enabled = True
                            cboDecrirePar.Enabled = True
                            btnAjouterContrainte.Enabled = True

                            'Sortir de la boucle
                            Exit For
                        End If
                    End If
                End If
            Next

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pStandaloneTableColl = Nothing
            pStandaloneTable = Nothing
            pDataset = Nothing
        End Try
    End Sub

    Private Sub treContraintes_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles treContraintes.AfterSelect
        Try
            'Afficher l'information du noeud
            Call AfficherNoeud(e.Node)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub treContraintes_DoubleClick(sender As Object, e As EventArgs) Handles treContraintes.DoubleClick
        'Déclarer les variables de travail
        Try
            'Modifier le noeud sélectionné
            Call ModifierNoeud(treContraintes.SelectedNode)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
        End Try
    End Sub

    Private Sub treContraintes_GotFocus(sender As Object, e As EventArgs) Handles treContraintes.GotFocus
        Try
            'Afficher l'information du noeud
            Call AfficherNoeud(treContraintes.SelectedNode)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub treContraintes_LostFocus(sender As Object, e As EventArgs) Handles treContraintes.LostFocus
        Try
            'Désactiver les boutons
            btnAjouterRequete.Enabled = False
            btnDetruire.Enabled = False
            btnModifier.Enabled = False
            btnExecuter.Enabled = False
            btnExecuterBatch.Enabled = False

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub
#End Region

#Region "Routines et fonctions publiques"
    '''<summary>
    ''' Permet d'initialiser le menu des contraintes d'intégrité selon les valeurs par défaut.
    '''</summary>
    '''
    Public Sub Init()
        'Déclarer les variables de travail
        Dim iSelectedIndex As Integer = -1      'Contient le numéro de l'index à sélectionner.
        Dim sNomClasseDecoupage = "ges_Decoupage_SNRC50K_Canada_2"
        Dim sNomAttributDecoupage = "DATASET_NAME"
        Dim sNomRepertoireDefaut As String = "D:" 'Contient le nom du répertoire par défaut.
        Dim sNomRepertoireErreurs As String = sNomRepertoireDefaut & "\Erreurs_[DATE_TIME]_[DATASET_NAME].mdb"
        Dim sNomRapportErreurs As String = sNomRepertoireDefaut & "\Rapport_[DATE_TIME].txt"
        Dim sNomFichierJournal As String = sNomRepertoireDefaut & "\Journal_[DATE_TIME].log"

        Try
            'Vider toutes les tables des contraintes d'intégrité
            treContraintes.Nodes.Clear()
            'Désactiver les boutons
            btnImporter.Enabled = False
            cboTrierPar.Enabled = False
            cboDecrirePar.Enabled = False
            btnAjouterContrainte.Enabled = False
            btnAjouterRequete.Enabled = False
            btnDeplacerRequeteHaut.Enabled = False
            btnDeplacerRequeteBas.Enabled = False
            btnEnlever.Enabled = False
            btnDetruire.Enabled = False
            btnModifier.Enabled = False
            btnExecuter.Enabled = False
            btnExecuterBatch.Enabled = False

            'Initialiser le nom de la classe de découpage
            cboClasseDecoupage.Items.Clear()
            cboClasseDecoupage.Items.Add("")
            cboClasseDecoupage.Items.Add(sNomClasseDecoupage)
            cboClasseDecoupage.Items.Add("ges_Decoupage_SNRC50K_2")
            cboClasseDecoupage.Items.Add("ges_Decoupage_RHN_2")
            cboClasseDecoupage.Items.Add("ges_Decoupage_Province_2")
            cboClasseDecoupage.Items.Add("ges_Decoupage_Canada_2")
            cboClasseDecoupage.Items.Add("ges_Decoupage_SNRC50K_2")
            cboClasseDecoupage.Items.Add("D:\ges_Decoupage_SNRC50K_2.lyr")
            cboClasseDecoupage.Text = sNomClasseDecoupage

            'Initialiser le nom de la classe de découpage
            cboAttributDecoupage.Items.Clear()
            cboAttributDecoupage.Items.Add("")
            cboAttributDecoupage.Items.Add(sNomAttributDecoupage)
            cboAttributDecoupage.Text = sNomAttributDecoupage

            'Initialiser le nom du répertoire d'erreurs
            cboRepertoireErreurs.Items.Clear()
            cboRepertoireErreurs.Items.Add("")
            cboRepertoireErreurs.Items.Add(sNomRepertoireErreurs)
            cboRepertoireErreurs.Items.Add(sNomRepertoireDefaut & "\Erreurs_[DATE_TIME].mdb")
            cboRepertoireErreurs.Items.Add(sNomRepertoireDefaut & "\Erreurs")
            cboRepertoireErreurs.Text = sNomRepertoireErreurs

            'Initialiser le nom du rapport d'erreurs
            cboRapportErreurs.Items.Clear()
            cboRapportErreurs.Items.Add("")
            cboRapportErreurs.Items.Add(sNomRapportErreurs)
            cboRapportErreurs.Text = sNomRapportErreurs

            'Initialiser le nom du fichier journal
            cboFichierJournal.Items.Clear()
            cboFichierJournal.Items.Add("")
            cboFichierJournal.Items.Add(sNomFichierJournal)
            cboFichierJournal.Text = sNomFichierJournal

            'Initialiser les courriels
            cboCourriel.Items.Clear()
            cboFichierJournal.Items.Add("")

            'Afficher le nombre de contraintes
            lblInformation.Text = "0 Contrainte"

            'Définir les valeurs par défaut
            m_AttributTrier = "NOM_TABLE"   'Nom de l'attribut utilisé pour trier les contraintes.
            m_AttributDecrire = "GROUPE"    'Nom de l'attribut utilisé pour décrire les contraintes.
            m_AttributRequetes = "REQUETES" 'Nom de l'attribut utilisé pour exécuter les requêtes.

            'Initialiser la Géodatabase
            m_Geodatabase = Nothing
            'Remplir le ComboBox pour définir la Geodatabase
            iSelectedIndex = RemplirComboBoxGeodatabase(cboGeodatabaseClasses, "BDRS_PRO_BDG_DBA")
            'Sélectionner l'index
            cboGeodatabaseClasses.SelectedIndex = iSelectedIndex

            'Initialiser la table des contraintes
            m_TableContraintes = Nothing
            'Remplir le ComboBox pour définir la table d'intégrité spatiale
            iSelectedIndex = RemplirComboBoxTable(cboTableContraintes, "")
            'Sélectionner l'index
            cboTableContraintes.SelectedIndex = iSelectedIndex

            'Remplir le comboBox de l'attribut pour trier les contraintes
            iSelectedIndex = RemplirComboBoxAttribut(cboTrierPar, m_AttributTrier)
            'Sélectionner l'index
            cboTrierPar.SelectedIndex = iSelectedIndex

            'Remplir le comboBox de l'attribut pour décrire les contraintes
            iSelectedIndex = RemplirComboBoxAttribut(cboDecrirePar, m_AttributDecrire)
            'Sélectionner l'index
            cboDecrirePar.SelectedIndex = iSelectedIndex

            'Rafraîchir le menu
            Me.Refresh()

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            iSelectedIndex = Nothing
            sNomRepertoireDefaut = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions privées"
    '''<summary>
    ''' Fonction qui permet de créer une nouvelle table des contraintes d'intégrité dans une géodatabase existante.
    '''</summary>
    '''
    '''<param name="sNomGeodatabase">Nom de la géodatabase existante.</param>
    '''<param name="sNomTableContraintes">Nom de la nouvelle table des contraintes d'intégrité.</param>
    '''<param name="sNomProprietaireDefaut"> Contient le nom du propriétaire des tables par défaut pour les Géodatabase Enterprise.</param>
    ''' 
    '''<returns>"ITable" contenant une nouvelle table des contraintes d'intégrité.</returns>
    '''
    Private Function CreerNouvelleTableContraintes(ByVal sNomGeodatabase As String, ByVal sNomTableContraintes As String,
                                                   Optional ByVal sNomProprietaireDefaut As String = "BDG_DBA") As ITable
        'Déclarer les variables de travail
        Dim pWorkspace2 As IWorkspace2 = Nothing                'Interface pour vérifier si la table existe.
        Dim pWorkspace As IWorkspace = Nothing                  'Interface pour vérifier le type de Géodatabase.
        Dim pGeodatabaseType As Type = Nothing                  'Interface utilisé pour définir le type de géodatabase.
        Dim pWorkspaceFactory As IWorkspaceFactory = Nothing    'Interface pour créer un Workspace en mémoire.
        Dim pFeatureWorkspace As IFeatureWorkspace = Nothing    'Interface contenant un FeatureWorkspace.
        Dim pFields As IFields = Nothing                        'Interface pour contenant les attributs de la Featureclass.
        Dim pFieldsEdit As IFieldsEdit = Nothing                'Interface pour créer les attributs.
        Dim pFieldEdit As IFieldEdit = Nothing                  'Interface pour créer un attribut.
        Dim pUID As New UID                                     'Interface pour générer un UID.

        'Définir la valeur par défaut
        CreerNouvelleTableContraintes = Nothing

        Try
            'Vérifier si la Geodatabse est SDE
            If sNomGeodatabase.Contains(".sde") Then
                'Définir le type de workspace : SDE
                pGeodatabaseType = Type.GetTypeFromProgID("esriDataSourcesGDB.SdeWorkspaceFactory")

                'Si la Geodatabse est une File Geodatabase
            ElseIf sNomGeodatabase.Contains(".gdb") Then
                'Définir le type de workspace : GDB
                pGeodatabaseType = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory")

                'Si la Geodatabse est une personnelle Geodatabse
            ElseIf sNomGeodatabase.Contains(".mdb") Then
                'Définir le type de workspace : MDB
                pGeodatabaseType = Type.GetTypeFromProgID("esriDataSourcesGDB.AccessWorkspaceFactory")

                'Si le format est invalide
            Else
                'Retourner l'erreur
                Err.Raise(-1, , "Erreur : Le format spécifié de la Geodatabase est invalide !")
            End If

            'Interface pour ouvrir la Géodatabase
            pWorkspaceFactory = CType(Activator.CreateInstance(pGeodatabaseType), IWorkspaceFactory)

            'Ouvrir la Géodatabase
            pFeatureWorkspace = CType(pWorkspaceFactory.OpenFromFile(sNomGeodatabase, 0), IFeatureWorkspace)

            'Vérifier si la Géodatabase est présente
            If pFeatureWorkspace IsNot Nothing Then
                'Interface pour vérifier le type de Géodatabase.
                pWorkspace = CType(pFeatureWorkspace, IWorkspace)
                'Vérifier si la Géodatabase est de type "Enterprise" 
                If pWorkspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                    'Vérifier si le nom de la table contient le nom du propriétaire
                    If Not sNomTableContraintes.Contains(".") Then
                        'Définir le nom de la table avec le nom du propriétaire
                        sNomTableContraintes = sNomProprietaireDefaut & "." & sNomTableContraintes
                    End If
                End If

                'Interface pour vérifier si la table existe
                pWorkspace2 = CType(pFeatureWorkspace, IWorkspace2)
                'Vérifier si la table existe
                If pWorkspace2.NameExists(esriDatasetType.esriDTTable, sNomTableContraintes) Then
                    'Ouvrir la table des contraintes
                    CreerNouvelleTableContraintes = pFeatureWorkspace.OpenTable(sNomTableContraintes)
                End If

                'Vérifier si la table est absente, on doit la créer
                If CreerNouvelleTableContraintes Is Nothing Then
                    'Définir le type d'élément
                    pUID.Value = "esriGeodatabase.Row"

                    'Interface pour créer des attributs
                    pFieldsEdit = New Fields

                    'Définir le nombre d'attributs
                    pFieldsEdit.FieldCount_2 = 10

                    'Créer l'attribut du OBJECTID
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "OBJECTID"
                        .AliasName_2 = "OBJECTID"
                        .Type_2 = esriFieldType.esriFieldTypeOID
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(0) = pFieldEdit

                    'Créer l'attribut ETAMPE
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "ETAMPE"
                        .AliasName_2 = "ETAMPE"
                        .Type_2 = esriFieldType.esriFieldTypeString
                        .Length_2 = 23
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(1) = pFieldEdit

                    'Créer l'attribut DT_C
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "DT_C"
                        .AliasName_2 = "DT_C"
                        .Type_2 = esriFieldType.esriFieldTypeDate
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(2) = pFieldEdit

                    'Créer l'attribut DT_M
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "DT_M"
                        .AliasName_2 = "DT_M"
                        .Type_2 = esriFieldType.esriFieldTypeDate
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(3) = pFieldEdit

                    'Créer l'attribut GROUPE
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "GROUPE"
                        .AliasName_2 = "GROUPE"
                        .Type_2 = esriFieldType.esriFieldTypeString
                        .Length_2 = 50
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(4) = pFieldEdit

                    'Créer l'attribut DESCRIPTION
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "DESCRIPTION"
                        .AliasName_2 = "DESCRIPTION"
                        .Type_2 = esriFieldType.esriFieldTypeString
                        .Length_2 = 2000
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(5) = pFieldEdit

                    'Créer l'attribut MESSAGE
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "MESSAGE"
                        .AliasName_2 = "MESSAGE"
                        .Type_2 = esriFieldType.esriFieldTypeString
                        .Length_2 = 2000
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(6) = pFieldEdit

                    'Créer l'attribut REQUETES
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "REQUETES"
                        .AliasName_2 = "REQUETES"
                        .Type_2 = esriFieldType.esriFieldTypeString
                        .Length_2 = 2000
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(7) = pFieldEdit

                    'Créer l'attribut NOM_REQUETE
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "NOM_REQUETE"
                        .AliasName_2 = "NOM_REQUETE"
                        .Type_2 = esriFieldType.esriFieldTypeString
                        .Length_2 = 50
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(8) = pFieldEdit

                    'Créer l'attribut NOM_TABLE
                    pFieldEdit = New Field
                    With pFieldEdit
                        .Name_2 = "NOM_TABLE"
                        .AliasName_2 = "NOM_TABLE"
                        .Type_2 = esriFieldType.esriFieldTypeString
                        .Length_2 = 50
                        .IsNullable_2 = False
                    End With
                    'Ajouter l'attribut
                    pFieldsEdit.Field_2(9) = pFieldEdit

                    'Créer la Featureclass
                    CreerNouvelleTableContraintes = pFeatureWorkspace.CreateTable(sNomTableContraintes, pFieldsEdit, pUID, Nothing, "")
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pWorkspace2 = Nothing
            pWorkspace = Nothing
            pGeodatabaseType = Nothing
            pWorkspaceFactory = Nothing
            pFeatureWorkspace = Nothing
            pFields = Nothing
            pFieldsEdit = Nothing
            pFieldEdit = Nothing
            pUID = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'écrire une nouvelle contrainte d'intégrité et de la retourner.
    '''</summary>
    ''' 
    '''<param name="pTable"> Interface contenant la table des contraintes d'intégrité.</param>
    '''<param name="sGroupe"> Groupe de la contrainte d'intégrité.</param>
    '''<param name="sDescription"> Description de la contrainte d'intégrité.</param>
    '''<param name="sMessage"> Message de la contrainte d'intégrité.</param>
    '''<param name="sRequetes"> Requetes de la contrainte d'intégrité.</param>
    '''<param name="sNomRequete"> Nom de la requête de la contrainte d'intégrité.</param>
    '''<param name="sNomTable"> Nom de la table de la contrainte d'intégrité.</param>
    ''' 
    ''' <returns>IRow contenant l'information de la contrainte d'intégrité.</returns>
    '''
    Private Function AjouterContrainte(ByRef pTable As ITable, _
                                       ByVal sGroupe As String, ByVal sDescription As String, ByVal sMessage As String, _
                                       ByVal sRequetes As String, ByVal sNomRequete As String, ByVal sNomTable As String) As IRow
        'Déclarer les variables de travail
        Dim pRow As IRow = Nothing      'Interface ESRI contenant la contrainte créée.
        Dim dDate As DateTime = System.DateTime.Now 'Contient la date et l'heure courante.

        'Définir la contrainte par défaut
        AjouterContrainte = Nothing

        Try
            'Sortir si la classe d'erreurs est absente
            If pTable Is Nothing Then Exit Function

            'Créer une nouvelle contrainte vide
            pRow = pTable.CreateRow

            'Définir l'usager dans ETAMPE
            pRow.Value(1) = System.Environment.GetEnvironmentVariable("USERNAME")

            'Définir la date de création DT_C
            pRow.Value(2) = dDate

            'Définir la date de modification DT_M
            pRow.Value(3) = dDate

            'Définir le groupe
            pRow.Value(4) = sGroupe

            'Définir la description
            pRow.Value(5) = sDescription

            'Définir le message
            pRow.Value(6) = sMessage

            'Définir les requêtes
            pRow.Value(7) = sRequetes

            'Définir le nom de la requête
            pRow.Value(8) = sNomRequete

            'Définir le nom de la table
            pRow.Value(9) = sNomTable

            'Sauver l'ajout
            pRow.Store()

            'Retourner la contrainte
            AjouterContrainte = pRow

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pRow = Nothing
            dDate = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'afficher l'information selon son type de noeud.
    '''</summary>
    '''
    '''<param name="oTreeNode">Noeud d'un TreeView pour lequel on affiche l'information nécessaire.</param>
    ''' 
    ''' <remarks>
    ''' Certaines informations sont présentes dans le ToolTip du noeud.
    '''</remarks>
    Private Sub AfficherNoeud(ByVal oTreeNode As TreeNode)
        'Déclarer les variables de travail
        Dim iNbContraintes As Integer = 0   'Contient le nombre de contraintes
        Dim iPosAtt As Integer = -1         'contient la position de l'attribut

        Try
            'Vérifier si le noeud est valide
            If oTreeNode IsNot Nothing Then
                'Vérifier si le noeud est une TABLE
                If oTreeNode.Tag.ToString = "[TABLE]" Then
                    'Activer les boutons
                    btnAjouterContrainte.Enabled = True
                    btnEnlever.Enabled = True
                    btnExecuter.Enabled = True
                    btnExecuterBatch.Enabled = True
                    'Désactiver les boutons
                    btnAjouterRequete.Enabled = False
                    btnDetruire.Enabled = False
                    btnModifier.Enabled = False
                    btnDeplacerRequeteBas.Enabled = False
                    btnDeplacerRequeteHaut.Enabled = False
                    'Déclarer un noeud de travail
                    Dim oNode As TreeNode = Nothing
                    'Vérifier tous les enfants
                    For Each oNode In oTreeNode.Nodes
                        'Compter le nombre de contraintes
                        iNbContraintes = iNbContraintes + oNode.Nodes.Count
                    Next
                    'Vider la mémoire
                    oNode = Nothing
                    'Afficher l'information
                    lblInformation.Text = iNbContraintes.ToString & "  Contrainte(s), " & oTreeNode.ToolTipText.Replace("[X]", oTreeNode.Nodes.Count.ToString)

                    'Si le noeud est une CONTRAINTE
                ElseIf oTreeNode.Tag.ToString = "[CONTRAINTE]" Then
                    'Activer les boutons
                    btnEnlever.Enabled = True
                    btnDetruire.Enabled = True
                    btnExecuter.Enabled = True
                    btnExecuterBatch.Enabled = True
                    'Désactiver les boutons
                    btnAjouterContrainte.Enabled = False
                    btnAjouterRequete.Enabled = False
                    btnModifier.Enabled = False
                    btnDeplacerRequeteBas.Enabled = False
                    btnDeplacerRequeteHaut.Enabled = False
                    'Afficher l'information
                    lblInformation.Text = "1 Contrainte, " & oTreeNode.Nodes.Count.ToString & "  attribut(s)"

                    'Si le noeud est une REQUETES
                ElseIf oTreeNode.Tag.ToString = "[REQUETES]" Then
                    'Activer les boutons
                    btnAjouterRequete.Enabled = True
                    btnModifier.Enabled = True
                    btnExecuter.Enabled = True
                    btnExecuterBatch.Enabled = True
                    'Désactiver les boutons
                    btnAjouterContrainte.Enabled = False
                    btnEnlever.Enabled = False
                    btnDetruire.Enabled = False
                    btnDeplacerRequeteBas.Enabled = False
                    btnDeplacerRequeteHaut.Enabled = False
                    'Afficher l'information
                    lblInformation.Text = "1 Contrainte, " & oTreeNode.Nodes.Count.ToString & "  requête(s)" & ", " & oTreeNode.ToolTipText

                    'Si le noeud est une REQUETE
                ElseIf oTreeNode.Tag.ToString = "[REQUETE]" Then
                    'Activer les boutons
                    btnAjouterRequete.Enabled = True
                    btnModifier.Enabled = True
                    btnExecuter.Enabled = True
                    btnExecuterBatch.Enabled = True
                    'Désactiver les boutons
                    btnEnlever.Enabled = False
                    btnAjouterContrainte.Enabled = False
                    'Vérifier le nombre de requêtes
                    If oTreeNode.Parent.Nodes.Count > 1 Then
                        'Activer le bouton
                        btnDetruire.Enabled = True
                    Else
                        'Désactiver le bouton
                        btnDetruire.Enabled = False
                    End If
                    'Vérifier si on doit activer ou désactiver le déplacement vers la bas
                    If oTreeNode.Parent.Nodes.IndexOf(oTreeNode) < oTreeNode.Parent.Nodes.Count - 1 Then
                        btnDeplacerRequeteBas.Enabled = True
                    Else
                        btnDeplacerRequeteBas.Enabled = False
                    End If
                    'Vérifier si on doit activer ou désactiver le déplacement vers la bas
                    If oTreeNode.Parent.Nodes.IndexOf(oTreeNode) > 0 Then
                        btnDeplacerRequeteHaut.Enabled = True
                    Else
                        btnDeplacerRequeteHaut.Enabled = False
                    End If
                    'Liste des champs d'une requête de contrainte d'intégrité.
                    Dim sParam() As String = Nothing
                    'Séparer les champs de la requête
                    Using rdr As New IO.StringReader(oTreeNode.Text)
                        Using parser As New FileIO.TextFieldParser(rdr)
                            parser.TextFieldType = FileIO.FieldType.Delimited
                            parser.Delimiters = New String() {" "}
                            parser.HasFieldsEnclosedInQuotes = True
                            sParam = parser.ReadFields()
                        End Using
                    End Using
                    'Afficher l'information
                    lblInformation.Text = "1 Contrainte, " & oTreeNode.Text.Length.ToString & " caractère(s), " & (sParam.Length - 1).ToString & "  paramètre(s)"
                    'Vider la mémoire
                    sParam = Nothing

                    'Si le noeud est un ATTRIBUT
                ElseIf oTreeNode.Tag.ToString = "[ATTRIBUT]" Then
                    'Désactiver les boutons
                    btnAjouterContrainte.Enabled = False
                    btnEnlever.Enabled = False
                    btnDetruire.Enabled = False
                    btnModifier.Enabled = False
                    btnExecuter.Enabled = False
                    btnExecuterBatch.Enabled = False
                    btnDeplacerRequeteBas.Enabled = False
                    btnDeplacerRequeteHaut.Enabled = False
                    'Afficher l'information
                    lblInformation.Text = "1 Contrainte, " & oTreeNode.Nodes.Count.ToString & "  valeur(s)"

                    'Si le noeud est une VALEUR
                ElseIf oTreeNode.Tag.ToString = "[VALEUR]" Then
                    'Désactiver les boutons
                    btnAjouterContrainte.Enabled = False
                    btnAjouterRequete.Enabled = False
                    btnModifier.Enabled = False
                    btnEnlever.Enabled = False
                    btnDetruire.Enabled = False
                    btnExecuter.Enabled = False
                    btnExecuterBatch.Enabled = False
                    btnDeplacerRequeteBas.Enabled = False
                    btnDeplacerRequeteHaut.Enabled = False
                    'Extraire la position de l'attribut
                    iPosAtt = m_TableContraintes.Table.FindField(oTreeNode.Name)
                    'vérifer si l'attribut est présent
                    If iPosAtt >= 0 Then
                        'vérifier si l'attribut est éditable
                        If m_TableContraintes.Table.Fields.Field(iPosAtt).Editable Then
                            'Activer les boutons
                            btnModifier.Enabled = True
                        End If
                    End If
                    'Afficher l'information
                    lblInformation.Text = "1 Contrainte, " & oTreeNode.Text.Length.ToString & "  caractère(s)"

                    'Si le noeud est TRIER
                ElseIf oTreeNode.Tag.ToString.Contains("[TRIER]") Then
                    'Activer les boutons
                    btnEnlever.Enabled = True
                    btnExecuter.Enabled = True
                    btnExecuterBatch.Enabled = True
                    'Désactiver les boutons
                    btnAjouterContrainte.Enabled = False
                    btnAjouterRequete.Enabled = False
                    btnDetruire.Enabled = False
                    btnModifier.Enabled = False
                    btnDeplacerRequeteBas.Enabled = False
                    btnDeplacerRequeteHaut.Enabled = False
                    'Afficher l'information
                    lblInformation.Text = oTreeNode.Nodes.Count.ToString & "  Contrainte(s)"
                End If

                'Vérifier si la classe de sélection est invalide
                If m_FeatureLayer Is Nothing Then
                    'Désactiver les boutons
                    btnAjouterContrainte.Enabled = False
                    btnAjouterRequete.Enabled = False
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw erreur
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de modifier la valeur d'un attribut à partir d'un noeud de type [VALEUR], [REQUETE] OU  [REQUETES].
    '''</summary>
    '''
    '''<param name="oTreeNode">Noeud d'un TreeView de type [VALEUR], [REQUETE] OU  [REQUETES] à modifier.</param>
    ''' 
    ''' <remarks>
    ''' Rien ne sera modifié si le TAG du noeud n'est pas [VALEUR], [REQUETE] OU  [REQUETES].
    '''</remarks>
    Private Sub ModifierNoeud(ByRef oTreeNode As TreeNode)
        'Déclarer les variables de travail
        Dim oTreeNodeTable As TreeNode = Nothing        'Objet contenant le Node du TreeView pour la table des contraintes.
        Dim oTreeNodeContrainte As TreeNode = Nothing   'Objet contenant le Node du TreeView pour la contrainte.
        Dim oTreeNodeAttribut As TreeNode = Nothing     'Objet contenant le Node du TreeView pour un attribut de la contrainte.
        Dim oTreeNodeOid As TreeNode = Nothing          'Objet contenant le Node du TreeView pour un attribut OID.
        Dim oNode As TreeNode = Nothing                 'Contient un noeud de travail.
        Dim pRow As IRow = Nothing                      'Interface contenant la contrainte d'intégrité de la table.
        Dim oInputBoxForm As InputBoxForm = Nothing     'Objet utilisé pour répondre à une question. 
        Dim sNomAttribut As String = ""     'Contient le nom de l'attribut à modifier.
        Dim sValeur As String = ""          'Valeur de l'attribut de la table des contraintes à modifier.
        Dim iPosAtt As Integer = -1         'Position de l'attribut de la table des contraintes.

        Try
            'Vérifier si le noeud est valide
            If oTreeNode IsNot Nothing Then
                'Vérifier si le nom de l'attribut est [VALEUR]
                If oTreeNode.Tag.ToString = "[VALEUR]" Then
                    'Définir le nom de l'attribut à modifier
                    sNomAttribut = oTreeNode.Name
                    'Définir la valeur de l'attribut
                    sValeur = oTreeNode.Text
                    'Définir le noeud contenant un attribut de la contrainte
                    oTreeNodeAttribut = oTreeNode.Parent
                    'Définir le noeud contenant la contrainte
                    oTreeNodeContrainte = oTreeNodeAttribut.Parent
                    'Définir le noeud contenant la table
                    oTreeNodeTable = oTreeNodeContrainte.Parent.Parent

                    'Si le nom de l'attribut est [REQUETE]
                ElseIf oTreeNode.Tag.ToString = "[REQUETE]" Then
                    'Définir le nom de l'attribut à modifier
                    sNomAttribut = oTreeNode.Name
                    'Définir la valeur de l'attribut
                    sValeur = oTreeNode.Text
                    'Définir le noeud contenant un attribut de la contrainte
                    oTreeNodeAttribut = oTreeNode.Parent

                    'Si le nom de l'attribut est [REQUETES]
                ElseIf oTreeNode.Tag.ToString = "[REQUETES]" Then
                    'Définir le nom de l'attribut à modifier
                    sNomAttribut = m_AttributRequetes
                    'Traiter toutes les requêtes
                    For Each oNode In oTreeNode.Nodes
                        'Vérifier si aucune valeur présente
                        If sValeur.Length = 0 Then
                            'Définir la valeur de l'attribut
                            sValeur = oNode.Text
                        Else
                            'Définir la valeur de l'attribut
                            sValeur = sValeur & vbCrLf & oNode.Text
                        End If
                    Next
                    'Définir le noeud contenant un attribut de la contrainte
                    oTreeNodeAttribut = oTreeNode

                    'Sortir sinon
                Else
                    Exit Sub
                End If

                'Définir le noeud contenant la contrainte
                oTreeNodeContrainte = oTreeNodeAttribut.Parent

                'Définir le noeud contenant la table
                oTreeNodeTable = oTreeNodeContrainte.Parent.Parent

                'Vérifier si la table correspond à celle active
                If oTreeNodeTable.Text = cboTableContraintes.Text Then
                    'Définir le noeud contenant le OID
                    oTreeNodeOid = oTreeNodeContrainte.FirstNode.FirstNode
                    'Extraire la contrainte de la table
                    pRow = m_TableContraintes.Table.GetRow(CInt(oTreeNodeOid.Text))
                    'Définir la position de l'attribut des requêtes
                    iPosAtt = m_TableContraintes.Table.FindField(sNomAttribut)

                    'Vérifier si l'attribut est éditable
                    If m_TableContraintes.Table.Fields.Field(iPosAtt).Editable Then
                        'Déclarer le menu InputBoxForm
                        oInputBoxForm = New InputBoxForm("Modifier la valeur de l'attribut :", _
                                                         treContraintes.SelectedNode.Name, sValeur)
                        'Vérifier si le bouton OK a été sélectionné
                        If oInputBoxForm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                            'Vérifier si le noeud sélectionné est [VALEUR]
                            If oTreeNode.Tag.ToString = "[VALEUR]" Then
                                'Modifier la valeur de l'attribut 
                                pRow.Value(iPosAtt) = oInputBoxForm.Reponse.Replace(Chr(10), vbCrLf)

                                'Modifier la valeur de l'attribut dans le TreeView
                                treContraintes.SelectedNode.Text = oInputBoxForm.Reponse.Replace(Chr(10), vbCrLf)

                                'si le noeud sélectionné est [REQUETE]
                            ElseIf oTreeNode.Tag.ToString = "[REQUETE]" Then
                                'Modifier la requête parmis les requêtes
                                pRow.Value(iPosAtt) = pRow.Value(iPosAtt).ToString.Replace(sValeur, oInputBoxForm.Reponse)

                                'Modifier la valeur de l'attribut dans le TreeView
                                treContraintes.SelectedNode.Text = oInputBoxForm.Reponse

                                'Si le noeud sélectionné est [REQUETES]
                            ElseIf oTreeNode.Tag.ToString = "[REQUETES]" Then
                                'Modifier la valeur de l'attribut 
                                pRow.Value(iPosAtt) = oInputBoxForm.Reponse.Replace(Chr(10), vbCrLf)

                                'Détruire les noeuds
                                oTreeNode.Nodes.Clear()
                                'Modifier la valeur de l'attribut
                                Call AjouterRequetesNoeud(oTreeNode, m_AttributRequetes, oInputBoxForm.Reponse.Replace(Chr(10), vbCrLf))
                            End If

                            'Sauver la modification
                            pRow.Store()
                        End If
                        'Vider la mémoire
                        oInputBoxForm = Nothing
                    End If
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw erreur
        Finally
            'Vider la mémoire
            oTreeNodeTable = Nothing
            oTreeNodeContrainte = Nothing
            oTreeNodeAttribut = Nothing
            oTreeNodeOid = Nothing
            oNode = Nothing
            pRow = Nothing
            oInputBoxForm = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de remplir un ComboBox à partir des attributs de la table des contraintes
    ''' et retourner le numéro de l'index correspondant au texte à trouver.
    '''</summary>
    '''
    '''<param name="qComboBox">ComboBox à remplir.</param>
    '''<param name="sTexte">Texte utilisé pour trouver la table par défaut.</param>
    '''
    ''' <returns>L'index correspondant au texte pour trouver la table, -1 par défaut</returns>
    ''' 
    ''' <remarks>
    ''' Si le texte est absent, le dernier de la liste sera utilisé comme celui par défaut.
    '''</remarks>
    Private Function RemplirComboBoxAttribut(ByRef qComboBox As ToolStripComboBox, ByVal sTexte As String) As Integer
        'Déclarer la variables de travail

        'Contient l'item à sélectionner par défaut.
        RemplirComboBoxAttribut = -1

        Try
            'Initialiser le ComboBox
            qComboBox.Items.Clear()

            'Vérifier si la table des contraintes est valide
            If m_TableContraintes IsNot Nothing Then
                'Traiter tous les StandaloneTable
                For i = 0 To m_TableContraintes.Table.Fields.FieldCount - 1
                    'Ajouter l'attribut dans le ComboBox
                    qComboBox.Items.Add(m_TableContraintes.Table.Fields.Field(i).Name)

                    'Vérifier la présence du texte recherché pour la valeur par défaut
                    If m_TableContraintes.Table.Fields.Field(i).Name.ToUpper = sTexte.ToUpper Then
                        'Définir la valeur par défaut
                        RemplirComboBoxAttribut = qComboBox.Items.Count - 1
                    End If
                Next
            End If

            'Définir le nom de l'attribut
            qComboBox.Text = sTexte

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de remplir un ComboBox à partir des StandaloneTables présentes dans la Map active
    ''' et retourner le numéro de l'index correspondant au texte à trouver.
    '''</summary>
    '''
    '''<param name="qComboBox">ComboBox à remplir.</param>
    '''<param name="sTexte">Texte utilisé pour trouver la table par défaut.</param>
    '''
    ''' <returns>L'index correspondant au texte pour trouver la table, -1 par défaut</returns>
    '''
    ''' <remarks>
    ''' Si le texte est absent, le dernier de la liste sera utilisé comme celui par défaut.
    '''</remarks>
    Private Function RemplirComboBoxTable(ByRef qComboBox As ComboBox, ByVal sTexte As String) As Integer
        'Déclarer la variables de travail
        Dim pStandaloneTableColl As IStandaloneTableCollection = Nothing    'Interface qui permet d'extraire les StandaloneTable de la Map.
        Dim pStandaloneTable As IStandaloneTable = Nothing                  'Interface contenant une classe de données.
        Dim pDataset As IDataset = Nothing                                  'Interface contenant le nom complet de la table.
        Dim sNomTable As String = ""                                        'Contient le nom de la table

        'Contient l'item à sélectionner par défaut.
        RemplirComboBoxTable = -1

        Try
            'Initialiser le ComboBox
            qComboBox.Items.Clear()

            'Ajouter le mot clé de la table des contraintes par défaut dans le ComboBox
            qComboBox.Items.Add("CONTRAINTE_INTEGRITE_SPATIALE")

            'Définir la liste des StandaloneTable
            pStandaloneTableColl = CType(m_MxDocument.FocusMap, IStandaloneTableCollection)

            'Traiter tous les StandaloneTable
            For i = 0 To pStandaloneTableColl.StandaloneTableCount - 1
                'Définir le StandaloneTable
                pStandaloneTable = pStandaloneTableColl.StandaloneTable(i)

                'Vérifier si la table est valide
                If pStandaloneTable.Table IsNot Nothing Then
                    'Vérifier si la table contient des OID
                    If pStandaloneTable.Table.HasOID Then
                        'Interface pour extraire le nom complet
                        pDataset = CType(pStandaloneTable.Table, IDataset)

                        'Vérifier si la table contient l'attribut "REQUETES"
                        If pStandaloneTable.Table.FindField("REQUETES") >= 0 Then
                            'Vérifier la présence de Path de la Géodatabase
                            If pDataset.Workspace.PathName.Length = 0 Then
                                'Définir le nom complet de la table
                                sNomTable = pDataset.Name
                            Else
                                'Définir le nom complet de la table
                                sNomTable = pDataset.Workspace.PathName & "\" & pDataset.Name
                            End If
                            'Ajouter le StandaloneTable dans le ComboBox
                            qComboBox.Items.Add(sNomTable)

                            'Vérifier la présence du texte recherché pour la valeur par défaut
                            If sNomTable.ToUpper = sTexte.ToUpper Then
                                'Définir la valeur par défaut
                                RemplirComboBoxTable = qComboBox.Items.Count - 1
                            End If
                        End If
                    End If
                End If
            Next

            'Vérifier si le texte correspond à la table par défaut
            If sTexte = "CONTRAINTE_INTEGRITE_SPATIALE" Then
                'Contient l'item à sélectionner par défaut.
                RemplirComboBoxTable = 0
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pStandaloneTableColl = Nothing
            pStandaloneTable = Nothing
            pDataset = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de remplir un ComboBox à partir des Géodatabases présentes dans la Map active
    ''' et retourner le numéro de l'index correspondant au texte à trouver.
    '''</summary>
    '''
    '''<param name="qComboBox">ComboBox à remplir.</param>
    '''<param name="sTexte">Texte utilisé pour trouver la Geodatabase par défaut.</param>
    '''
    ''' <returns>L'index correspondant au texte pour trouver la table, -1 par défaut</returns>
    '''
    ''' <remarks>
    ''' Si le texte est absent, le dernier de la liste sera utilisé comme celui par défaut.
    '''</remarks>
    Private Function RemplirComboBoxGeodatabase(ByRef qComboBox As ComboBox, ByVal sTexte As String) As Integer
        'Déclarer la variables de travail
        Dim qGeodatabaseColl As Collection = Nothing        'Contient la liste des Géodatabases.
        Dim pTableCollection As ITableCollection = Nothing  'Interface contenant une collection de tables.
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer.
        Dim pDataset As IDataset = Nothing                  'Interface utilisé pour extraire le Workspace.
        Dim pWorkspace As IWorkspace = Nothing              'Interface contenant une Géodatabase.

        'Contient l'item à sélectionner par défaut
        RemplirComboBoxGeodatabase = -1

        Try
            'Initialiser le ComboBox
            qComboBox.Items.Clear()

            'Ajouter le mot clé de la Géodatabase PRO dans le ComboBox
            qComboBox.Items.Add("BDRS_PRO_BDG_DBA")

            'Ajouter le mot clé de la Géodatabase TST dans le ComboBox
            qComboBox.Items.Add("BDRS_TST_BDG_DBA")

            'Créer une nouvelle collection de géodatabase vide
            qGeodatabaseColl = MapGeodatabaseColl(m_MxDocument.FocusMap)

            'Traiter toutes les Géodatabases trouvées
            For i = 1 To qGeodatabaseColl.Count
                'Définir la Géodatabase
                pWorkspace = CType(qGeodatabaseColl.Item(i), IWorkspace)

                'Vérifier si le Path de la Géodatabase est présente
                If pWorkspace.PathName.Length > 0 Then
                    'Ajouter la Géodatabase dans le ComboBox
                    qComboBox.Items.Add(pWorkspace.PathName)

                    'Vérifier la présence du texte recherché pour la valeur par défaut
                    If pWorkspace.PathName.ToUpper = sTexte.ToUpper Then
                        'Définir la valeur par défaut
                        RemplirComboBoxGeodatabase = qComboBox.Items.Count - 1
                    End If
                End If
            Next

            'Si le texte correspond à la Géodatabase PRO
            If sTexte = "BDRS_PRO_BDG_DBA" Then RemplirComboBoxGeodatabase = 0

            'Si le texte correspond à la Géodatabase TST
            If sTexte = "BDRS_TST_BDG_DBA" Then RemplirComboBoxGeodatabase = 1

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            qGeodatabaseColl = Nothing
            pTableCollection = Nothing
            pLayer = Nothing
            pDataset = Nothing
            pWorkspace = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de retourner des Géodatabases des FeatureLayers et des StandaloneTables présents dans la Map active.
    '''</summary>
    '''
    '''<param name="pMap"> Map contenant les FeatureLayer à traiter.</param>
    '''
    ''' <returns> Collection contenant les Géodatabases</returns>
    ''' 
    ''' <remarks>
    ''' Si aucun FeatureLayer n'est présent, la collection sera vide.
    '''</remarks>
    Private Function MapGeodatabaseColl(ByVal pMap As IMap) As Collection
        'Déclarer la variables de travail
        Dim oMapLayer As clsGererMapLayer = Nothing         'Objet utilisé pour extraire les FeatureLayers.
        Dim oFeatureLayerColl As Collection = Nothing       'Objet contenant une collection de FeatureLayers.
        Dim pTableCollection As ITableCollection = Nothing  'Interface contenant une collection de tables.
        Dim pTable As ITable = Nothing                      'Interface contenant une table.
        Dim pDataset As IDataset = Nothing                  'Interface utilisé pour extraire le Workspace.

        'Définir une collection vide de Géodatabases
        MapGeodatabaseColl = New Collection

        Try
            'Définir l'objet pour extraire les FeatureLayers
            oMapLayer = New clsGererMapLayer(pMap, True)

            'Définir la collection des FeatureLayers
            oFeatureLayerColl = oMapLayer.FeatureLayerCollection

            'Traiter tous les StandaloneTable
            For i = 1 To oFeatureLayerColl.Count
                'Interface pour extraire la Géodatabase
                pDataset = CType(oFeatureLayerColl.Item(i), IDataset)

                'Vérifier si c'est une Géodatabase
                If pDataset.Workspace.Type <> esriWorkspaceType.esriFileSystemWorkspace Then
                    'Vérifier si la géodatabase est absente de la colelction
                    If Not MapGeodatabaseColl.Contains(pDataset.Workspace.PathName) Then
                        'Ajouter la Géodatabase
                        MapGeodatabaseColl.Add(pDataset.Workspace, pDataset.Workspace.PathName)
                    End If
                End If
            Next

            'Définir la liste des Table
            pTableCollection = CType(pMap, ITableCollection)

            'Traiter tous les StandaloneTable
            For i = 0 To pTableCollection.TableCount - 1
                'Définir le StandaloneTable
                pTable = CType(pTableCollection.Table(i), ITable)
                'vérifier si la table est valide
                If pTable IsNot Nothing Then
                    'Vérifier si la table possède des OIDs
                    If pTable.HasOID Then
                        'Interface pour extraire la Géodatabase
                        pDataset = CType(pTable, IDataset)

                        'Vérifier si c'est une Géodatabase
                        If pDataset.Workspace.Type <> esriWorkspaceType.esriFileSystemWorkspace Then
                            'Vérifier si la géodatabase est absente de la colelction
                            If Not MapGeodatabaseColl.Contains(pDataset.Workspace.PathName) Then
                                'Ajouter la Géodatabase
                                MapGeodatabaseColl.Add(pDataset.Workspace, pDataset.Workspace.PathName)
                            End If
                        End If
                    End If
                End If
            Next

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            oMapLayer = Nothing
            oFeatureLayerColl = Nothing
            pTableCollection = Nothing
            pTable = Nothing
            pDataset = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'importer les contraintes d'intégrité spatiales dans un TreeView
    ''' et retourner le noeud du TreeView de la table.
    '''</summary>
    '''
    '''<param name="pStandaloneTable"> Nom de la table contenant les contraintes.</param>
    '''<param name="sNomAttributTrier"> Nom de l'attribut utilisé pour trier les contraintes.</param>
    '''<param name="sNomAttributDecrire"> Nom de l'attribut utilisé pour décrire les contraintes.</param>
    '''<param name="sNomAttributRequetes"> Nom de l'attribut contenant les requêtes spatiales</param>
    ''' 
    Private Function ImporterContraintes(ByVal pStandaloneTable As IStandaloneTable, _
                                    Optional ByVal sNomAttributTrier As String = "NOM_TABLE", _
                                    Optional ByVal sNomAttributDecrire As String = "GROUPE",
                                    Optional ByVal sNomAttributRequetes As String = "REQUETES") As TreeNode
        'Déclarer la variables de travail
        Dim pMouseCursor As IMouseCursor = Nothing          'Interface qui permet de changer le curseur de la souris.
        Dim pDataset As IDataset = Nothing                  'Interface qui permet d'ouvrir une table.
        Dim pTableSelection As ITableSelection = Nothing    'Interface qui permet d'extraire les contraintes sélectionnées.
        Dim pQueryFilter As IQueryFilterDefinition = Nothing 'Interface utilisé pour trier les contraintes.
        Dim pCursor As ICursor = Nothing                    'Interface contenant les contraintes à afficher.
        Dim pRow As IRow = Nothing                          'Interface contenant l'information d'une contrainte.
        Dim oTreeNodeTable As TreeNode = Nothing            'Objet contenant le Node du TreeView pour la table des contraintes.
        Dim iPosAttTrier As Integer = -1        'Contient la position de l'attribut pour trier les contraintes.
        Dim iPosAttDecrire As Integer = -1      'Contient la position de l'attribut pour décrire les contraintes.

        'Définir le noeud de la table par défaut
        ImporterContraintes = Nothing

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Vérifier si la table est valide
            If pStandaloneTable IsNot Nothing Then
                'On recharge la table suite au ca ou il y a eu des changement
                pDataset = CType(pStandaloneTable.Table, IDataset)
                pStandaloneTable.Table = CType(pDataset.FullName.Open(), ITable)

                'Désactiver l'affichage du TreeView
                treContraintes.BeginUpdate()

                'Définir le noeud de la table
                oTreeNodeTable = AjouterTableNoeud(pStandaloneTable, sNomAttributTrier)

                'Définir la position de l'attribut pour regrouper les contraintes
                iPosAttTrier = pStandaloneTable.Table.FindField(sNomAttributTrier)

                'Définir la position de l'attribut pour décrire les contraintes
                iPosAttDecrire = pStandaloneTable.Table.FindField(sNomAttributDecrire)

                'Définir la tables des contraintes d'intégrités spatiales
                pTableSelection = CType(pStandaloneTable, ITableSelection)

                'vérifier si une sélection est présente
                If pTableSelection.SelectionSet.Count = 0 Then
                    'Sélectionner toutes les contraintes
                    pTableSelection.SelectRows(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                End If

                'Définir les noms d'attributs utilisés pour trier la table des contraintes
                pQueryFilter = New QueryFilter
                pQueryFilter.PostfixClause = "ORDER BY " & sNomAttributTrier & "," & sNomAttributDecrire

                'Interface pour extraire les contraintes
                pTableSelection.SelectionSet.Search(CType(pQueryFilter, IQueryFilter), False, pCursor)

                'Extraire la première contrainte
                pRow = pCursor.NextRow

                'Traiter toutes les contraintes
                Do Until pRow Is Nothing
                    'Ajouter une contrainte d'intégrité dans le TreeView
                    Call AjouterContrainteNoeud(oTreeNodeTable, pRow, iPosAttTrier, iPosAttDecrire, sNomAttributRequetes)

                    'Extraire la prochaine contrainte
                    pRow = pCursor.NextRow
                Loop

                'Ouvrir le noeud de la table des contraintes
                oTreeNodeTable.Expand()

                'Activer l'affichage du TreeView
                treContraintes.EndUpdate()

                'Sélectionner la table dans le TreeView
                treContraintes.SelectedNode = oTreeNodeTable

                'Retourner le noeud de la table
                ImporterContraintes = oTreeNodeTable
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
            pDataset = Nothing
            pTableSelection = Nothing
            pQueryFilter = Nothing
            pCursor = Nothing
            pRow = Nothing
            treContraintes.EndUpdate()
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'ajouter le noeud de la table des contraintes d'intégrité spatiales dans un TreeView.
    '''</summary>
    '''
    '''<param name="pStandaloneTable"> Nom de la table contenant les contraintes.</param>
    '''<param name="sNomAttributTrier"> Nom de l'attribut utilisé pour trier les contraintes.</param>
    '''
    '''<returns>TreeNode de la table des contraintes d'intégrité.</returns>
    '''
    '''<remarks>Si le noeud de la table est déjà présente, elle sera vidé et retourné.</remarks>
    '''
    Private Function AjouterTableNoeud(ByVal pStandaloneTable As IStandaloneTable, ByVal sNomAttributTrier As String) As TreeNode
        'Déclarer la variables de travail
        Dim pDataset As IDataset = Nothing      'Interface contenant le nom complet de la table.
        Dim sNomTable As String = ""            'Contient le nom complet de la table.

        'Définir le noeud de la table par défaut
        AjouterTableNoeud = Nothing

        Try
            'Vérifier si la table est valide
            If pStandaloneTable IsNot Nothing Then
                'Interface pour extraire le nom complet de la table
                pDataset = CType(pStandaloneTable.Table, IDataset)
                'Vérifier la présence du Path de la Géodatabase
                If pDataset.Workspace.PathName.Length = 0 Then
                    'Définir le nom complet de la table
                    sNomTable = pDataset.Name
                Else
                    'Définir le nom complet de la table
                    sNomTable = pDataset.Workspace.PathName & "\" & pDataset.Name
                End If

                'Vérifier si la valeur pour la table des contraintes est présente
                If treContraintes.Nodes.ContainsKey(sNomTable) Then
                    'Extraire le noeud existant pour la table des contraintes
                    AjouterTableNoeud = treContraintes.Nodes.Item(sNomTable)
                    'Vider tous les Noeuds de la table des contraintes
                    AjouterTableNoeud.Nodes.Clear()
                Else
                    'Créer un nouveau noeud pour la table des contraintes
                    AjouterTableNoeud = treContraintes.Nodes.Add(sNomTable, sNomTable)
                End If

                'Définir le type de noeud
                AjouterTableNoeud.Tag = "[TABLE]"

                'Conserver le nom de l'attribut utilisé pour trier les contraintes dans le noeud de la table
                AjouterTableNoeud.ToolTipText = "[X] " & sNomAttributTrier.ToLower & "(s)"
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pDataset = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'ajouter une contrainte d'intégrité spatiale dans un TreeView à partir du noeud de la table.
    '''</summary>
    '''
    '''<param name="oTreeNodeTable"> Objet contenant le Node du TreeView pour la table des contraintes.</param>
    '''<param name="pRow"> Interface contenant l'information d'une contrainte.</param>
    '''<param name="iPosAttTrier"> Contient la position de l'attribut pour trier les contraintes.</param>
    '''<param name="iPosAttDecrire"> Contient la position de l'attribut pour décrire les contraintes.</param>
    '''<param name="sNomAttributRequetes"> Nom de l'attribut contenant les requêtes spatiales.</param>
    '''
    '''<returns>TreeNode de la contrainte d'intégrité créée.</returns>
    '''
    '''<remarks>Si le noeud de la table ou la contrainte est invalide, Nothing sera retourné.</remarks>
    '''
    Private Function AjouterContrainteNoeud(ByRef oTreeNodeTable As TreeNode, ByVal pRow As IRow, _
                                            ByVal iPosAttTrier As Integer, ByVal iPosAttDecrire As Integer, _
                                            ByVal sNomAttributRequetes As String) As TreeNode
        'Déclarer la variables de travail
        Dim oTreeNodeTrier As TreeNode = Nothing            'Objet contenant le Node du TreeView pour trier les contraintes.
        Dim oTreeNodeContrainte As TreeNode = Nothing       'Objet contenant le Node du TreeView pour décrire une contrainte.
        Dim oTreeNodeAttribut As TreeNode = Nothing         'Objet contenant le Node du TreeView pour un attribut d'une contrainte.
        Dim oTreeNode As TreeNode = Nothing                 'Objet contenant le Node du TreeView pour la valeur d'un atribut d'une contrainte.

        'Définir le noeud de la contrainte par défaut
        AjouterContrainteNoeud = Nothing

        Try
            'Vérifier si la table est valide
            If oTreeNodeTable IsNot Nothing And pRow IsNot Nothing Then
                'Vérifier si la valeur pour trier est présente
                If oTreeNodeTable.Nodes.ContainsKey(pRow.Value(iPosAttTrier).ToString) Then
                    'Extraire le noeud existant pour trier les contraintes
                    oTreeNodeTrier = oTreeNodeTable.Nodes.Item(pRow.Value(iPosAttTrier).ToString)
                Else
                    'Créer un nouveau noeud pour trier les contraintes
                    oTreeNodeTrier = oTreeNodeTable.Nodes.Add(pRow.Value(iPosAttTrier).ToString, pRow.Value(iPosAttTrier).ToString)
                End If
                'Définir le type de noeud
                oTreeNodeTrier.Tag = "[TRIER]:" & m_AttributTrier

                'Créer le noeud pour décrire la contrainte
                oTreeNodeContrainte = oTreeNodeTrier.Nodes.Add(pRow.Value(iPosAttDecrire).ToString, pRow.Value(iPosAttDecrire).ToString)
                'Définir le type de noeud
                oTreeNodeContrainte.Tag = "[CONTRAINTE]"

                'Ajouter toute l'information de la contrainte
                For i = 0 To pRow.Fields.FieldCount - 1
                    'Créer le noeud contenant le nom de l'attribut d'information de la contrainte
                    oTreeNodeAttribut = oTreeNodeContrainte.Nodes.Add(pRow.Fields.Field(i).Name & "=", pRow.Fields.Field(i).Name)

                    'Vérifier si l'attribut est celui des requêtes
                    If pRow.Fields.Field(i).Name = sNomAttributRequetes Then
                        'Définir le type de noeud
                        oTreeNodeAttribut.Tag = "[REQUETES]"
                        'Définir le nombre de caractères
                        oTreeNodeAttribut.ToolTipText = pRow.Value(i).ToString.Length.ToString & " caractère(s)"
                        'Ajouter les noeuds des requêtes spatiales
                        Call AjouterRequetesNoeud(oTreeNodeAttribut, pRow.Fields.Field(i).Name, pRow.Value(i).ToString)

                        'Si l'attribut n'est pas celui des requêtes
                    Else
                        'Définir le type de noeud
                        oTreeNodeAttribut.Tag = "[ATTRIBUT]"
                        'Créer le noeud contenant la valeur de l'attribut d'information de la contrainte
                        oTreeNode = oTreeNodeAttribut.Nodes.Add(pRow.Fields.Field(i).Name, pRow.Value(i).ToString)
                        'Définir le type de noeud
                        oTreeNode.Tag = "[VALEUR]"
                    End If

                    'Ouvrir le noeud
                    oTreeNodeAttribut.ExpandAll()
                Next

                'Retourner le noeud de la contrainte créée
                AjouterContrainteNoeud = oTreeNodeContrainte
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            oTreeNodeTrier = Nothing
            oTreeNodeContrainte = Nothing
            oTreeNodeAttribut = Nothing
            oTreeNode = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'ajouter toutes les requêtes spatiales dans un TreeView à partir du noeud des requêtes spatiales
    ''' et de retourner au besoin le noeud de la requête qui a été ajoutée.
    '''</summary>
    '''
    '''<param name="oTreeNodeRequetes"> Objet contenant le Node du TreeView pour les requêtes spatiales.</param>
    '''<param name="sNomAttributRequetes"> Nom de l'attribut contenant les requêtes spatiales.</param>
    '''<param name="sRequetes"> Contient la requête spatiale.</param>
    '''<param name="sRequeteAjouter"> Contient la requête spatiale qui a été ajoutée.</param>
    '''
    '''<returns>TreeNode de la requête spatiale ajoutée.</returns>
    '''
    '''<remarks>Si le noeud des requêtes ou la requête est invalide, Nothing sera retourné.</remarks>
    '''
    Private Function AjouterRequetesNoeud(ByRef oTreeNodeRequetes As TreeNode, ByVal sNomAttributRequetes As String, _
                                          ByVal sRequetes As String, Optional ByVal sRequeteAjouter As String = "") As TreeNode
        'Déclarer la variables de travail
        Dim oTreeNodeRequete As TreeNode = Nothing     'Objet contenant le Node du TreeView pour une requête spatiale.

        'Définir le noeud par défaut
        AjouterRequetesNoeud = Nothing

        Try
            'Vérifier si le noeud de l'attribut est valide
            If oTreeNodeRequetes IsNot Nothing And sRequetes.Length > 0 Then
                'Vérifier le type de noeud
                If oTreeNodeRequetes.Tag.ToString = "[REQUETES]" Then
                    'Traiter toutes les requêtes de la contrainte
                    For Each sRequete In Split(sRequetes, vbCrLf)
                        'Ajouter le noeud de la requête
                        oTreeNodeRequete = AjouterRequeteNoeud(oTreeNodeRequetes, sNomAttributRequetes, sRequete)
                        'Vérifier si le noeud correspond à la requête demandée
                        If sRequete = sRequeteAjouter Then
                            'Retourner la requête ajoutée
                            AjouterRequetesNoeud = oTreeNodeRequete
                        End If
                    Next
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            oTreeNodeRequete = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'ajouter une requête spatiale dans un TreeView à partir du noeud des requêtes spatiales.
    '''</summary>
    '''
    '''<param name="oTreeNodeRequetes"> Objet contenant le Node du TreeView pour les requêtes spatiales.</param>
    '''<param name="sNomAttributRequetes"> Nom de l'attribut contenant les requêtes spatiales.</param>
    '''<param name="sRequete"> Contient la requête spatiale.</param>
    '''
    '''<returns>TreeNode de la requête spatiale créée.</returns>
    '''
    '''<remarks>Si le noeud des requêtes ou la requête est invalide, Nothing sera retourné.</remarks>
    '''
    Private Function AjouterRequeteNoeud(ByRef oTreeNodeRequetes As TreeNode, _
                                         ByVal sNomAttributRequetes As String, ByVal sRequete As String) As TreeNode
        'Déclarer la variables de travail

        'Définir le noeud par défaut
        AjouterRequeteNoeud = Nothing

        Try
            'Vérifier si le noeud de l'attribut est valide
            If oTreeNodeRequetes IsNot Nothing And sRequete.Length > 0 Then
                'Vérifier le type de noeud
                If oTreeNodeRequetes.Tag.ToString = "[REQUETES]" Then
                    'Créer le noeud contenant la valeur de l'attribut d'information de la contrainte
                    AjouterRequeteNoeud = oTreeNodeRequetes.Nodes.Add(sNomAttributRequetes, sRequete)
                    'Définir le type de noeud
                    AjouterRequeteNoeud.Tag = "[REQUETE]"
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
        End Try
    End Function
#End Region
End Class