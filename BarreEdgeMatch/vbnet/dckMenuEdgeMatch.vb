Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem

''' <summary>
''' Designer class of the dockable window add-in. It contains user interfaces that
''' make up the dockable window.
''' </summary>
Public Class dckMenuEdgeMatch
    'Déclarer les variables de travail
    Private m_hook As Object

    ''' <summary>
    ''' Appeler lorsque le menu est créé
    ''' </summary>
    Public Sub New(ByVal hook As Object)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.Hook = hook

        'Définir les variables de travail
        Dim pEditor As IEditor = Nothing

        Try
            'Définir l'application
            m_Application = CType(hook, IApplication)

            'Définir le document
            m_MxDocument = CType(m_Application.Document, IMxDocument)

            'Définir le menu de EdgeMatch
            m_MenuEdgeMatch = Me

            'Appel de l'initialisation du formulaire
            Call InitialiserFormulaire()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pEditor = Nothing
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        Try
            'Détruire les Handlers
            Call DeleteHandler()

        Catch ex As Exception
            'On ne fait rien
        Finally
            MyBase.Finalize()
        End Try
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

        Private m_windowUI As dckMenuEdgeMatch

        Protected Overrides Function OnCreateChild() As System.IntPtr
            m_windowUI = New dckMenuEdgeMatch(Me.Hook)
            Return m_windowUI.Handle
        End Function

        Protected Overrides Sub Dispose(ByVal Param As Boolean)
            If m_windowUI IsNot Nothing Then
                m_windowUI.Dispose(Param)
            End If

            MyBase.Dispose(Param)
        End Sub
    End Class

#Region "Routines et fonctions d'événements"
    Private Sub cboClasseDecoupage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClasseDecoupage.SelectedIndexChanged
        'Déclarer les variables de travail
        Dim pFields As IFields = Nothing    'Interface contenant tous les attributs de la classe de découpage
        Dim i As Integer = Nothing          'Compteur

        Try
            'Remplir les attributs pour l'identifiant de la classe de découpage
            cboIdentifiantDecoupage.Items.Clear()

            'Extraire le FeatureLayer de découpage
            m_FeatureLayerDecoupage = m_MapLayer.ExtraireFeatureLayerByName(cboClasseDecoupage.Text)

            'Vérifier si le FeatureLayer est valide
            If m_FeatureLayerDecoupage IsNot Nothing Then
                'Interface contenant tou sles attributs de la classe de découage
                pFields = m_FeatureLayerDecoupage.FeatureClass.Fields

                'Ajouter tous les attributs de type String
                For i = 0 To pFields.FieldCount - 1
                    'Vérifier si l'attribut est de type string
                    If pFields.Field(i).Type = esriFieldType.esriFieldTypeString Then
                        'Ajouter le nom de l'attribut dans le ComboBox
                        cboIdentifiantDecoupage.Items.Add(pFields.Field(i).Name)
                        'Vérifier si l'attribut contient le mot "Dataset"
                        If InStr(pFields.Field(i).Name.ToUpper, "DATASET".ToUpper) > 0 Then
                            'Ajouter le nom de l'attribut dans le ComboBox
                            cboIdentifiantDecoupage.Text = pFields.Field(i).Name
                        End If
                    End If
                Next

                'Remplir le ListBox de tous les types de FeatureLayer en excluant le FeatureLayer de découpage et en excluant les classes vides.
                RemplirListBoxFeatureLayer(lstClasses, lstAttributAdjacence, "CODE_SPEC", m_FeatureLayerDecoupage.Name, False)
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pFields = Nothing
        End Try
    End Sub

    Private Sub cboIdentifiantDecoupage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboIdentifiantDecoupage.SelectedIndexChanged
        Try
            'Définir l'attribut contenant l'identifiant de découpage
            m_IdentifiantDecoupage = cboIdentifiantDecoupage.Text

            'Vider la liste des attribut d'adjacence
            lstAttributAdjacence.Items.Clear()

            'Remplir le ListBox de tous les types de FeatureLayer en excluant le FeatureLayer de découpage et en excluant les classes vides.
            RemplirListBoxFeatureLayer(lstClasses, lstAttributAdjacence, "CODE_SPEC", m_FeatureLayerDecoupage.Name, False)

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    Private Sub chkAdjacenceUnique_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAdjacenceUnique.CheckedChanged
        'Conserver si un élément doit avoir seulement une adjacence
        m_AdjacenceUnique = chkAdjacenceUnique.Checked
    End Sub

    Private Sub chkClasseDifferente_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkClasseDifferente.CheckedChanged
        'Conserver si on doit permettre les classes différentes aux points d'adjacences
        m_ClasseDifferente = chkClasseDifferente.Checked
    End Sub

    Private Sub chkIdentifiantPareil_CheckedChanged(sender As Object, e As EventArgs) Handles chkIdentifiantPareil.CheckedChanged
        'Conserver si on ne tient pas compte des identifiants
        m_SansIdentifiant = chkIdentifiantPareil.Checked
    End Sub

    Private Sub lstAttributAdjacence_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstAttributAdjacence.SelectedIndexChanged
        'Déclarer les variables de travail
        Dim i As Integer = Nothing  'Compteur

        Try
            'Initialiser la liste des attributs d'adjacence
            m_AttributAdjacence = New Collection

            'Traiter tous les noms d'attributs sélectionnés
            For i = 0 To lstAttributAdjacence.SelectedItems.Count - 1
                'Ajouter le nom de l'attribut sélectionné 
                m_AttributAdjacence.Add(lstAttributAdjacence.SelectedItems.Item(i).ToString)
            Next

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    Private Sub trePoints_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trePoints.AfterSelect
        'Déclarer les variables de travail
        Dim pEnvelope As IEnvelope = Nothing    'Interface ESRI contenant l'enveloppe de la fenêtre graphique
        Dim pPoint As IPoint = Nothing          'Interface ESRI contenant le point d'adjacence sélectionné
        Dim i As Integer = Nothing              'Index du point d'adjacence

        Try
            'Définir l'enveloppe de la fenêtre graphique
            pEnvelope = m_MxDocument.ActiveView.Extent

            'Définir l'index du point d'adjacence sélectionné
            If trePoints.SelectedNode.Parent Is Nothing Then
                i = trePoints.SelectedNode.Index
            Else
                i = trePoints.SelectedNode.Parent.Index
            End If

            'Définir le point d'adjacence sélectionné
            pPoint = CType(m_ListePointAdjacence.Geometry(i), IPoint)
            pPoint.Project(m_MxDocument.FocusMap.SpatialReference)

            'Centrer l'enveloppe selon le point d'adjacence
            pEnvelope.CenterAt(pPoint)

            'Définir la nouvelle fenêtre graphique
            m_MxDocument.ActiveView.Extent = pEnvelope

            'Dessiner la géométrie traitée et son identifiant
            Call bDessinerGeometrie(pPoint, True)
            Call bDessinerTexte(pPoint, pPoint.ID.ToString, False)

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pEnvelope = Nothing
            pPoint = Nothing
        End Try
    End Sub

    Private Sub treErreurs_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treErreurs.AfterSelect
        'Déclarer les variables de travail
        Dim oErreur As Structure_Erreur = Nothing   'Structure contenant une erreur d'attribut
        Dim pEnvelope As IEnvelope = Nothing        'Interface ESRI contenant l'enveloppe de la fenêtre graphique
        Dim pPoint As IPoint = Nothing              'Interface ESRI contenant le point d'adjacence sélectionné
        Dim i As Integer = Nothing                  'Index du point d'adjacence

        Try
            'Définir l'enveloppe de la fenêtre graphique
            pEnvelope = m_MxDocument.ActiveView.Extent

            'Définir l'index de l'erreur d'adjacence sélectionné
            If treErreurs.SelectedNode.Parent Is Nothing Then
                i = treErreurs.SelectedNode.Index + 1
            Else
                i = treErreurs.SelectedNode.Parent.Index + 1
            End If

            'Définir l'erreur sélectionné
            oErreur = CType(m_ErreurFeature.Item(i), Structure_Erreur)

            'Vérifier si le point B est présent
            If oErreur.PointB Is Nothing Then
                'Définir le point d'erreur sélectionné
                pPoint = oErreur.PointA
                pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
            Else
                'Définir le point d'erreur sélectionné
                pPoint = oErreur.PointB
                pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
            End If

            'Centrer l'enveloppe selon le point d'erreur
            pEnvelope.CenterAt(pPoint)

            'Définir la nouvelle fenêtre graphique
            m_MxDocument.ActiveView.Extent = pEnvelope

            'Dessiner la géométrie traitée et son identifiant
            Call bDessinerGeometrie(pPoint, True)
            Call bDessinerTexte(pPoint, i.ToString, False)

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            oErreur = Nothing
            pEnvelope = Nothing
            pPoint = Nothing
        End Try
    End Sub

    Private Sub chkPrecision_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrecision.CheckedChanged, chkAdjacence.CheckedChanged, chkAttribut.CheckedChanged
        Try
            'Vider le treview d'erreurs
            treErreurs.Nodes.Clear()

            'Définir une nouvelle Collection pour les éléments en erreur
            m_ErreurFeature = New Collection

            'Remplir le formulaires contenant les erreurs de précision
            Call RemplirFormulaireErreurPrecision()

            'Remplir le formulaires contenant les erreurs d'adjacence
            Call RemplirFormulaireErreurAdjacence()

            'Remplir le formulaires contenant les erreurs d'attribut
            Call RemplirFormulaireErreurAttribut()

            'Montrer toutes l'information pour toutes les erreurs
            If chkOuvrirErreurs.Checked And treErreurs.Nodes.Count > 0 Then treErreurs.ExpandAll()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Écrire une trace de fin
            'Call citsMod_FichierTrace.EcrireMessageTrace("Fin")
        End Try
    End Sub

    Private Sub chkOuvrirPoints_CheckedChanged(sender As Object, e As EventArgs) Handles chkOuvrirPoints.CheckedChanged
        'Vérifier si les points sont présents
        If trePoints.Nodes.Count > 0 Then
            'Vérifier si on doit ouvrir
            If chkOuvrirPoints.Checked Then
                'Ouvrir toute l'information pour tous les points d'adjacence
                trePoints.ExpandAll()
            Else
                'Fermer toute l'information pour tous les points d'adjacence
                trePoints.CollapseAll()
            End If
        End If
    End Sub

    Private Sub chkOuvrirErreurs_CheckedChanged(sender As Object, e As EventArgs) Handles chkOuvrirErreurs.CheckedChanged
        'Vérifier si les erreurs sont présentes
        If treErreurs.Nodes.Count > 0 Then
            'Vérifier si on doit ouvrir
            If chkOuvrirErreurs.Checked Then
                'Ouvrir toute l'information pour toutes les erreurs
                treErreurs.ExpandAll()
            Else
                'Fermer toute l'information pour toutes erreurs
                treErreurs.CollapseAll()
            End If
        End If
    End Sub

    Private Sub txtTolAdjacence_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTolAdjacence.TextChanged
        'Vérifier si le texte est numérique
        If IsNumeric(txtTolAdjacence.Text) Then
            m_TolAdjacence = CDbl(txtTolAdjacence.Text)
            m_TolAdjacenceOri = m_TolAdjacence
        Else
            txtTolAdjacence.Text = m_TolAdjacenceOri.ToString
        End If
    End Sub

    Private Sub txtTolRecherche_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTolRecherche.TextChanged
        'Vérifier si le texte est numérique
        If IsNumeric(txtTolRecherche.Text) Then
            m_TolRecherche = CDbl(txtTolRecherche.Text)
            m_TolRechercheOri = m_TolRecherche
        Else
            txtTolRecherche.Text = m_TolRechercheOri.ToString
        End If
    End Sub

    Private Sub txtPrecision_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecision.TextChanged
        'Vérifier si le texte est numérique
        If IsNumeric(txtPrecision.Text) Then
            'Changer la précision
            m_Precision = CDbl(txtPrecision.Text)
        Else
            'Remettre l'ancienne valeur
            txtPrecision.Text = m_Precision.ToString
        End If
    End Sub

    Private Sub cmdInitialiser_Click(sender As Object, e As EventArgs) Handles cmdInitialiser.Click
        Try
            'Permet d'initialiser le traitement au complet
            Call InitialiserFormulaire()

            'Initialiser les AddHandler
            Me.InitHandler()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub
    '''<summary>
    ''' Routine qui permet de réinitialiser le menu lorsqu'un nouveau document est ouvert.
    '''</summary>
    '''
    Private Sub OnOpenDocument()
        Try
            'Détruire les Handlers
            RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).SpatialReferenceChanged, AddressOf OnSpatialReferenceChanged
            RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemAdded, AddressOf OnItemAdded
            RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemDeleted, AddressOf OnItemDeleted

            'Définir le document
            m_MxDocument = CType(m_Application.Document, IMxDocument)

            'Initialiser les Handlers
            AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).SpatialReferenceChanged, AddressOf OnSpatialReferenceChanged
            AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemAdded, AddressOf OnItemAdded
            AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemDeleted, AddressOf OnItemDeleted

            'Permet d'initialiser le traitement au complet
            Call InitialiserFormulaire()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser le menu lorsque le contenu change.
    '''</summary>
    '''
    Private Sub OnContentsChanged()
        Try
            'Initialiser le menu
            m_MenuEdgeMatch.InitialiserClasses()

            'Initialiser les tolérances
            m_MenuEdgeMatch.InitialiserTolerance()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser les tolérances lors d'un changement de référence spatiale.
    '''</summary>
    '''
    Private Sub OnSpatialReferenceChanged()
        Try
            'Initialiser les tolérances
            m_MenuEdgeMatch.InitialiserTolerance()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser le menu lorsqu'un nouvel item est ajouter dans la Map active.
    '''</summary>
    '''
    Private Sub OnItemAdded(ByVal Item As Object)
        Try
            'Réinitialiser les classes
            m_MenuEdgeMatch.InitialiserClasses()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser le menu lorsqu'un item est retiré dans la Map active.
    '''</summary>
    '''
    Private Sub OnItemDeleted(ByVal Item As Object)
        Try
            'Réinitialiser les classes
            m_MenuEdgeMatch.InitialiserClasses()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub


    Private Sub dckMenuEdgeMatch_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        'Déclarer les variables de travail
        Dim DockWindow As ESRI.ArcGIS.Framework.IDockableWindow
        Dim windowID As UID = New UIDClass

        Try
            'Créer un nouveau menu
            windowID.Value = "MPO_BarreEdgeMatch_dckMenuEdgeMatch"
            DockWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(windowID)

            'Vérifier si le menu est visible
            If DockWindow.IsVisible() Then
                'Initialiser les AddHandler
                Call InitHandler()

                'si le menu n'est pas visible
            Else
                'Détruire les AddHandler
                Call DeleteHandler()
            End If

        Catch erreur As Exception
            'Message d'erreur
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            DockWindow = Nothing
            windowID = Nothing
        End Try
    End Sub

    Private Sub dckMenuEdgeMatch_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        'Déclarer les variables de travail
        Dim iDeltaHeight As Integer
        Dim iDeltaWidth As Integer

        'Calculer les deltas
        iDeltaHeight = Me.Height - m_Height
        iDeltaWidth = Me.Width - m_Width
        'Redimensionner les objets du menu
        tabEdgeMatch.Height = tabEdgeMatch.Height + iDeltaHeight
        tabEdgeMatch.Width = tabEdgeMatch.Width + iDeltaWidth
        cboClasseDecoupage.Width = cboClasseDecoupage.Width + iDeltaWidth
        cboIdentifiantDecoupage.Width = cboIdentifiantDecoupage.Width + iDeltaWidth
        lstClasses.Height = lstClasses.Height + iDeltaHeight
        lstClasses.Width = lstClasses.Width + iDeltaWidth
        lstAttributAdjacence.Height = lstAttributAdjacence.Height + iDeltaHeight
        lstAttributAdjacence.Width = lstAttributAdjacence.Width + iDeltaWidth
        trePoints.Height = trePoints.Height + iDeltaHeight
        trePoints.Width = trePoints.Width + iDeltaWidth
        treErreurs.Height = treErreurs.Height + iDeltaHeight
        treErreurs.Width = treErreurs.Width + iDeltaWidth
        'Réajustement de la dimension des listbox en fonction de la différence de hauteur initiale
        lstClasses.Height = lstClasses.Height + (Me.Height - lstClasses.Height) - m_ClasseHeight
        lstAttributAdjacence.Height = lstAttributAdjacence.Height + (Me.Height - lstAttributAdjacence.Height) - m_AttributHeight
        'Réajustement de la position de la commande d'initialisation
        cmdInitialiser.Top = cmdInitialiser.Top + iDeltaHeight
        lblNbPointAdjacence.Top = lblNbPointAdjacence.Top + iDeltaHeight
        lblNbErreur.Top = lblNbErreur.Top + iDeltaHeight
        'Initialiser les variables de dimension
        m_Height = Me.Height
        m_Width = Me.Width
    End Sub
#End Region

#Region "Routines et fonctions publiques"
    '''<summary>
    ''' Routine permet d'initialiser la Map et la liste des classes visibles traitées.
    '''</summary>
    ''' 
    Public Sub InitialiserClasses()
        Try
            'Définir le MapLayer et tous les FeatureLayer visibles
            m_MapLayer = New clsGererMapLayer(m_MxDocument.FocusMap, False)

            'Remplir le comboBox de la classe de découpage
            RemplirComboBox(cboClasseDecoupage, "decoupage", esriGeometryType.esriGeometryPolygon)

            'Vider la liste des points d'adjacence et d'erreur
            'trePoints.Nodes.Clear()
            'treErreurs.Nodes.Clear()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'intialiser les tolérances et la précision. 
    '''</summary>
    ''' 
    Public Sub InitialiserTolerance()
        'Déclarer les variables de travail
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing   'Interface contenant la tolérance XY de la référence spatiale.

        Try
            'Définir la tolérance d'adjacence en mètre
            m_TolAdjacence = 3.0
            m_TolAdjacenceOri = 3.0
            'Définir la tolérance de recherche en mètre
            m_TolRecherche = 1.0
            m_TolRechercheOri = 1.0

            'Vérifier si la référence spatiale est valide
            If m_MxDocument.FocusMap.SpatialReference IsNot Nothing Then
                'Interface pour extraire la tolérance de la référence spatiale
                pSpatialRefTol = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceTolerance)
                'Définir la précision
                m_Precision = pSpatialRefTol.XYTolerance
            Else
                'Définir la précision
                m_Precision = 0.001
            End If

            'Remplir le formulaire des précision
            txtPrecision.Text = m_Precision.ToString("0.0#######")
            txtTolAdjacence.Text = m_TolAdjacenceOri.ToString("0.0##")
            txtTolRecherche.Text = m_TolRechercheOri.ToString("0.0##")

        Catch ex As Exception
            'Message d'erreur
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pSpatialRefTol = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de remplir le formulaire pour visualiser les erreurs de précision. 
    '''</summary>
    ''' 
    Public Sub RemplirFormulairePointAdjacence()
        'Déclarer les variables de travail
        Dim qElementTraiter As Structure_Element_Traiter = Nothing  'Objet contenant l'information de l'élément.
        Dim qElementColl As Collection = Nothing    'Objet qui contient la liste des éléments adjacents.
        Dim pFeature As IFeature = Nothing          'Interface contenant l'élément.
        Dim pPoint As IPoint = Nothing              'Interface contenant le sommet traité.
        Dim oNode As TreeNode = Nothing             'Objet contenant le noeud en erreur
        Dim i As Integer = Nothing                  'Compteur

        Try
            'Sortir si le menu n'est pas défini
            If m_MenuEdgeMatch Is Nothing Then Return

            'Traiter tous les points d'adjacence
            For i = 1 To m_ListeElementPointAdjacent.Count
                'Définir le point d'adjacence
                pPoint = CType(m_ListePointAdjacence.Geometry(i - 1), IPoint)

                'Définir la liste des sommets d'éléments adjacents
                qElementColl = CType(m_ListeElementPointAdjacent.Item(i), Collection)

                'Ajouter le point d'adjacence dans le formulaire
                oNode = m_MenuEdgeMatch.trePoints.Nodes.Add("PointID=" & pPoint.ID.ToString)

                'Afficher le nombre d'elements
                oNode.Nodes.Add("Erreur présente = " & m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString).ToString)

                'Afficher le nombre d'elements
                oNode = oNode.Nodes.Add("Nombre de sommets d'éléments = " & qElementColl.Count.ToString)

                'Traiter tous les éléments
                For j = 1 To qElementColl.Count
                    'Définir l'élément
                    qElementTraiter = CType(m_ListeElementTraiter(qElementColl.Item(j)), Structure_Element_Traiter)
                    pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                    'Afficher le nombre d'elements
                    oNode.Nodes.Add("Oid=" & pFeature.OID.ToString & ", Classe=" & pFeature.Class.AliasName)
                Next
            Next

            'Afficher toutes l'information de l'erreur
            'm_MenuEdgeMatch.trePoints.ExpandAll()

        Catch erreur As Exception
            'Message d'erreur
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            qElementTraiter = Nothing
            qElementColl = Nothing
            pFeature = Nothing
            pPoint = Nothing
            oNode = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de remplir le formulaire pour visualiser les erreurs de précision. 
    '''</summary>
    ''' 
    Public Sub RemplirFormulaireErreurPrecision()
        'Déclarer les variables de travail
        Dim qElementTraiter As Structure_Element_Traiter = Nothing  'Objet contenant l'information de l'élément.
        Dim pFeature As IFeature = Nothing          'Interface contenant l'élément.
        Dim oErreur As Structure_Erreur         'Structure contenant une erreur de précision
        Dim oNode As TreeNode = Nothing         'Objet contenant le noeud en erreur
        Dim i As Integer = Nothing              'Compteur

        Try
            'Sortir si le menu n'est pas défini
            If m_MenuEdgeMatch Is Nothing Then Return

            'Vérifier si les erreurs de précision doivent être affichées
            If m_MenuEdgeMatch.chkPrecision.Checked Then
                'Traiter toutes les erreurs de précision
                For i = 1 To m_ErreurFeaturePrecision.Count
                    'Définir l'erreur de précision
                    oErreur = CType(m_ErreurFeaturePrecision.Item(i), Structure_Erreur)
                    'Ajouter l'erreur dans la liste globale
                    m_ErreurFeature.Add(oErreur)

                    'Définir l'élément
                    qElementTraiter = CType(m_ListeElementTraiter(oErreur.ElementA), Structure_Element_Traiter)
                    pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)

                    'Ajouter l'erreur dans le formulaire
                    oNode = m_MenuEdgeMatch.treErreurs.Nodes.Add(i.ToString & ", Erreur de précision : " & oErreur.Description)

                    oNode.Nodes.Add("PointID=" & oErreur.PointA.ID.ToString)
                    oNode.Nodes.Add(oErreur.IdentifiantA & ", OID=" & oErreur.FeatureA.OID.ToString _
                                    & ", Classe=" & oErreur.FeatureA.Class.AliasName)
                Next

                'Afficher toutes l'information de l'erreur
                'm_MenuEdgeMatch.treErreurs.ExpandAll()
            End If

        Catch erreur As Exception
            'Message d'erreur
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            qElementTraiter = Nothing
            pFeature = Nothing
            oErreur = Nothing
            oNode = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de remplir le formulaire pour visualiser les erreurs d'adjacence. 
    '''</summary>
    ''' 
    Public Sub RemplirFormulaireErreurAdjacence()
        'Déclarer les variables de travail
        Dim qElementTraiter As Structure_Element_Traiter = Nothing  'Objet contenant l'information de l'élément.
        Dim pFeature As IFeature = Nothing          'Interface contenant l'élément.
        Dim oErreur As Structure_Erreur         'Structure contenant une erreur d'adjacence
        Dim oNode As TreeNode = Nothing         'Objet contenant le noeud en erreur
        Dim i As Integer = Nothing              'Compteur

        Try
            'Sortir si le menu n'est pas défini
            If m_MenuEdgeMatch Is Nothing Then Return

            'Vérifier si les erreurs de précision doivent être affichées
            If m_MenuEdgeMatch.chkAdjacence.Checked Then
                'Traiter toutes les erreurs d'adjacence
                For i = 1 To m_ErreurFeatureAdjacence.Count
                    'Définir l'erreur d'adjacence
                    oErreur = CType(m_ErreurFeatureAdjacence.Item(i), Structure_Erreur)
                    'Ajouter l'erreur dans la liste globale
                    m_ErreurFeature.Add(oErreur)

                    'Définir l'élément
                    'qElementTraiter = CType(m_ListeElementTraiter(oErreur.ElementA), Structure_Element_Traiter)
                    'pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                    pFeature = oErreur.FeatureA

                    'Ajouter l'erreur dans le formulaire
                    oNode = m_MenuEdgeMatch.treErreurs.Nodes.Add(i.ToString & ", Erreur d'adjacence : " & oErreur.Description)
                    'Ajouter la description du point d'adjacence
                    oNode.Nodes.Add("PointID=" & oErreur.PointA.ID.ToString)
                    'Ajouter la description de l'élément
                    oNode.Nodes.Add(oErreur.IdentifiantA & ", OID=" & oErreur.FeatureA.OID.ToString _
                                    & ", Classe=" & oErreur.FeatureA.Class.AliasName)

                    'Vérifier si le point correspondant est trouvé
                    If Not oErreur.PointB Is Nothing Then
                        'Ajouter la description du point d'adjacence
                        oNode.Nodes.Add("PointID=" & oErreur.PointB.ID.ToString)
                        'Vérifier si l'élément B est présent
                        If oErreur.FeatureB IsNot Nothing Then
                            'Définir l'élément
                            'qElementTraiter = CType(m_ListeElementTraiter(oErreur.ElementB), Structure_Element_Traiter)
                            'pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                            pFeature = oErreur.FeatureB
                            'Ajouter la description de l'élément
                            oNode.Nodes.Add(oErreur.IdentifiantB & ", OID=" & oErreur.FeatureB.OID.ToString _
                                            & ", Classe=" & oErreur.FeatureB.Class.AliasName)
                        End If
                    End If
                Next

                'Afficher toutes l'information de l'erreur
                'm_MenuEdgeMatch.treErreurs.ExpandAll()
            End If

        Catch erreur As Exception
            'Message d'erreur
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            qElementTraiter = Nothing
            pFeature = Nothing
            oErreur = Nothing
            oNode = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de remplir le formulaire pour visualiser les erreurs d'attribut. 
    '''</summary>
    ''' 
    Public Sub RemplirFormulaireErreurAttribut()
        'Déclarer les variables de travail
        Dim qElementTraiter As Structure_Element_Traiter = Nothing  'Objet contenant l'information de l'élément.
        Dim pFeature As IFeature = Nothing          'Interface contenant l'élément.
        Dim oErreur As Structure_Erreur         'Structure contenant une erreur d'attribut
        Dim oNode As TreeNode = Nothing         'Objet contenant le noeud en erreur
        Dim i As Integer = Nothing              'Compteur

        Try
            'Sortir si le menu n'est pas défini
            If m_MenuEdgeMatch Is Nothing Then Return

            'Vérifier si les erreurs de précision doivent être affichées
            If m_MenuEdgeMatch.chkAttribut.Checked Then
                'Traiter toutes les erreurs d'attribut
                For i = 1 To m_ErreurFeatureAttribut.Count
                    'Définir l'erreur d'attribut
                    oErreur = CType(m_ErreurFeatureAttribut.Item(i), Structure_Erreur)
                    'Ajouter l'erreur dans la liste globale
                    m_ErreurFeature.Add(oErreur)

                    'Ajouter l'erreur dans le formulaire
                    oNode = m_MenuEdgeMatch.treErreurs.Nodes.Add(i.ToString & ", Erreur d'attribut : " & oErreur.Description)
                    oNode.Nodes.Add("PointID=" & oErreur.PointA.ID.ToString)

                    'Définir l'élément
                    'qElementTraiter = CType(m_ListeElementTraiter(oErreur.ElementA), Structure_Element_Traiter)
                    'pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                    oNode.Nodes.Add(oErreur.IdentifiantA & ", OID=" & oErreur.FeatureA.OID.ToString _
                                    & ", Classe=" & oErreur.FeatureA.Class.AliasName & ", Valeur=" & oErreur.ValeurA)

                    'Définir l'élément
                    'qElementTraiter = CType(m_ListeElementTraiter(oErreur.ElementB), Structure_Element_Traiter)
                    'pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                    oNode.Nodes.Add(oErreur.IdentifiantB & ", OID=" & oErreur.FeatureB.OID.ToString _
                                    & ", Classe=" & oErreur.FeatureB.Class.AliasName & ", Valeur=" & oErreur.ValeurB)
                Next

                'Afficher toutes l'information de l'erreur
                'm_MenuEdgeMatch.treErreurs.ExpandAll()
            End If

        Catch erreur As Exception
            'Message d'erreur
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            qElementTraiter = Nothing
            pFeature = Nothing
            oErreur = Nothing
            oNode = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine permet d'initialiser le formulaire.
    '''</summary>
    ''' 
    Public Sub InitialiserFormulaire()
        'Définir les variables de travail
        Dim qMapLayer As New clsGererMapLayer(m_MxDocument.FocusMap)

        Try
            'Écrire une trace de début
            'Call citsMod_FichierTrace.EcrireMessageTrace("Début")

            'Initialiser les variables globales
            m_ListePointAdjacence = Nothing
            m_ListeElementTraiter = Nothing
            m_ListeElementPointAdjacent = Nothing
            m_ListeErreurPointAdjacent = Nothing
            m_ErreurFeature = Nothing
            m_ErreurFeaturePrecision = Nothing
            m_ErreurFeatureAdjacence = Nothing
            m_ErreurFeatureAttribut = Nothing
            m_FeatureLayerDecoupage = Nothing
            m_IdentifiantDecoupage = Nothing
            m_LimiteDecoupage = Nothing
            m_ListePointAdjacence = Nothing
            m_AttributAdjacence = Nothing

            'Initialiser les listbox d'adjacence
            lstAttributAdjacence.Items.Clear()
            lstClasses.Items.Clear()
            cboClasseDecoupage.Items.Clear()
            cboClasseDecoupage.Text = ""
            cboIdentifiantDecoupage.Items.Clear()
            cboIdentifiantDecoupage.Text = ""
            trePoints.Nodes.Clear()
            treErreurs.Nodes.Clear()

            'Définir les variables par défaut
            Call DefinirVariablesDefaut()

            'Définir le MapLayer et tous les FeatureLayer visibles
            Call InitialiserClasses()

            'Initialiser les tolérances
            Call InitialiserTolerance()

            'Afficher la page des classes
            m_MenuEdgeMatch.tabEdgeMatch.SelectedIndex = 0

            'Initialiser le nombre de points d.Adjacence
            lblNbPointAdjacence.Text = "Nombre de points d'Adjacence: 0"
            'Initialiser le nombre d'erreurs
            lblNbErreur.Text = "Nombre total d'erreurs: 0"

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Écrire une trace de fin
            'Call citsMod_FichierTrace.EcrireMessageTrace("Fin")
        End Try
    End Sub

    '''<summary>
    ''' Permet d'initialiser les AddHandler.
    '''</summary>
    '''
    Public Sub InitHandler()
        Try
            'Détruire les Handler
            Call DeleteHandler()

            'Définition des événements pour initialiser le menu
            'm_DocumentEventsOpenDocument = New IDocumentEvents_OpenDocumentEventHandler(AddressOf OnOpenDocument)
            'AddHandler CType(m_MxDocument, IDocumentEvents_Event).OpenDocument, m_DocumentEventsOpenDocument
            AddHandler CType(m_MxDocument, IDocumentEvents_Event).OpenDocument, AddressOf OnOpenDocument

            'm_ActiveViewEventsContentsChanged = New IActiveViewEvents_ContentsChangedEventHandler(AddressOf OnContentsChanged)
            'AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ContentsChanged, m_ActiveViewEventsContentsChanged
            'AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ContentsChanged, AddressOf OnContentsChanged

            'm_ActiveViewEventsSpatialReferenceChanged = New IActiveViewEvents_SpatialReferenceChangedEventHandler(AddressOf OnSpatialReferenceChanged)
            'AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).SpatialReferenceChanged, m_ActiveViewEventsSpatialReferenceChanged
            AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).SpatialReferenceChanged, AddressOf OnSpatialReferenceChanged

            'm_ActiveViewEventsItemAdded = New IActiveViewEvents_ItemAddedEventHandler(AddressOf OnItemAdded)
            'AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemAdded, m_ActiveViewEventsItemAdded
            AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemAdded, AddressOf OnItemAdded

            'm_ActiveViewEventsItemDeleted = New IActiveViewEvents_ItemDeletedEventHandler(AddressOf OnItemDeleted)
            'AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemDeleted, m_ActiveViewEventsItemDeleted
            AddHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemDeleted, AddressOf OnItemDeleted

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    '''<summary>
    ''' Permet de détruire les AddHandler.
    '''</summary>
    '''
    Public Sub DeleteHandler()
        Try
            'Step 4: Dynamically remove an event handler.
            'RemoveHandler CType(m_MxDocument, IDocumentEvents_Event).OpenDocument, m_DocumentEventsOpenDocument
            'RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ContentsChanged, m_ActiveViewEventsContentsChanged
            'RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).SpatialReferenceChanged, m_ActiveViewEventsSpatialReferenceChanged
            'RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemAdded, m_ActiveViewEventsItemAdded
            'RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemDeleted, m_ActiveViewEventsItemDeleted

            RemoveHandler CType(m_MxDocument, IDocumentEvents_Event).OpenDocument, AddressOf OnOpenDocument
            'RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ContentsChanged, m_ActiveViewEventsContentsChanged
            RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).SpatialReferenceChanged, AddressOf OnSpatialReferenceChanged
            RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemAdded, AddressOf OnItemAdded
            RemoveHandler CType(m_MxDocument.FocusMap, IActiveViewEvents_Event).ItemDeleted, AddressOf OnItemDeleted

        Catch erreur As Exception
            'Message d'erreur
            'MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub
#End Region

#Region "Routines et fonctions privées"
    '''<summary>
    ''' Routine permet de définir les variables par défaut.
    '''</summary>
    '''
    Private Sub DefinirVariablesDefaut()
        Try
            'Écrire une trace de début
            'Call citsMod_FichierTrace.EcrireMessageTrace("Début")

            'Définir la symbologie
            Call InitSymbole()

            'Définir la tolérance d'adjacence en mètre
            m_TolAdjacence = 3.0
            m_TolAdjacenceOri = 3.0
            txtTolAdjacence.Text = m_TolAdjacenceOri.ToString

            'Définir la tolérance de recherche en mètre
            m_TolRecherche = 1.0
            m_TolRechercheOri = 1.0
            txtTolRecherche.Text = m_TolRechercheOri.ToString

            'Définir la précision
            m_Precision = 0.001
            txtPrecision.Text = m_Precision.ToString

            'Par défaut, l'adjacence unique est inactive
            m_AdjacenceUnique = False
            chkAdjacenceUnique.Checked = m_AdjacenceUnique

            'Indiquer si les classes d'adjacence ne peuvent être différentes par défaut
            m_ClasseDifferente = False
            chkClasseDifferente.Checked = m_ClasseDifferente

            'Indiquer si les identifiants peuvent être pareils aux points d'adjacence par défaut
            m_SansIdentifiant = False
            chkIdentifiantPareil.Checked = m_SansIdentifiant

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Écrire une trace de fin
            'Call citsMod_FichierTrace.EcrireMessageTrace("Fin")
        End Try
    End Sub

    '''<summary>
    ''' Routine permet de remplir un ComboBox à partir de FeatureLayers visibles.
    '''</summary>
    '''
    '''<param name="qComboBox">ComboBox à remplir.</param>
    '''<param name="sTexte">Texte utilisé pour trouver le FeatureLayer par défaut.</param>
    '''
    ''' <remarks>
    ''' Si le texte est absent, le dernier de la liste sera utilisé comme celui par défaut.
    '''</remarks>
    Private Sub RemplirComboBox(ByRef qComboBox As ComboBox, ByVal sTexte As String, ByVal pEsriGeometryType As esriGeometryType)
        'Déclarer la variables de travail
        Dim qFeatureLayerColl As Collection = Nothing   'Contient la liste des FeatureLayer visibles
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant une classe de données
        Dim sNom As String = Nothing                    'Nom du FeatureLayer sélectionné
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Initialiser le nom du FeatureLayer
            sNom = qComboBox.Text

            'Initialiser le ComboBox
            qComboBox.Items.Clear()

            'Définir la liste des FeatureLayer
            qFeatureLayerColl = m_MapLayer.FeatureLayerCollection

            'vérifier si les FeatureLayers sont présents
            If qFeatureLayerColl IsNot Nothing Then
                'Traiter tous les FeatureLayer
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    'Vérifier le type de géométrie est valide
                    If pEsriGeometryType = esriGeometryType.esriGeometryAny _
                    Or pEsriGeometryType = pFeatureLayer.FeatureClass.ShapeType Then
                        'Ajouter le FeatureLayer dans le ComboBox
                        qComboBox.Items.Add(pFeatureLayer.Name)

                        'Vérifier si le nom spécifié correspond
                        If sNom = pFeatureLayer.Name Then
                            'Définir la valeur par défaut
                            qComboBox.Text = pFeatureLayer.Name

                            'Vérifier la présence du texte recherché pour la valeur par défaut
                        ElseIf InStr(pFeatureLayer.Name.ToUpper, sTexte.ToUpper) > 0 Then
                            'Définir la valeur par défaut
                            qComboBox.Text = pFeatureLayer.Name
                        End If
                    End If
                Next
            End If

            ''Vérifier si aucun FeatureLayer n'a été trouvé
            'If qComboBox.Text.Length = 0 And qComboBox.Items.Count > 0 Then
            '    'Définir le dernier présent dans le ComboxBox
            '    qComboBox.Text = qComboBox.Items(qComboBox.Items.Count - 1).ToString
            'End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
            qFeatureLayerColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine permet de remplir les ListBox des classes et des attributs à partir de FeatureLayers visibles.
    '''</summary>
    '''
    '''<param name="qClasses">ListBox contenant les classes à remplir.</param>
    '''<param name="lstAttributAdjacence">ListBox contenant les attributs d'adjacence à remplir.</param>
    '''<param name="sSelectionnerAttribut">Texte utilisé pour sélectionner les attributs.</param>
    '''<param name="sExclureClasse">Texte utilisé pour exclure des FeatureLayer.</param>
    '''<param name="bExclureClasseVide">Texte utilisé pour exclure des FeatureLayer.</param>
    '''
    ''' <remarks>
    ''' "*" pour sélectionner tout, si le texte est absent, aucun ne sera sélectionné.
    '''</remarks>
    Private Sub RemplirListBoxFeatureLayer(ByRef qClasses As ListBox, ByRef lstAttributAdjacence As ListBox, _
                                           ByVal sSelectionnerAttribut As String, ByVal sExclureClasse As String, ByVal bExclureClasseVide As Boolean)
        'Déclarer la variables de travail
        Dim qFeatureLayerColl As Collection = Nothing   'Contient la liste des FeatureLayer visibles
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant une classe de données
        Dim pFields As IFields = Nothing                'Interface contenant les attributs de la classe
        Dim bAjouter As Boolean = Nothing               'Indique si on doit ajouter le FeatureLayer dans le Listbox
        Dim i As Integer = Nothing                      'Compteur
        Dim j As Integer = Nothing                      'Compteur

        Try
            'Initialiser le ComboBox
            qClasses.Items.Clear()

            'Définir la liste des FeatureLayer
            qFeatureLayerColl = m_MapLayer.FeatureLayerCollection

            'Traiter tous les FeatureLayer
            For i = 1 To qFeatureLayerColl.Count
                'Définir le FeatureLayer
                pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)

                'Indiquer qu'on n'ajoute pas le FeatureLayer par défaut
                bAjouter = False
                'Vérifier si on doit exclure les l
                If bExclureClasseVide Then
                    'Vérifier la présence d'éléments dans la classe
                    If pFeatureLayer.FeatureClass.FeatureCount(Nothing) > 0 Then
                        'Indiquer que l'on veut ajouter le FeatureLayer
                        bAjouter = True
                    End If
                Else
                    'Indiquer que l'on veut ajouter le FeatureLayer
                    bAjouter = True
                End If

                'Vérifier si on doit ajouter le FeatureLayer
                If bAjouter Then
                    'Vérifier si la classe ne doit pas être excluse
                    If InStr(sExclureClasse, pFeatureLayer.Name) = 0 Then
                        'Rendre la classe sélectionnable
                        pFeatureLayer.Selectable = True
                        'Vérifier si le nom est absent dans la liste
                        If qClasses.FindStringExact(pFeatureLayer.Name) = -1 Then
                            'Ajouter le FeatureLayer dans le ListBox
                            qClasses.Items.Add(pFeatureLayer.Name)

                            'Définir les attributs de la classe
                            pFields = pFeatureLayer.FeatureClass.Fields
                            'Traiter tous les attributs du FeatureLayer
                            For j = 0 To pFeatureLayer.FeatureClass.Fields.FieldCount - 1
                                'Vérifier si l'attribut est modifiable, que ce n'est pas une géométrie
                                'et que ce n'est pas l'attribut de découpage
                                If pFields.Field(j).Editable And pFields.Field(j).Type <> esriFieldType.esriFieldTypeGeometry _
                                And pFields.Field(j).Name <> cboIdentifiantDecoupage.Text Then
                                    'Vérifier si le nom est absent dans la liste
                                    If lstAttributAdjacence.FindStringExact(pFields.Field(j).Name) = -1 Then
                                        'Ajouter le FeatureLayer dans le ListBox
                                        lstAttributAdjacence.Items.Add(pFields.Field(j).Name)
                                        'Vérifier la présence du texte recherché pour la valeur par défaut
                                        If sSelectionnerAttribut = "*" Or InStr(sSelectionnerAttribut, pFields.Field(j).Name) > 0 Then
                                            'Sélectionner l'item
                                            lstAttributAdjacence.SelectedIndex = lstAttributAdjacence.Items.Count - 1
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Next

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
            qFeatureLayerColl = Nothing
            pFields = Nothing
        End Try
    End Sub
#End Region
End Class