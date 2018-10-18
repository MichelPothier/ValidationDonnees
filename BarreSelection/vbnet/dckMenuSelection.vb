Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.esriSystem

''' <summary>
''' Designer class of the dockable window add-in. It contains user interfaces that
''' make up the dockable window.
''' </summary>
Public Class dckMenuSelection
    'Variables de travail
    Private m_hook As Object

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

        'Conserver le menu en mémoire
        m_MenuParametresSelection = Me

        'Initialiser le menu
        Me.Init()
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

        Private m_windowUI As dckMenuSelection

        Protected Overrides Function OnCreateChild() As System.IntPtr
            m_windowUI = New dckMenuSelection(Me.Hook)
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

#Region "Routines et fonctions d'événements"
    Private Sub btnInitialiser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInitialiser.Click
        Try
            'Initialiser les paramètres de sélection
            Me.Init()

            'Initialiser les AddHandler
            Me.InitHandler()

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub txtPrecision_GotFocus(sender As Object, e As EventArgs) Handles txtPrecision.GotFocus
        'Déclarer les variables de travail
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance XY de la référence spatiale.

        Try
            'Interface pour extraire la précision de la référence spatiale
            pSpatialRefTol = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceTolerance)
            'Extraire la précision de la référence spatiale
            m_Precision = pSpatialRefTol.XYTolerance
            'Afficher la précision de la référence spatiale
            txtPrecision.Text = m_Precision.ToString("0.0#######")

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pSpatialRefTol = Nothing
        End Try
    End Sub

    Private Sub txtPrecision_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecision.TextChanged
        'Déclarer les variables de travail
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance XY de la référence spatiale.

        Try
            'Vérifier si la précision est numérique
            If IsNumeric(txtPrecision.Text) Then
                'Vérifier si la précision n'est pas négative
                If CDbl(txtPrecision.Text) > 0 Then
                    'Vérifier si la référence spatiale est définie
                    If m_MxDocument.FocusMap.SpatialReference IsNot Nothing Then
                        'Interface pour extraire la précision de la référence spatiale
                        pSpatialRefTol = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceTolerance)
                        'Définir la précision de la référence spatiale
                        pSpatialRefTol.XYTolerance = CDbl(txtPrecision.Text)
                        'Valider la précision de la référence spatiale
                        If pSpatialRefTol.XYToleranceValid = esriSRToleranceEnum.esriSRToleranceOK Then
                            'Définir la tolérance de précision
                            m_Precision = CDbl(txtPrecision.Text)
                        Else
                            'Remettre l'ancienne valeur
                            txtPrecision.Text = m_Precision.ToString("0.0#######")
                        End If
                    End If
                Else
                    'Remettre l'ancienne valeur
                    txtPrecision.Text = m_Precision.ToString("0.0#######")
                End If
            Else
                'Remettre l'ancienne valeur
                txtPrecision.Text = m_Precision.ToString("0.0#######")
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pSpatialRefTol = Nothing
        End Try
    End Sub

    Private Sub cboLayerDecoupage_GotFocus(sender As Object, e As EventArgs) Handles cboLayerDecoupage.GotFocus
        Try
            'Remplir le ComboBox du Layer de découpage
            Call RemplirComboBox(cboLayerDecoupage, cboLayerDecoupage.Text, esriGeometryType.esriGeometryPolygon)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub cboLayerDecoupage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLayerDecoupage.SelectedIndexChanged
        Try
            'Extraire le FeatureLayer de découpage
            m_FeatureLayerDecoupage = m_MapLayer.ExtraireFeatureLayerByName(cboLayerDecoupage.Text, True)

            'vérifier si le FeatureLayer de découpage est valide
            If m_FeatureLayerDecoupage IsNot Nothing Then
                'Déclarer les variables de travail
                Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance XY de la référence spatiale.
                'Interface pour extraire la précision de la référence spatiale
                pSpatialRefTol = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceTolerance)
                'Extraire la précision de la référence spatiale
                m_Precision = pSpatialRefTol.XYTolerance
                'Afficher la précision de la référence spatiale
                txtPrecision.Text = m_Precision.ToString("0.0#######")
                'Vider la mémoire
                pSpatialRefTol = Nothing
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub dckMenuParametresSelection_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        'Déclarer les variables de travail
        Dim iDeltaHeight As Integer
        Dim iDeltaWidth As Integer

        Try
            'Calculer les deltas
            iDeltaHeight = Me.Height - m_Height
            iDeltaWidth = Me.Width - m_Width

            'Vérifier si le menu est initialisé
            If m_MenuParametresSelection IsNot Nothing Then
                'Vérifier si la hauteur du menu est supérieure à la limite
                If m_MenuParametresSelection.Size.Height > 100 Then
                    'Redimensionner les objets du menu en hauteur
                    btnInitialiser.Top = btnInitialiser.Top + iDeltaHeight
                    tabMenuSelection.Height = tabMenuSelection.Height + iDeltaHeight
                    'Initialiser les variables de dimension
                    m_Height = Me.Height
                End If
                'Vérifier si la largeur du menu est supérieure à la limite
                If m_MenuParametresSelection.Size.Width > 100 Then
                    'Redimensionner les objets du menu en largeur
                    tabMenuSelection.Width = tabMenuSelection.Width + iDeltaWidth
                    cboLayerDecoupage.Width = cboLayerDecoupage.Width + iDeltaWidth
                    'Initialiser les variables de dimension
                    m_Width = Me.Width
                End If
            End If

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub chkZoomGeometrieErreur_CheckedChanged(sender As Object, e As EventArgs) Handles chkZoomGeometrieErreur.CheckedChanged
        'Indiquer si on doit effectuer un Zoom selon les géométries d'erreurs
        m_ZoomGeometrieErreur = chkZoomGeometrieErreur.Checked
    End Sub

    Private Sub chkCreerClasseErreur_CheckedChanged(sender As Object, e As EventArgs) Handles chkCreerClasseErreur.CheckedChanged
        'Indiquer si on doit créer la classe d'erreurs en mémoire
        m_CreerClasseErreur = chkCreerClasseErreur.Checked
    End Sub

    Private Sub chkAfficherTable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAfficherTable.CheckedChanged
        'Vérifier si on doit afficher la table d'erreur
        If chkAfficherTable.Checked Then
            'Afficher la table d'erreur
            m_AfficherTableErreur = True
        Else
            'Ne pas afficher la table d'erreur
            m_AfficherTableErreur = False
        End If
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

            'Permet d'initialiser le ComboBox
            m_cboFeatureLayer.RemplirComboBox()
            'Initialiser le menu des paramètres de sélection
            m_MenuParametresSelection.Init()

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser le menu lorsqu'un nouveau document est ouvert.
    '''</summary>
    '''
    Private Sub OnContentsChanged()
        Try
            'Permet d'initialiser le ComboBox
            m_cboFeatureLayer.RemplirComboBox()
            'Initialiser le menu des paramètres de sélection
            m_MenuParametresSelection.Init()

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser les tolérances lors d'un changement de référence spatiale.
    '''</summary>
    '''
    Private Sub OnSpatialReferenceChanged()
        Try
            'Initialiser le menu des paramètres de sélection
            m_MenuParametresSelection.Init()

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser le menu lorsqu'un nouvel item est ajouter dans la Map active.
    '''</summary>
    '''
    Private Sub OnItemAdded(ByVal Item As Object)
        Try
            'Permet d'initialiser le ComboBox
            m_cboFeatureLayer.RemplirComboBox()
            'Initialiser le menu des paramètres de sélection
            m_MenuParametresSelection.Init()

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de réinitialiser le menu lorsqu'un item est retiré dans la Map active.
    '''</summary>
    '''
    Private Sub OnItemDeleted(ByVal Item As Object)
        Try
            'Permet d'initialiser le ComboBox
            m_cboFeatureLayer.RemplirComboBox()
            'Initialiser le menu des paramètres de sélection
            m_MenuParametresSelection.Init()

        Catch erreur As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub dckMenuSelection_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        'Déclarer les variables de travail
        Dim DockWindow As ESRI.ArcGIS.Framework.IDockableWindow
        Dim windowID As UID = New UIDClass

        Try
            'Créer un nouveau menu
            windowID.Value = "MPO_BarreSelection_dckMenuSelection"
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
#End Region

#Region "Routines et fonctions publiques"
    '''<summary>
    ''' Permet d'initialiser le menu pour l'affichage et la modification des paramètres de la barre de sélection.
    '''</summary>
    '''
    Public Sub Init()
        'Déclarer les variables de travail
        Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.

        Try
            If m_MxDocument.FocusMap Is Nothing Then Exit Sub

            'Interface contenant la tolrance de précision de la Référence spatiale
            pSRTolerance = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceTolerance)

            'vérifier si la référence spatiale est spécifiée
            If pSRTolerance Is Nothing Then
                'Définir la précision par défaut
                txtPrecision.Text = m_Precision.ToString
            Else
                'Définir la précision selon la référence spatiale par défaut
                txtPrecision.Text = pSRTolerance.XYTolerance.ToString("0.0#######")
            End If

            'Aucun Zoom selon les géométries d'erreurs par défaut
            m_ZoomGeometrieErreur = chkZoomGeometrieErreur.Checked

            'Créer une classe d'erreur en mémoire par défaut
            chkCreerClasseErreur.Checked = True

            'Afficher la table d'erreur par défaut
            chkAfficherTable.Checked = True

            'Remplir le ComboBox du Layer de découpage
            Call RemplirComboBox(cboLayerDecoupage, "decoupage", esriGeometryType.esriGeometryPolygon)

            'Rafraîchir le menu
            Me.Refresh()

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pSRTolerance = Nothing
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
    Public Sub RemplirComboBox(ByRef qComboBox As ComboBox, ByVal sTexte As String, ByVal pEsriGeometryType As esriGeometryType)
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
            qComboBox.Text = ""

            'Définir l'objet pour extraire les FeatureLayer
            m_MapLayer = New clsGererMapLayer(m_MxDocument.FocusMap, True)

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
                        ElseIf sNom.Length = 0 And sTexte.Length > 0 And pFeatureLayer.Name.ToUpper.Contains(sTexte.ToUpper) Then
                            'Définir la valeur par défaut
                            qComboBox.Text = pFeatureLayer.Name
                        End If
                    End If
                Next
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
            qFeatureLayerColl = Nothing
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
End Class