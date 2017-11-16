Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.DataSourcesRaster

'**
'Nom de la composante : modBarreSelection.vb 
'
'''<summary>
''' Librairies de routines contenant toutes les variables, routines et fonctions globales.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 14 avril 2015
'''</remarks>
''' 
Module modBarreSelection
    'Liste des variables publiques utilisées
    '''<summary> Classe contenant le menu des contraintes d'intégrité. </summary>
    Public m_MenuContrainteIntegrite As dckMenuContrainteIntegrite
    '''<summary> Classe contenant le menu des paramètres de la barre de sélection. </summary>
    Public m_MenuParametresSelection As dckMenuSelection
    ''' <summary>ComboBox utilisé pour gérer le Layer de sélection.</summary>
    Public m_cboFeatureLayer As cboFeatureLayer = Nothing
    ''' <summary>ComboBox utilisé pour gérer les valeurs des paramètres de la contraite de sélection utilisée.</summary>
    Public m_cboParametres As cboParametres = Nothing
    ''' <summary>ComboBox utilisé pour gérer le nom de l'attribut utilisé pour la sélection de groupe.</summary>
    Public m_cboAttributGroupe As cboAttributGroupe = Nothing
    '''<summary> Interface ESRI contenant les FeatureClass à valider.</summary>
    Public m_Geodatabase As IWorkspace = Nothing
    '''<summary> Interface ESRI contenant les contraintes d'intégrités spatiales.</summary>
    Public m_TableContraintes As IStandaloneTable = Nothing
    ''' <summary>Interface ESRI contenant la topologie des éléments visibles.</summary>
    Public m_TopologyGraph As ITopologyGraph4 = Nothing
    ''' <summary>Indique si on doit créer la classe d'erreurs.</summary>
    Public m_CreerClasseErreur As Boolean = True
    ''' <summary>Indique si on doit afficher la table d'erreurs.</summary>
    Public m_AfficherTableErreur As Boolean = True
    ''' <summary>Indique si on doit effectuer un Zoom selon les géométries d'erreurs.</summary>
    Public m_ZoomGeometrieErreur As Boolean = True
    ''' <summary>Interface utilisé pour arrêter un traitement en progression.</summary>
    Public m_TrackCancel As ITrackCancel = Nothing
    ''' <summary>Interface ESRI contenant l'application ArcMap.</summary>
    Public m_Application As IApplication = Nothing
    ''' <summary>Interface ESRI contenant le document ArcMap.</summary>
    Public m_MxDocument As IMxDocument = Nothing
    ''' <summary>Interface ESRI contenant le FeatureLayer de sélection.</summary>
    Public m_FeatureLayer As IFeatureLayer = Nothing
    ''' <summary>Interface ESRI contenant le FeatureLayer de découpage.</summary>
    Public m_FeatureLayerDecoupage As IFeatureLayer = Nothing
    ''' <summary>Indique la requête de sélection utilisée.</summary>
    Public m_Requete As clsRequete = Nothing
    ''' <summary>Indique les valeurs des paramètres de la requête de sélection utilisée.</summary>
    Public m_Parametres As String = Nothing
    ''' <summary>Indique la valeur de la tolérance de précision par défaut.</summary>
    Public m_Precision As Double = 0.001
    ''' <summary>Indique le type de sélection utilisée.</summary>
    Public m_TypeSelection As String = Nothing
    '''<summary> Interface ESRI contenant les géométries sélectionnées.</summary>
    Public m_GeometrieSelection As IGeometry = Nothing
    ''' <summary>Objet qui permet la gestion des FeatureLayer.</summary>
    Public m_MapLayer As clsGererMapLayer = Nothing

    '''<summary>Valeur initiale de la dimension en hauteur du menu.</summary>
    Public m_Height As Integer = 300
    '''<summary>Valeur initiale de la dimension en largeur du menu.</summary>
    Public m_Width As Integer = 300

    '''<summary> Interface ESRI contenant le symbol pour les sommets d'une géométrie.</summary>
    Dim gpSymboleSommet As ISymbol = Nothing
    '''<summary> Interface ESRI contenant le symbol pour la géométrie de type point.</summary>
    Dim gpSymbolePoint As ISymbol = Nothing
    '''<summary> Interface ESRI contenant le symbol pour la géométrie de type ligne.</summary>
    Dim gpSymboleLigne As ISymbol = Nothing
    '''<summary> Interface ESRI contenant le symbol pour la géométrie de type surface.</summary>
    Dim gpSymboleSurface As ISymbol = Nothing
    '''<summary>Interface ESRI contenant le symbol de texte par défaut.</summary>
    Dim gpSymboleText As ITextSymbol = Nothing

    ' '''<summary>Interface ESRI qui permet de gérer l'événement lorsque qu'un nouveau document est ouvert.</summary>
    'Public m_DocumentEventsOpenDocument As IDocumentEvents_OpenDocumentEventHandler
    ' ''''<summary>Interface ESRI qui permet de gérer l'événement lorsqu'un nouveau document est activé.</summary>
    'Public m_ActiveViewEventsContentsChanged As IActiveViewEvents_ContentsChangedEventHandler
    ' '''<summary>Interface ESRI qui permet de gérer l'événement lors d'un changement de référence spatiale à la Map active.</summary>
    'Public m_ActiveViewEventsSpatialReferenceChanged As IActiveViewEvents_SpatialReferenceChangedEventHandler
    ' '''<summary>Interface ESRI qui permet de gérer l'événement lorsqu'un nouvel item est ajouté à la Map active.</summary>
    'Public m_ActiveViewEventsItemAdded As IActiveViewEvents_ItemAddedEventHandler
    ' '''<summary>Interface ESRI qui permet de gérer l'événement lorsqu'un nouvel item est retiré de la Map active</summary>
    'Public m_ActiveViewEventsItemDeleted As IActiveViewEvents_ItemDeletedEventHandler

    '''<summary>
    '''Fonction qui permet de retourner l'interface d'édition de ArcMap.
    '''</summary>
    '''
    '''<returns>"IEditor" contenant l'interface d'édition des données, "Nothing" sinon.</returns>
    '''
    Public Function GetEditor() As IEditor
        'Déclarer les variables de travail
        Dim qType As Type = Nothing                 'Interface contenant le type d'objet à créer
        Dim oObjet As System.Object = Nothing       'Interface contenant l'objet correspondant à l'application
        Dim pApplication As IApplication = Nothing  'Interface contenant l'application
        Dim pEditor As IEditor = Nothing            'Interface ESRI utilisée pour effectuer l'édition.

        'Définir la valeur de retour par défaut
        GetEditor = Nothing

        Try
            'Définir l'application pour extraire l'extension de la topologie
            qType = Type.GetTypeFromProgID("esriFramework.AppRef")
            oObjet = Activator.CreateInstance(qType)
            pApplication = CType(oObjet, IApplication)

            'Interface pour vérifer si on est en mode édition
            pEditor = CType(pApplication.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Retourner l'interface pour l'édition
            GetEditor = pEditor

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            qType = Nothing
            oObjet = Nothing
            pApplication = Nothing
            pEditor = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de créer et retourner la topologie en mémoire des éléments entre une collection de FeatureLayer.
    '''</summary>
    '''
    '''<param name="pEnvelope">Interface ESRI contenant l'enveloppe de création de la topologie traitée.</param>
    '''<param name="qFeatureLayerColl">Interface ESRI contenant les FeatureLayers pour traiter la topologie.</param> 
    '''<param name="dTolerance">Tolerance de proximité.</param> 
    '''
    '''<returns>"ITopologyGraph" contenant la topologie des classes de données, "Nothing" sinon.</returns>
    '''
    Public Function CreerTopologyGraph2(ByVal pEnvelope As IEnvelope, ByVal qFeatureLayerColl As Collection, ByVal dTolerance As Double) As ITopologyGraph4
        'Déclarer les variables de travail
        Dim qType As Type = Nothing                         'Interface contenant le type d'objet à créer.
        Dim oObjet As System.Object = Nothing               'Interface contenant l'objet correspondant à l'application.
        Dim pApplication As IApplication = Nothing          'Interface contenant l'application.
        Dim pTopologyUID As UID = New UIDClass()            'Interface contenant l'identifiant de l'extension de la topologie.
        Dim pTopologyExt As ITopologyExtension = Nothing    'Interface contenant l'extension de la topologie.
        Dim pMapTopology As IMapTopology2 = Nothing         'Interface utilisé pour créer la topologie.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la classe de données.

        'Définir la valeur de retour par défaut
        CreerTopologyGraph2 = Nothing

        Try
            'Définir l'application pour extraire l'extension de la topologie
            qType = Type.GetTypeFromProgID("esriFramework.AppRef")
            oObjet = Activator.CreateInstance(qType)
            pApplication = CType(oObjet, IApplication)

            'Définir l'extension de topologie
            pTopologyUID.Value = "esriEditorExt.TopologyExtension"
            pTopologyExt = CType(pApplication.FindExtensionByCLSID(pTopologyUID), ITopologyExtension)

            'Définir l'interface pour créer la topologie
            pMapTopology = CType(pTopologyExt.MapTopology, IMapTopology2)

            'S'assurer que laliste des Layers est vide
            pMapTopology.ClearLayers()

            'Traiter tous les FeatureLayer présents
            For Each pFeatureLayer In qFeatureLayerColl
                'Ajouter le FeatureLayer à la topologie
                pMapTopology.AddLayer(pFeatureLayer)
            Next

            'Changer la référence spatiale selon l'enveloppe
            pMapTopology.SpatialReference = pEnvelope.SpatialReference

            'Définir la tolérance de connexion et de partage
            pMapTopology.ClusterTolerance = dTolerance

            'Interface pour construre la topologie
            pTopologyGraph = CType(pMapTopology.Cache, ITopologyGraph4)
            pTopologyGraph.SetEmpty()

            Try
                'Construire la topologie
                pTopologyGraph.Build(pEnvelope, False)
            Catch ex As Exception
                'Retourner une erruer de création de la topologie
                Err.Raise(-1, , "Incapable de créer la topologie, méomoire insuffisante!")
                pTopologyGraph = Nothing
                GC.Collect()
            End Try

            'Retourner la topologie
            CreerTopologyGraph2 = pTopologyGraph

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            qType = Nothing
            oObjet = Nothing
            pApplication = Nothing
            pTopologyUID = Nothing
            pTopologyExt = Nothing
            pTopologyGraph = Nothing
            pMapTopology = Nothing
            pFeatureLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de trouvée tout le réseau rattaché à un Node de départ sélectionné.
    ''' 
    ''' Un réseau est un ensemble de lignes interconnecté entre eux.
    '''</summary>
    '''
    '''<param name="pPolyline"> Interface ESRI contenant le réseau trouvé.</param>
    '''<param name="pTopologyGraph"> Interface ESRI contenant la topologie des éléments visibles.</param>
    '''<param name="pTopoNode"> Interface ESRI contenant le Node de la topologie utilisé pour rechercher le réseau.</param>
    '''<param name="bVerifierSelection"> Permettre d'indiquer si on doit vérifier si le Node est déjà sélectionné.</param>
    ''' 
    Public Sub SelectionnerReseau(ByRef pPolyline As IPolyline, ByRef pTopologyGraph As ITopologyGraph4, ByRef pTopoNode As ITopologyNode, _
                                  Optional ByVal bVerifierSelection As Boolean = True)
        'Déclarer les variables de travail
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface pour extraire les Nodes du Edge traité.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le Edge traité.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter la géométrie du Edge dans la Poilyline.

        Try
            'Vérifier si le Node a été traité
            If Not pTopoNode.IsSelected Or bVerifierSelection = False Then
                'Interface pour ajouter les lignes trouvées
                pGeomColl = CType(pPolyline, IGeometryCollection)

                'Indiquer que le Node est traité
                pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoNode)

                'Interface pour extraire tous les edges du Node
                pEnumNodeEdge = pTopoNode.Edges(True)

                'Extraire le premier edge du Node
                pEnumNodeEdge.Reset()
                pEnumNodeEdge.Next(pTopoEdge, True)

                'Traiter tous les Edges du Node
                Do Until pTopoEdge Is Nothing
                    'Vérifier si le Edge a été traité
                    If Not pTopoEdge.IsSelected Then
                        'Ajouter le Edge dans la Polyline
                        pGeomColl.AddGeometryCollection(CType(pTopoEdge.Geometry, IGeometryCollection))

                        'Indiquer que le Node est traité
                        pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoEdge)

                        'Sélectionner le réseau par le FromNode
                        SelectionnerReseau(pPolyline, pTopologyGraph, pTopoEdge.FromNode)

                        'Sélectionner le réseau par le ToNode
                        SelectionnerReseau(pPolyline, pTopologyGraph, pTopoEdge.ToNode)
                    End If

                    'Extraire le prochain edge du Node
                    pEnumNodeEdge.Next(pTopoEdge, True)
                Loop
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pEnumNodeEdge = Nothing
            pTopoEdge = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de créer une nouvelle fenêtre de données (MapFrame et Map) et d'ajouter tous les FeatureLayers dans la Map.
    '''
    ''' Ce qui est important de comprendre lorsqu'on est dans un projet, c'est qu'il y a deux modes pour
    ''' visualiser les données,DATA VIEW et PAGE LAYOUT. Dans le mode DataView, la fenêtre de données
    ''' correspond au MAP et dans le mode PageLayout, la fenêtre de données correspond au MapFrame. Lorsque
    ''' l'on veut créer une nouvelle fenêtre de données, il faut créer une nouvelle Map dans le mode
    ''' DataView et un nouveau MapFrame dans le mode PageLayout.
    '''</summary>
    '''
    '''<param name="sNomFenetre">Nom de la fenêtre de données à créer.</param>
    '''<param name="qFeatureLayerColl">Collection des FeatureLayer en relation utilisés pour extraire les éléments en relation.</param>
    ''' 
    '''<returns>"IMapFrame" vide avec des propriétés par défaut, "Nothing" sinon</returns>
    '''
    Public Function fpCreerMapFrame(ByVal sNomFenetre As String, ByVal qFeatureLayerColl As Collection) As IMapFrame
        'Déclarer les varibles de travail
        Dim pMap As IMap = Nothing              'Interface ESRI contenant des Layers.
        Dim pMapFrame As IMapFrame = Nothing    'Interface ESRI contenant un MapFrame (Fenêtre de données).
        Dim pElement As IElement = Nothing      'Interface ESRI contenant un objet dans le PageLayout.
        Dim pEnvelope As IEnvelope = Nothing    'Interface ESRI utilisée pour positionner le MapFrame.
        Dim pFeatureLayer As IFeatureLayer = Nothing ' Interface ESRI contenant une classe de données.

        'Définir la valeur de retour par défaut
        fpCreerMapFrame = Nothing

        Try
            'Créer et nommé une nouvelle Fenêtre
            pMap = New Map
            pMap.Name = sNomFenetre

            'Créer un Frame associé à la Fenêtre
            pMapFrame = New MapFrame
            pMapFrame.Map = pMap

            'Positionner le Frame
            pElement = CType(pMapFrame, IElement)
            pEnvelope = New Envelope
            pEnvelope.PutCoords(0, 0, 5, 5)
            pElement.Geometry = pEnvelope

            'Ajouter tous les FeatureLayer à la fenêtre de données
            For Each pFeatureLayer In qFeatureLayerColl
                'Ajouter un FeatureLayer à la fenêtre de données
                pMap.AddLayer(pFeatureLayer)
            Next

            'Retourner la fenêtre de données vide
            fpCreerMapFrame = pMapFrame

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pMap = Nothing
            pMapFrame = Nothing
            pElement = Nothing
            pEnvelope = Nothing
            pFeatureLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Initialiser les couleurs par défaut pour le texte qui sera affiché avec les
    ''' géométries déssiné et initialiser les 3 couleurs pour les géométries de type
    ''' point, ligne et surface.
    '''</summary>
    '''
    Public Sub InitSymbole()
        'Déclarer les variables de travail
        Dim pRgbColor As IRgbColor = Nothing                'Interface ESRI contenant la couleur RGB.
        Dim pTextSymbol As ITextSymbol = Nothing            'Interface ESRI contenant un symbole de texte.
        Dim pMarkerSymbol As ISimpleMarkerSymbol = Nothing  'Interface ESRI contenant un symbole de point.
        Dim pLineSymbol As ISimpleLineSymbol = Nothing      'Interface ESRi contenant un symbole de ligne.
        Dim pFillSymbol As ISimpleFillSymbol = Nothing      'Interface ESRI contenant un symbole de surface.

        'Permet d'initialiser la symbologie
        Try
            'Vérifier si le symbole est invalide
            If gpSymboleSommet Is Nothing Then
                'Définir la couleur rouge pour le polygon
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                pRgbColor.Green = 100
                'Définir la symbologie pour la limite d'un polygone
                pMarkerSymbol = New SimpleMarkerSymbol
                pMarkerSymbol.Color = pRgbColor
                pMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSSquare
                pMarkerSymbol.Size = 3
                'Conserver le symbole
                gpSymboleSommet = CType(pMarkerSymbol, ISymbol)
            End If

            'Vérifier si le symbole est invalide
            If gpSymbolePoint Is Nothing Then
                'Définir la couleur rouge pour le polygon
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                'Définir la symbologie pour la limite d'un polygone
                pMarkerSymbol = New SimpleMarkerSymbol
                pMarkerSymbol.Color = pRgbColor
                pMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCircle
                pMarkerSymbol.Size = 5
                'Conserver le symbole 
                gpSymbolePoint = CType(pMarkerSymbol, ISymbol)
            End If

            'Vérifier si le symbole est invalide
            If gpSymboleLigne Is Nothing Then
                'Définir la couleur rouge pour le polygon
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                'Définir la symbologie pour la limite d'un polygone
                pLineSymbol = New SimpleLineSymbol
                pLineSymbol.Color = pRgbColor
                pLineSymbol.Style = esriSimpleLineStyle.esriSLSDot
                pLineSymbol.Width = 1.5
                'Conserver le symbole en mémoire
                gpSymboleLigne = CType(pLineSymbol, ISymbol)
            End If

            'Vérifier si le symbole est invalide
            If gpSymboleSurface Is Nothing Then
                'Définir la couleur rouge pour le polygon
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                'Définir la symbologie pour la limite d'un polygone
                pLineSymbol = New SimpleLineSymbol
                pLineSymbol.Color = pRgbColor
                'Définir la symbologie pour l'intérieur d'un polygone
                pFillSymbol = New SimpleFillSymbol
                pFillSymbol.Color = pRgbColor
                pFillSymbol.Outline = pLineSymbol
                pFillSymbol.Style = esriSimpleFillStyle.esriSFSBackwardDiagonal
                'Conserver le symbole 
                gpSymboleSurface = CType(pFillSymbol, ISymbol)
            End If

            'Vérifier si le symbole est invalide
            If gpSymboleText Is Nothing Then
                'Définir la couleur Noir pour le texte
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                'Créer le symbole pour le texte
                gpSymboleText = New TextSymbol
                gpSymboleText.Color = pRgbColor
                gpSymboleText.Font.Bold = True
                gpSymboleText.HorizontalAlignment = esriTextHorizontalAlignment.esriTHAFull
                gpSymboleText.Size = 9
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            pRgbColor = Nothing
            pTextSymbol = Nothing
            pMarkerSymbol = Nothing
            pLineSymbol = Nothing
            pFillSymbol = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de dessiner dans la vue active les géométries Point, MultiPoint, 
    ''' Polyline et/ou Polygon. Ces géométries peuvent être contenu dans un GeometryBag. 
    ''' Un point est représenté par un carré, une ligne est représentée par une ligne pleine 
    ''' et la surface est représentée par une ligne pleine pour la limite et des lignes à
    ''' 45 dégrés pour l'intérieur. On peut afficher le numéro de la géométrie pour un GeometryBag.
    '''</summary>
    '''
    '''<param name="pMxDoc"> Interface ESRI utilisé pour dessiner les géométries.</param>
    '''<param name="pGeometry"> Interface ESRI contenant la géométrie èa dessiner.</param>
    '''<param name="bRafraichir"> Indique si on doit rafraîchir la vue active.</param>
    '''<param name="pTrackCancel"> Permettre d'annuler le traitement avec la touche ESC.</param>
    ''' 
    '''<return>Un booleen est retourner pour indiquer la fonction s'est bien exécutée.</return>
    ''' 
    Public Function bDessinerGeometrie(ByRef pMxDoc As IMxDocument, ByRef pTrackCancel As ITrackCancel, ByRef pGeometry As IGeometry, _
                                       Optional ByVal bRafraichir As Boolean = False) As Boolean
        'Déclarer les variables de travail
        Dim pScreenDisplay As IScreenDisplay = Nothing  'Interface ESRI contenant le document de ArcMap.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface ESRI contenant la fenêtre d'affichage.
        Dim pArea As IArea = Nothing                    'Interface ESRI contenant un symbole de texte.
        Dim pGeomTexte As IGeometry = Nothing           'Interface ESRI pour la position du texte d'une surface.
        Dim pRelOp As IRelationalOperator = Nothing     'Interface pour vérifier la relation.

        Try
            'Vérifier si on doit rafraichir l'écran
            If bRafraichir Then
                'Rafraîchier l'affichage
                pMxDoc.ActiveView.Refresh()
                System.Windows.Forms.Application.DoEvents()
            End If

            'Vérifier si la géométrie est absente
            If pGeometry Is Nothing Then Exit Function

            'Vérifier si la géométrie est vide
            If pGeometry.IsEmpty Then Exit Function

            'Interface pour vérifier l'intersection avec la zone d'affichage
            pRelOp = CType(pMxDoc.ActiveView.Extent, IRelationalOperator)

            'Initialiser les variables de travail
            pScreenDisplay = pMxDoc.ActiveView.ScreenDisplay

            'Transformation du système de coordonnées selon la vue active
            pGeometry.Project(pMxDoc.FocusMap.SpatialReference)

            'Vérifier si la géométrie est un GeometryBag
            If pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                'Interface pour traiter toutes les géométries présentes dans le GeometryBag
                pGeomColl = CType(pGeometry, IGeometryCollection)

                'Dessiner toutes les géométrie présentes dans une collectiopMxDocn de géométrie
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Vérifier si la géométrie intersecte
                    'If Not pRelOp.Disjoint(pGeomColl.Geometry(i)) Then
                    'Dessiner la géométrie contenu dans un GeometrieBag
                    Call bDessinerGeometrie(pMxDoc, pTrackCancel, pGeomColl.Geometry(i), False)
                    'Vérifier si la géométrie n'est pas de type point ou multipoint
                    If pGeomColl.Geometry(i).Dimension > esriGeometryDimension.esriGeometry0Dimension Then
                        'Dessiner les sommets de la géométrie contenu dans un GeometrieBag
                        Call bDessinerSommet(pMxDoc, pGeomColl.Geometry(i), False)
                    End If
                    'End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next i

                'Vérifier si la géométrie est un point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(gpSymbolePoint)
                    .DrawPoint(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est un multi-point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(gpSymbolePoint)
                    .DrawMultipoint(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est une ligne
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(gpSymboleLigne)
                    .DrawPolyline(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est une surface
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon _
            Or pGeometry.GeometryType = esriGeometryType.esriGeometryEnvelope Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(gpSymboleSurface)
                    .DrawPolygon(pGeometry)
                    .FinishDrawing()
                End With
            End If

            'Retourner le résultat
            bDessinerGeometrie = True

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            pScreenDisplay = Nothing
            pGeomColl = Nothing
            pArea = Nothing
            pGeomTexte = Nothing
            pRelOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de dessiner dans la vue active les sommets des géométries Point, MultiPoint, Polyline et/ou Polygon.
    ''' Ces géométries peuvent être contenu dans un GeometryBag. Les sommets sont représentés par un cercle.
    '''</summary>
    '''
    '''<param name="pMxDoc"> Interface ESRI utilisé pour dessiner les géométries.</param>
    '''<param name="pGeometry"> Interface ESRI contenant la géométrie utilisée pour dessiner les sommets.</param>
    '''<param name="bRafraichir"> Indique si on doit rafraîchir la vue active.</param>
    ''' 
    '''<return>Un booleen est retourner pour indiquer la fonction s'est bien exécutée.</return>
    ''' 
    Public Function bDessinerSommet(ByRef pMxDoc As IMxDocument, ByVal pGeometry As IGeometry, Optional ByVal bRafraichir As Boolean = False) As Boolean
        'Déclarer les variables de travail
        Dim pScreenDisplay As IScreenDisplay = Nothing  'Interface ESRI contenant le document de ArcMap.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface ESRI contenant la fenêtre d'affichage.
        Dim pMultiPoint As IMultipoint = Nothing        'Interface contenant les sommets de la géométrie
        Dim pPointColl As IPointCollection = Nothing    'Interface utilisée pour transformer la géométrie en multipoint

        Try
            'Vérifier si la géométrie est absente
            If pGeometry Is Nothing Then Exit Function

            'Vérifier si la géométrie est vide
            If pGeometry.IsEmpty Then Exit Function

            'Initialiser les variables de travail
            pScreenDisplay = pMxDoc.ActiveView.ScreenDisplay

            'Vérifier si on doit rafraichir l'écran
            If bRafraichir Then
                'Rafraîchier l'affichage
                pMxDoc.ActiveView.Refresh()
                System.Windows.Forms.Application.DoEvents()
            End If

            'Transformation du système de coordonnées selon la vue active
            pGeometry.Project(pMxDoc.FocusMap.SpatialReference)

            'Vérifier si la géométrie est un GeometryBag
            If pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                'Interface pour traiter toutes les géométries présentes dans le GeometryBag
                pGeomColl = CType(pGeometry, IGeometryCollection)

                'Dessiner toutes les géométrie présentes dans une collection de géométrie
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Dessiner les sommets de la géométrie contenu dans un GeometrieBag
                    Call bDessinerSommet(pMxDoc, pGeomColl.Geometry(i), False)
                Next i

                'Vérifier si la géométrie est un point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(gpSymboleSommet)
                    .DrawPoint(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est un multi-point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(gpSymboleSommet)
                    .DrawMultipoint(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est une ligne ou une surface
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline _
            Or pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Créer un nouveau multipoint vide
                pMultiPoint = CType(New Multipoint, IMultipoint)
                pMultiPoint.SpatialReference = pGeometry.SpatialReference

                'Transformer la géométrie en multipoint
                pPointColl = CType(pMultiPoint, IPointCollection)
                pPointColl.AddPointCollection(CType(pGeometry, IPointCollection))

                'Afficher les sommets de la géométrie avec la symbologie des sommets dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(gpSymboleSommet)
                    .DrawMultipoint(pMultiPoint)
                    .FinishDrawing()
                End With
            End If

            'Retourner le résultat
            bDessinerSommet = True

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            pScreenDisplay = Nothing
            pGeomColl = Nothing
            pMultiPoint = Nothing
            pPointColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de dessiner un pixel en utilisant un point et un texte contenant la valeur du pixel.
    '''</summary>
    '''
    '''<param name="pScreenDisplay">Interface contenant la fenêtre d'affichage.</param>
    '''<param name="pPoint">Interface contenant le Point du centre du pixel.</param>
    '''<param name="sTexte">Contient le texte à afficher et représentant la valeur du pixel.</param>
    ''' 
    Public Sub DessinerPixel(ByVal pScreenDisplay As IScreenDisplay, ByVal pPoint As IPoint, ByVal sTexte As String)
        'Déclarer les variables de travail

        Try
            'Afficher le polygone, le point du centre et le texte dans la vue active
            With pScreenDisplay
                'Débuter l'affichage
                .StartDrawing(pScreenDisplay.hDC, -1)

                'Afficher le texte avec sa symbologie dans la vue active
                .SetSymbol(CType(gpSymboleText, ISymbol))
                .DrawText(pPoint, sTexte)

                'Terminer l'affichage
                .FinishDrawing()
            End With

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet définir et retourner la requête spécifiée.
    '''</summary>
    ''' 
    '''<returns>"intRequete" contenant toute l'information de la requête spécifiée.</returns>
    ''' 
    Public Function DefinirRequete() As intRequete
        'Déclarer la variables de travail
        Dim oRequete As intRequete       'Objet contenant la requête spécifiée.

        'Définir la requête par défaut
        DefinirRequete = Nothing

        Try
            'Définir la requête
            oRequete = m_Requete.Requete
            'Définir l'application
            oRequete.Application = m_Application
            'Définir la Map active
            oRequete.Map = m_MxDocument.FocusMap
            'Définir le FeatureLayer de sélection
            oRequete.FeatureLayerSelection = m_FeatureLayer
            'Définir le FeatureLayer de découpage, l'élément, le polygone et les limites selon les éléments sélectionnés.
            oRequete.DefinirLimiteLayerDecoupage(m_FeatureLayerDecoupage)
            'Définir les paramètres
            oRequete.Parametres = m_Parametres
            'Définir la tolérance de précision
            oRequete.Precision = m_Precision
            'Définir si on doit créer la table d'erreur
            oRequete.CreerClasseErreur = m_CreerClasseErreur
            'Définir si on doit afficher la table d'erreur
            oRequete.AfficherTableErreur = m_AfficherTableErreur
            'Vérifier si la requête doit avoir des FeatureLayer en relation
            If oRequete.FeatureLayersRelation IsNot Nothing Then
                'Définir les FeatureLayers en relation
                oRequete.FeatureLayersRelation = ExtraireSelectedFeatureLayer(m_MxDocument)
            End If
            'Vérifier si la requête doit avoir des RasterLayer en relation
            If oRequete.RasterLayersRelation IsNot Nothing Then
                'Définir les RasterLayers en relation
                oRequete.RasterLayersRelation = ExtraireSelectedRasterLayer(m_MxDocument)
            End If

            'Retourner la requête par défaut
            DefinirRequete = oRequete

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            oRequete = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'extraire la collection des FeatureLayers sélectionnés dans le TableOfContent.
    '''</summary>
    ''' 
    '''<param name="pMxDoc">Interface contenant le document ArcMap.</param>
    ''' 
    '''<returns>"Collection" contenant les FeatureLayer sélectionnés dans le TableOfContent.</returns>
    ''' 
    Public Function ExtraireSelectedFeatureLayer(ByVal pMxDoc As IMxDocument) As Collection
        'Déclarer la variables de travail
        Dim pContentsViewSelection As IContentsViewSelection = Nothing  'Interface contenant les object sélectionnés dans le TableOfContent.
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant le FeatureLayer en relation.
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour sélectionner les éléments.
        Dim pSpatialFilter As ISpatialFilter = Nothing  'Interface contenant la requête spatiale.
        Dim pObject As Object = Nothing                 'Interface contenant un objet.
        Dim pSet As ISet = Nothing                      'Interface contenant les objets sélectionnés.

        'Définir la valeur de retour par défaut
        ExtraireSelectedFeatureLayer = New Collection

        Try
            'Sortir si le document ArcMap est invalide
            If pMxDoc Is Nothing Then Exit Function

            'Définir le tableOfContent
            pContentsViewSelection = CType(pMxDoc.CurrentContentsView, IContentsViewSelection)

            'Extraire les object sélectionnés dans le TableOfContent
            pSet = pContentsViewSelection.SelectedItems()

            'Vérifier si aucun objet sélectionnés
            If pSet Is Nothing Then Exit Function

            'Initialiser l'extraction
            pSet.Reset()

            'Extraire le premier object
            pObject = pSet.Next()

            'Traiter tous les object sélectionnés
            Do Until pObject Is Nothing
                'Vérifier si l'object est un FeatureLayer
                If TypeOf (pObject) Is IFeatureLayer Then
                    'Définir la FeatureLayer
                    pFeatureLayer = CType(pObject, IFeatureLayer)

                    'Vérifier si la FeatureClass est valide
                    If pFeatureLayer.FeatureClass Is Nothing Then
                        'Retourner une erreur
                        Err.Raise(1, , "FeatureClass invalide : " & pFeatureLayer.Name)
                    End If

                    'Interface pour vérifier la sélection
                    pFeatureSel = CType(pFeatureLayer, IFeatureSelection)
                    'Vérifier la sélection
                    If pFeatureSel.SelectionSet.Count = 0 Then
                        'Définir la requête spatiale
                        pSpatialFilter = New SpatialFilter
                        pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                        pSpatialFilter.OutputSpatialReference(pFeatureLayer.FeatureClass.ShapeFieldName) = pMxDoc.FocusMap.SpatialReference
                        pSpatialFilter.Geometry = pMxDoc.ActiveView.Extent
                        'Sélectionner les éléments selon l'affichage
                        pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
                    End If

                    'Ajouter le FeatureLayer dans la collection
                    ExtraireSelectedFeatureLayer.Add(pFeatureLayer, pFeatureLayer.Name)
                End If

                'Extraire le prochain object
                pObject = pSet.Next()
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            'Throw
        Finally
            'Vider la mémoire
            pContentsViewSelection = Nothing
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            pSpatialFilter = Nothing
            pObject = Nothing
            pSet = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'extraire la collection des RasterLayers sélectionnés dans le TableOfContent.
    '''</summary>
    ''' 
    '''<param name="pMxDoc">Interface contenant le document ArcMap.</param>
    ''' 
    '''<returns>"Collection" contenant les RasterLayer sélectionnés dans le TableOfContent.</returns>
    ''' 
    Public Function ExtraireSelectedRasterLayer(ByVal pMxDoc As IMxDocument) As Collection
        'Déclarer la variables de travail
        Dim pContentsViewSelection As IContentsViewSelection = Nothing  'Interface contenant les object sélectionnés dans le TableOfContent.
        Dim pRasterLayer As IRasterLayer = Nothing      'Interface contenant le RasterLayer en relation.
        Dim pObject As Object = Nothing                 'Interface contenant un objet.
        Dim pSet As ISet = Nothing                      'Interface contenant les objets sélectionnés.
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant le Layer d'un RasterCatalog.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface pour extraire les items d'un RasterCatalog.
        Dim pRasterCatalogItem As IRasterCatalogItem = Nothing 'Interface contenant un Item du RasterCatalog.

        'Définir la valeur de retour par défaut
        ExtraireSelectedRasterLayer = New Collection

        Try
            'Sortir si le document ArcMap est invalide
            If pMxDoc Is Nothing Then Exit Function

            'Définir le tableOfContent
            pContentsViewSelection = CType(pMxDoc.CurrentContentsView, IContentsViewSelection)

            'Extraire les object sélectionnés dans le TableOfContent
            pSet = pContentsViewSelection.SelectedItems()

            'Vérifier si aucun objet sélectionnés
            If pSet Is Nothing Then Exit Function

            'Initialiser l'extraction
            pSet.Reset()

            'Extraire le premier object
            pObject = pSet.Next()

            'Traiter tous les object sélectionnés
            Do Until pObject Is Nothing
                'Vérifier si l'object est un RasterLayer
                If TypeOf (pObject) Is IRasterLayer Then
                    'Définir le RasterLayer
                    pRasterLayer = CType(pObject, IRasterLayer)

                    'Vérifier si le Raster est valide
                    If pRasterLayer.Raster Is Nothing Then
                        'Retourner une erreur
                        Err.Raise(1, , "Raster invalide : " & pRasterLayer.Name)
                    End If

                    'Ajouter le Raster dans la collection
                    ExtraireSelectedRasterLayer.Add(pRasterLayer, pRasterLayer.Name)

                    'Si l'object est un GdbRasterCatalogLayer (Catalogue d'images)
                ElseIf TypeOf (pObject) Is IGdbRasterCatalogLayer Then
                    'Interface pour extraire le FeatureCatalog
                    pFeatureLayer = CType(pObject, IFeatureLayer)

                    'Interface pour extraire les Raster
                    pFeatureCursor = pFeatureLayer.Search(Nothing, True)

                    'Extraire le premier item du RasterCatalog
                    pRasterCatalogItem = CType(pFeatureCursor.NextFeature, IRasterCatalogItem)

                    'Traiter tous les Raster
                    Do Until pRasterCatalogItem Is Nothing
                        'Créer un nouveau RasterLayer vide
                        pRasterLayer = New RasterLayer
                        'Ajouter le Raster dans le Layer
                        pRasterLayer.CreateFromDataset(CType(pRasterCatalogItem.RasterDataset, IRasterDataset))
                        'Ajouter le Raster dans la collection
                        ExtraireSelectedRasterLayer.Add(pRasterLayer, pRasterLayer.Name)

                        'Extraire le premier item du RasterCatalog
                        pRasterCatalogItem = CType(pFeatureCursor.NextFeature, IRasterCatalogItem)
                    Loop
                End If

                'Extraire le prochain object
                pObject = pSet.Next()
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            'Throw
        Finally
            'Vider la mémoire
            pContentsViewSelection = Nothing
            pRasterLayer = Nothing
            pObject = Nothing
            pSet = Nothing
            pFeatureLayer = Nothing
            pFeatureCursor = Nothing
            pRasterCatalogItem = Nothing
        End Try
    End Function
End Module
