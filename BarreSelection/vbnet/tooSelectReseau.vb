Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.EditorExt

'**
'Nom de la composante : tooSelectReseau.vb
'
'''<summary>
'''Commande qui permet sélectionner les éléments par liens de conexion avec le premier élément du réseau trouvé.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''</remarks>
''' 
Public Class tooSelectReseau
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    'Déclarer les variable de travail
    Private gpApp As IApplication
    Private gpMxDoc As IMxDocument

    Public Sub New()
        'Définir les variables de travail
        Dim windowID As UID = New UIDClass

        Try
            'Vérifier si l'application est définie
            If Not Hook Is Nothing Then
                'Définir l'application
                gpApp = CType(Hook, IApplication)

                'Vérifier si l'application est ArcMap
                If TypeOf Hook Is IMxApplication Then
                    'Rendre active la commande
                    Enabled = True
                    'Définir le document
                    gpMxDoc = CType(gpApp.Document, IMxDocument)
                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            windowID = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        'Déclarer les variables de travail
        Dim oRectangle As tagRECT                               'Contient le rectangle en pixel.
        Dim pEnvelope As IEnvelope = Nothing                    'Interface contenant l'enveloppe de recherche.
        Dim pSelectionEnv As ISelectionEnvironment = Nothing    'Interface pour définir l'environnement de sélection.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing        'Interface pour extraire les Nodes sélectionnés.
        Dim pTopoNode As ITopologyNode = Nothing                'Interface contenant un Node.
        Dim pMouseCursor As IMouseCursor = Nothing              'Interface qui permet de changer l'image du curseur.
        Dim pTrackCancel As ITrackCancel = Nothing              'Interface qui permet d'annuler la sélection avec la touche ESC.

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Permettre d'annuler le traitement avec la touche ESC
            pTrackCancel = New CancelTracker
            pTrackCancel.CancelOnKeyPress = True
            pTrackCancel.CancelOnClick = False

            'Vider la sélection de la Map
            gpMxDoc.FocusMap.ClearSelection()

            'Créer une nouvelle géométrie polyline vide
            m_GeometrieSelection = New Polyline
            m_GeometrieSelection.SpatialReference = gpMxDoc.FocusMap.SpatialReference

            'Créer un nouvel enviromment de sélection
            pSelectionEnv = New SelectionEnvironmentClass()

            'Définir le rectangle en pixel
            oRectangle.left = arg.X - pSelectionEnv.SearchTolerance
            oRectangle.top = arg.Y - pSelectionEnv.SearchTolerance
            oRectangle.right = arg.X + pSelectionEnv.SearchTolerance
            oRectangle.bottom = arg.Y + pSelectionEnv.SearchTolerance
            'Transformer le rectangle en enveloppe '5 = esriTransformPosition + esriTransformToMap'.
            pEnvelope = New EnvelopeClass()
            gpMxDoc.ActiveView.ScreenDisplay.DisplayTransformation.TransformRect(pEnvelope, oRectangle, 5)

            'Vider la sélection des Edges
            m_TopologyGraph.SetSelectionEmpty(esriTopologyElementType.esriTopologyEdge)

            'Sélectionner un node
            m_TopologyGraph.SelectByGeometry(esriTopologyElementType.esriTopologyNode, esriTopologySelectionResultEnum.esriTopologySelectionResultNew, pEnvelope)

            'Interface pour extraire les Nodes trouvés
            pEnumTopoNode = m_TopologyGraph.NodeSelection

            'Extraire le premier Node
            pEnumTopoNode.Reset()
            pTopoNode = pEnumTopoNode.Next

            'Traiter tous les Nodes
            Do Until pTopoNode Is Nothing
                'Sélectionner le réseau par le Node sélectionné
                SelectionnerReseau(CType(m_GeometrieSelection, IPolyline), m_TopologyGraph, pTopoNode, False)

                'Extraire le prochain Node
                pTopoNode = pEnumTopoNode.Next
            Loop

            'Vider la sélection des Nodes
            m_TopologyGraph.SetSelectionEmpty(esriTopologyElementType.esriTopologyNode)

            'Vérifier si un réseau a été trouvé
            If Not m_GeometrieSelection.IsEmpty Then
                'Vérifier si on doit sélectionner le réseau (Plus lent)
                If arg.Shift Then
                    'Sélectionner les éléments du réseau
                    gpMxDoc.FocusMap.SelectByShape(m_GeometrieSelection, pSelectionEnv, False)
                End If

                'Définir l'enveloppe de base de la vue
                pEnvelope = m_GeometrieSelection.Envelope

                'Union entre l'enveloppe de la géométrie et celle de la vue active
                pEnvelope.Union(gpMxDoc.ActiveView.Extent)

                'Changer l'enveloppe de la vue active
                gpMxDoc.ActiveView.Extent = pEnvelope
            End If

            'Appel du module qui effectue le traitement
            Call bDessinerGeometrie(gpMxDoc, pTrackCancel, CType(m_GeometrieSelection, IGeometry), True)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oRectangle = Nothing
            pEnvelope = Nothing
            pSelectionEnv = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pMouseCursor = Nothing
            pTrackCancel = Nothing
            'Récupération de la mémoire disponble
            GC.Collect()
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            'Vérifier si la topologie est valide
            If m_TopologyGraph Is Nothing Then
                'Rendre désactive la commande
                Me.Enabled = False
            Else
                'Vérifier si la topologie contient des lignes
                If m_TopologyGraph.Edges.Count > 0 Then
                    'Rendre active la commande
                    Me.Enabled = True
                Else
                    'Rendre désactive la commande
                    Me.Enabled = False
                End If
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class
