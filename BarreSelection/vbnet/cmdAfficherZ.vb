Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.DataSourcesRaster
Imports System.IO

'''<summary>
''' Commande qui permet d'afficher le Z des extrémités des éléments du FeatureLayer de sélection à partir des élévations des MNE visibles.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
''' utilisé dans ArcMap (ArcGisESRI).
'''</remarks>
Public Class cmdAfficherZ
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    'Déclarer les variables de travail
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
        'Définir les variables de travail

        Try
            'Par défaut la commande est inactive
            Enabled = False
            'Vérifier si l'application est définie
            If Not Hook Is Nothing Then
                'Définir l'application
                gpApp = CType(Hook, IApplication)

                'Vérifier si l'application est ArcMap
                If TypeOf Hook Is IMxApplication Then
                    'Définir le document
                    gpMxDoc = CType(gpApp.Document, IMxDocument)
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        'Déclarer les variables de travail
        Dim oMapLayer As clsGererMapLayer = Nothing         'Objet utilisé pour extraire les RasterLayer visibles.
        Dim pRasterLayerColl As Collection = Nothing        'Contient la collection des RasterLayer visibles.
        Dim pRasterLayer As IRasterLayer = Nothing          'Contient un RasterLayer.
        Dim pMultiPoint As IMultipoint = Nothing            'Contient la liste des extrémités de ligne.
        Dim pMouseCursor As IMouseCursor = Nothing          'Interface qui permet de changer le curseur de la souris.

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Extraire les extrémités de ligne du Layer de sélection contenu dans la vue active
            pMultiPoint = ExtraireExtremiteLigneVueActive(m_FeatureLayer, gpMxDoc.ActivatedView.Extent)

            'Définir l'objet utilisé pour extraire la collection des RasterLayers visibles
            oMapLayer = New clsGererMapLayer(gpMxDoc.FocusMap)

            'Extraire la collection des RasterLayers visibles
            pRasterLayerColl = oMapLayer.RasterLayerCollection

            'Traiter tous les RasterLayer
            For i = 1 To pRasterLayerColl.Count
                'Définir le RasterLayer
                pRasterLayer = CType(pRasterLayerColl.Item(i), IRasterLayer)

                'Afficher les valeurs d'élévation pour toutes les extrémités des lignes
                AfficherValeurZ(gpMxDoc.ActiveView.ScreenDisplay, CType(pRasterLayer.Raster, IRaster2), pMultiPoint)
            Next

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oMapLayer = Nothing
            pRasterLayerColl = Nothing
            pRasterLayer = Nothing
            pMultiPoint = Nothing
            pMouseCursor = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si le FeatureLayer de sélection est invalide
        If m_FeatureLayer IsNot Nothing Then
            'Rendre active la commande
            Me.Enabled = True

            'Si le FeatureLayer de sélection est invalide
        Else
            'Rendre désactive la commande
            Me.Enabled = False
        End If
    End Sub

    '''<summary>
    ''' Fonction qui permet d'extraire toutes les extrémités de ligne d'un FeatureLayer contenues dans la fenêtre d'affichage active.
    ''' Les extrémités de lignes sont retournées dans un Multipoint.
    '''</summary>
    '''
    '''<param name="pFeatureLayer">Interface contenant les lignes dans lesquels les extrémités seraont extraites.</param>
    '''<param name="pEnvelope">Interface contenant l'enveloppe de la vue active dans lequel les extrémités de lignes seront extraites.</param>
    ''' 
    ''' <returns> Imultipoint contenant les extrémité de lignes. </returns>
    ''' 
    Private Function ExtraireExtremiteLigneVueActive(ByVal pFeatureLayer As IFeatureLayer, ByVal pEnvelope As IEnvelope) As IMultipoint
        'Déclarer les variables de travail
        Dim pMultiPoint As IMultipoint = Nothing            'Contient la liste des extrémités de ligne.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les extrémités de ligne.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les extrémités de ligne.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour effectuer l'union des extrémités de ligne.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface pour extraire les items d'un RasterCatalog.
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface utilisé pour sélectionner les Rasters du Catalogue dans la vue active.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la ligne à traiter.

        'Définir la valeur de retour par défaut
        ExtraireExtremiteLigneVueActive = New Multipoint
        ExtraireExtremiteLigneVueActive.SpatialReference = pFeatureLayer.AreaOfInterest.SpatialReference

        Try
            'Vérifier si le FeatureLayer est valide
            If pFeatureLayer IsNot Nothing Then
                'Vérifier si le FeatureLayer est valide
                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                    'Interface pour ajouter les extrémités de ligne
                    pGeomColl = New GeometryBag

                    'Créer une nouvelle requête spatiale
                    pSpatialFilter = New SpatialFilter

                    'Définir la requête spatiale
                    pSpatialFilter.Geometry = pEnvelope
                    pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    pSpatialFilter.OutputSpatialReference(pFeatureLayer.FeatureClass.ShapeFieldName) = ExtraireExtremiteLigneVueActive.SpatialReference
                    pSpatialFilter.GeometryField = pFeatureLayer.FeatureClass.ShapeFieldName

                    'Interface pour extraire tous les éléments à afficher le Z
                    pFeatureCursor = pFeatureLayer.Search(pSpatialFilter, True)

                    'Extraire le premier élément
                    pFeature = pFeatureCursor.NextFeature

                    'Traiter tous les éléments
                    Do Until pFeature Is Nothing
                        'Interface pour extraire le premier et le dernier sommet
                        pPolyline = CType(pFeature.Shape, IPolyline)

                        'Créer un Multipoint vide
                        pMultiPoint = New Multipoint
                        pMultiPoint.SpatialReference = ExtraireExtremiteLigneVueActive.SpatialReference
                        pMultiPoint.SnapToSpatialReference()

                        'Interface pour ajouter les extrémités de ligne
                        pPointColl = CType(pMultiPoint, IPointCollection)

                        'Ajouter le premier sommet
                        pPointColl.AddPoint(pPolyline.FromPoint)

                        'Ajouter le dernier sommet
                        pPointColl.AddPoint(pPolyline.ToPoint)

                        'Ajouter le dernier sommet
                        pGeomColl.AddGeometry(pMultiPoint)

                        'Extraire le prochain élément
                        pFeature = pFeatureCursor.NextFeature
                    Loop

                    'Créer un Multipoint vide
                    pMultiPoint = New Multipoint
                    pMultiPoint.SpatialReference = ExtraireExtremiteLigneVueActive.SpatialReference
                    pMultiPoint.SnapToSpatialReference()

                    'Union des extrémités
                    pTopoOp = CType(pMultiPoint, ITopologicalOperator2)
                    pTopoOp.ConstructUnion(CType(pGeomColl, IEnumGeometry))

                    'Retourner le multipoint contenant les extrémités de la vue active
                    ExtraireExtremiteLigneVueActive = pMultiPoint
                End If
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pPointColl = Nothing
            pGeomColl = Nothing
            pMultiPoint = Nothing
            pTopoOp = Nothing
            pFeatureCursor = Nothing
            pSpatialFilter = Nothing
            pFeature = Nothing
            pPolyline = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'afficher les valeurs des pixels de toutes les extrémités des lignes dans la fenêtre d'affichage active.
    '''</summary>
    '''
    '''<param name="pScreenDisplay">Interface contenant la fenêtre d'affichage.</param>
    '''<param name="pRaster">Interface contenant le MNE.</param>
    '''<param name="pMultiPoint">Interface contenant les extrémités de lignes à afficher.</param>
    ''' 
    Private Sub AfficherValeurZ(ByVal pScreenDisplay As IScreenDisplay, ByVal pRaster As IRaster2, ByVal pMultiPoint As IMultipoint)
        'Déclarer les variable de travail
        Dim pSpatialRef As ISpatialReference = Nothing  'Interface contenant les propriétés de l'image.
        Dim pRasterProps As IRasterProps = Nothing      'Interface contenant les propriétés de l'image.
        Dim pPointColl As IPointCollection = Nothing    'Interface utilisé pour extraire les points.
        Dim pPoint As IPoint = Nothing                  'Contient le point traité.
        Dim vValue As Object = Nothing                  'Contient la valeur du pixel traité.
        Dim iRow As Integer = -1            'Contient le numéro de la rangé de l'image.
        Dim iCol As Integer = -1            'Contient le numéro de la colonne de l'image.

        Try
            'Conserver la référence spatiale
            pSpatialRef = pMultiPoint.SpatialReference

            'Interface pour extraire les propriétés de l'image
            pRasterProps = CType(pRaster, IRasterProps)

            'Projeter les extrémités de ligne
            pMultiPoint.Project(pRasterProps.SpatialReference)

            'Interface pour extraire les extrémités de ligne
            pPointColl = CType(pMultiPoint, IPointCollection)

            'Traiter tous les sommets
            For i = 0 To pPointColl.PointCount - 1
                'Définir le point
                pPoint = pPointColl.Point(i)

                'Retourner la position en pixel du coin supérieur gauche
                pRaster.MapToPixel(pPoint.X, pPoint.Y, iCol, iRow)

                'Extraire la valeur du pixel
                vValue = pRaster.GetPixelValue(0, iCol, iRow)

                'Vérifier si la valeur est présente
                If vValue IsNot Nothing Then
                    'Projeter le point
                    pPoint.Project(pSpatialRef)

                    'Dessiner la valeur du pixel
                    DessinerPixel(pScreenDisplay, pPoint, CInt(vValue).ToString)
                End If
            Next

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pSpatialRef = Nothing
            pRasterProps = Nothing
            pPointColl = Nothing
            vValue = Nothing
            pPoint = Nothing
        End Try
    End Sub
End Class
