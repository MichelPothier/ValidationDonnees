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
''' Commande qui permet d'afficher les valeurs des pixels et leurs formes selon un nombre minimum et maximum de pixels à afficher en X et Y
''' pour tous les RasterLayer présents et visibles dans la fenêtre d'affichage active.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
''' utilisé dans ArcMap (ArcGisESRI).
'''</remarks>
Public Class cmdAfficherValeurPixel
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
        Dim pMouseCursor As IMouseCursor = Nothing          'Interface qui permet de changer le curseur de la souris.
        Dim pDocument As IDocument = Nothing                'Interface utilisé pour extraire la barre de statut.
        Dim oMapLayer As clsGererMapLayer = Nothing         'Objet utilisé pour extraire les RasterLayer visibles.
        Dim pRasterLayerColl As Collection = Nothing        'Contient la collection des RasterLayer visibles.
        Dim pRasterLayer As IRasterLayer = Nothing          'Contient un RasterLayer.
        Dim oPixel2Polygon As clsPixel2Polygon = Nothing    'Objet utilisé pour afficher les valeurs de pixels et leurs formes.
        Dim nNbPixelX As Integer = 36       'Contient le nombre de pixel à afficher en X.
        Dim nNbPixelY As Integer = 36       'Contient le nombre de pixel à afficher en Y.

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Interface pour définir le document afin d'extraire le StatusBar
            pDocument = CType(gpMxDoc, IDocument)

            'Définir le nonbre de pixel selon la forme de la fenêtre active
            If gpMxDoc.ActiveView.Extent.Width > gpMxDoc.ActiveView.Extent.Height Then
                nNbPixelY = CInt(gpMxDoc.ActiveView.Extent.Height / (gpMxDoc.ActiveView.Extent.Width / nNbPixelX))
            Else
                nNbPixelX = CInt(gpMxDoc.ActiveView.Extent.Width / (gpMxDoc.ActiveView.Extent.Height / nNbPixelY))
            End If

            'Définir l'objet utilisé pour extraire la collection des RasterLayers visibles
            oMapLayer = New clsGererMapLayer(gpMxDoc.FocusMap)

            'Extraire la collection des RasterLayers visibles
            pRasterLayerColl = oMapLayer.RasterLayerCollection

            'Traiter tous les RasterLayer
            For i = 1 To pRasterLayerColl.Count
                'Définir le RasterLayer
                pRasterLayer = CType(pRasterLayerColl.Item(i), IRasterLayer)

                'Définir un nouvel objet pour traiter l'affichage des pixels de l'image matricielle
                oPixel2Polygon = New clsPixel2Polygon(pRasterLayer)

                'Affichage des valeurs de pixels et leurs formes selon un nombre minimum et maximum de pixel par ligne et colonne
                oPixel2Polygon.AfficherValeurPixel(gpMxDoc.ActiveView, pDocument.Parent.StatusBar, nNbPixelX, nNbPixelY)
            Next

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
            pDocument = Nothing
            oMapLayer = Nothing
            pRasterLayerColl = Nothing
            pRasterLayer = Nothing
            oPixel2Polygon = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si aucun Layer n'est présent
        If gpMxDoc.FocusMap.LayerCount = 0 Then
            Enabled = False
        Else
            Enabled = True
        End If
    End Sub
End Class
