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

Public Class cmdAfficherTopologie
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
        Dim pMap As IMap = Nothing                          'Interface ESRI contenant la Map active.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pGeometry As IGeometry = Nothing                'Interface ESRI contenant des géométries en mémoire.
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface ESRI contenant des géométries en mémoire.
        Dim pTrackCancel As ITrackCancel = Nothing          'Interface qui permet d'annuler la sélection avec la touche ESC.

        Try
            'Vérifier si la topologie est présente
            If m_TopologyGraph IsNot Nothing Then
                'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
                pMouseCursor = New MouseCursorClass
                pMouseCursor.SetCursor(2)

                'Permettre d'annuler la sélection avec la touche ESC
                pTrackCancel = New CancelTracker
                pTrackCancel.CancelOnKeyPress = True
                pTrackCancel.CancelOnClick = False

                'Définir la Map courante
                pMap = gpMxDoc.FocusMap

                'Créer un nouveau Bag
                pGeometry = New GeometryBag
                pGeometry.SpatialReference = pMap.SpatialReference
                pGeometryColl = CType(pGeometry, IGeometryCollection)

                'Sélectionner les éléments de la vue active
                pMap.SelectByShape(gpMxDoc.ActiveView.Extent, Nothing, False)

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pMap.FeatureSelection, IEnumFeature)

                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traitre tous les éléments de la sélection
                Do Until pFeature Is Nothing
                    Try
                        'Définir la nouvelle géométrie de l'élément
                        pGeometry = m_TopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                        'Vérifier si la géométrie est invalide
                        If pGeometry IsNot Nothing Then
                            'ajouter la géométrie dans le Bag
                            pGeometryColl.AddGeometry(pGeometry)
                        End If

                    Catch ex As Exception
                        'On ne fait rien
                    End Try

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Vider la sélection
                pMap.ClearSelection()

                'Redéfinir la géométrie
                pGeometry = CType(pGeometryColl, IGeometry)
                'Dessiner la géométrie
                bDessinerGeometrie(gpMxDoc, pTrackCancel, pGeometry, True)
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
            pMap = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pTrackCancel = Nothing
            pGeometryColl = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            'Vérifier si la topologie est valide
            If m_TopologyGraph Is Nothing Then
                'Rendre désactive la commande
                Me.Enabled = False
            Else
                'Vérifier si la topologie n'est pas vide
                If m_TopologyGraph.Edges.Count > 0 Or m_TopologyGraph.Nodes.Count > 0 Then
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
