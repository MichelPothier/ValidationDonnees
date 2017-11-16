Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

'**
'Nom de la composante : tooSelectGroupe.vb
'
'''<summary>
'''Commande qui permet sélectionner un groupe d'éléments basé sur la valeur d'un attribut du premier élément trouvé.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''</remarks>
''' 
Public Class tooSelectGroupe
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
        Dim pEnvelope As IEnvelope = Nothing                    'Interface contenant l'enveloppe de recherche.
        Dim pSelectionEnv As ISelectionEnvironment = Nothing    'Interface pour définir l'environnement de sélection.
        Dim pSpatialFilter As New SpatialFilterClass            'Interface pour effectuer une requête spatiale.
        Dim pQueryFilter As New QueryFilterClass                'Interface pour effectuer une requête attributive.
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour sélectionner les éléments.
        Dim pEnumId As IEnumIDs = Nothing                       'Interface qui permet d'extraire les Ids des éléments sélectionnés.
        Dim pFeature As IFeature = Nothing                      'Interface contenant l'élément sélectionné.
        Dim iOid As Integer = Nothing                           'Contient un Id d'un élément sélectionné.
        Dim iPos As Integer = -1                                'Contient la position de l'attribut.
        Dim oRectangle As tagRECT                               'Contient le rectangle en pixel.

        Try
            'Vider la sélection de la Map
            gpMxDoc.FocusMap.ClearSelection()

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

            'Définir la requête spatiale
            pSpatialFilter.Geometry = pEnvelope
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            pSpatialFilter.OutputSpatialReference(m_FeatureLayer.FeatureClass.ShapeFieldName) = gpMxDoc.FocusMap.SpatialReference
            pSpatialFilter.GeometryField = m_FeatureLayer.FeatureClass.ShapeFieldName

            'Interface pour sélectionner les éléments.
            pFeatureSel = CType(m_FeatureLayer, IFeatureSelection)

            'Sélectionner les éléments selon l'enveloppe
            pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, True)

            'Vérifier si un élément a été trouvé
            If pFeatureSel.SelectionSet.Count = 1 Then
                'Interface pour extrire la liste des Ids
                pEnumId = pFeatureSel.SelectionSet.IDs

                'Extraire le premier élément sélectionné
                pEnumId.Reset()
                iOid = pEnumId.Next
                pFeature = m_FeatureLayer.FeatureClass.GetFeature(iOid)

                'Vérifier si l'élément est trouvé
                If pFeature IsNot Nothing Then
                    'Définir la position de l'attribut
                    iPos = m_FeatureLayer.FeatureClass.FindField(m_cboAttributGroupe.NomAttribut)

                    'vérifier si l'atrribut est présent
                    If iPos > -1 Then
                        'Vérifier si le type d'attribut est un texte
                        If m_FeatureLayer.FeatureClass.Fields.Field(iPos).Type = esriFieldType.esriFieldTypeString Then
                            'Définir la requête attributive
                            pQueryFilter.WhereClause = m_cboAttributGroupe.NomAttribut & "='" & pFeature.Value(iPos).ToString & "'"
                        Else
                            'Définir la requête attributive
                            pQueryFilter.WhereClause = m_cboAttributGroupe.NomAttribut & "=" & pFeature.Value(iPos).ToString
                        End If

                        'Sélectionner les éléments selon l'enveloppe et l'environnement de sélection
                        pFeatureSel.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultAdd, False)
                    End If
                End If
            End If

            'Affichage de la sélection
            gpMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, gpMxDoc.ActiveView.Extent)

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pEnvelope = Nothing
            pSelectionEnv = Nothing
            pSpatialFilter = Nothing
            pQueryFilter = Nothing
            pFeatureSel = Nothing
            pEnumId = Nothing
            pFeature = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            'Vérifier si le FeatureLayer de sélection est invalide
            If m_FeatureLayer Is Nothing Then
                'Rendre désactive la commande
                Me.Enabled = False

                'Si le FeatureLayer de sélection est invalide
            Else
                'Rendre active la commande
                Me.Enabled = True
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class
