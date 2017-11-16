Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Geodatabase

'**
'Nom de la composante : cmdCreerTopologie.vb
'
'''<summary>
''' Commande qui permet de créer la topologie en mémoire à partir des FeatureLayer visible.
''' 
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 7 juillet 2015
'''</remarks>
''' 
Public Class cmdCreerTopologie
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
        Dim oGererMapLayer As clsGererMapLayer = Nothing            'Objet pour extraire la collection des FeatureLayers visibles.
        Dim pMouseCursor As IMouseCursor = Nothing                  'Interface qui permet de changer l'image du curseur
        Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.

        Try
            'Afficher le message
            gpApp.StatusBar.Message(0) = "Création de la topologie des éléments visibles en cours ..."

            'Forcer l'exécution des événements
            System.Windows.Forms.Application.DoEvents()

            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Object pour extraire la collection des FeatureLayer visbles
            oGererMapLayer = New clsGererMapLayer(gpMxDoc.FocusMap)

            'Interface pour extraire la tolérance de précision de la référence spatiale
            pSRTolerance = CType(gpMxDoc.FocusMap.SpatialReference, ISpatialReferenceTolerance)
            'Extraire la précision
            m_Precision = pSRTolerance.XYTolerance
            'Mettre à jour la précision dans le menu des paramètres de sélection
            m_MenuParametresSelection.txtPrecision.Text = m_Precision.ToString

            'Vider la mémoire
            m_TopologyGraph = Nothing
            'Récupérer la mémoire
            GC.Collect()

            'Créer la topologie
            'm_TopologyGraph = CreerTopologyGraph2(oGererMapLayer.EnvelopeFeatureLayerCollection, oGererMapLayer.FeatureLayerCollection, m_Precision)
            m_TopologyGraph = CreerTopologyGraph2(gpMxDoc.ActiveView.Extent, oGererMapLayer.FeatureLayerCollection, m_Precision)

            'Permettre l'affichage du nombre de noeuds créés
            gpApp.StatusBar.Message(0) = "Succès du traitement : Nodes:" & m_TopologyGraph.Nodes.Count.ToString & ", Edges:" & m_TopologyGraph.Edges.Count.ToString

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oGererMapLayer = Nothing
            pMouseCursor = Nothing
            pSRTolerance = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            'Rendre acrif la commande
            Me.Enabled = True

        Catch erreur As Exception
            Throw
        End Try
    End Sub
End Class
