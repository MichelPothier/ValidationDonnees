'**
'Nom de la composante : cmdCreerSelection.vb
'
'''<summary>
'''Commande qui permet de créer un nouveau « FeatureLayer » contenant seulement les éléments de la sélection 
'''de chacun des « FeatureLayers » visibles dans lesquels une sélection est présente. 
'''Les anciens « FeatureLayers » seront détruits. Seuls les éléments sélectionnés seront conservés dans les 
'''« FeatureLayers » peu importe s’ils contiennent ou non une liste d’identifiants d’éléments. 
'''Les requêtes attributives seront conservées dans chacun des « FeatureLayers » si présentes, mais pas les jointures. 
'''Les « FeatureLayers » non visibles ou n’ayant aucun élément de sélectionné ne seront pas affectés par ce traitement.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''</remarks>
''' 

Public Class cmdCreerListeIdentifiant
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
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
                    'Vérifier si au moins un élément est sélectionné
                    If gpMxDoc.FocusMap.SelectionCount > 0 Then
                        'Rendre active la commande
                        Enabled = True
                    End If
                End If
            End If

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        Try
            'Appel du module qui effectue le traitement pour créer une liste d'identifiant d'éléments dans le FeatureLayer
            modBarreLayerSelection.CreerListeIdentifiant(gpApp, True)

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            'Rendre désactive la commande
            Enabled = False
            'Vérifier si au moins un élément est sélectionné
            If gpMxDoc.FocusMap.SelectionCount > 0 Then
                'Rendre active la commande
                Enabled = True
            End If

            Checked = True

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub
End Class
