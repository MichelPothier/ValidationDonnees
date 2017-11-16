'**
'Nom de la composante : cmdRetirerListeIdentifiant.vb
'
'''<summary>
'''Commande qui permet de retirer la sélection d'éléments de la liste d'identifiants des FeatureLayers visibles. 
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''</remarks>
''' 
Public Class cmdRetirerListeIdentifiant
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
            'Appel du module qui effectue le traitement pour combiner les Layers de même classe
            modBarreLayerSelection.RetirerListeIdentifiant(gpApp, True)

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            'Vérifier si au moins un élément est sélectionné
            If gpMxDoc.FocusMap.SelectionCount > 0 Then
                'Rendre active la commande
                Enabled = True
            Else
                'Rendre désactive la commande
                Enabled = False
            End If

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub
End Class
