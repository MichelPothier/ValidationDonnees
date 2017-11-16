Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : cmdDessinerGeometrie 
'
'''<summary>
''' Commande qui permet de dessiner dans la fenêtre graphique les géométries de travail en mémoire.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 21 avril 2015
'''</remarks>
Public Class cmdDessinerGeometrie
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    'Déclarer les variables globale de travail
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
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

                    'Conserver le document
                    m_MxDocument = gpMxDoc

                    'Initialiser les symbole d'affiche de géométrie
                    Call InitSymbole()
                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        'Déclarer les variables de travail
        Dim pMouseCursor As IMouseCursor = Nothing    'Interface qui permet de changer l'image du curseur
        Dim pTrackCancel As ITrackCancel = Nothing      'Interface qui permet d'annuler la sélection avec la touche ESC

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Permettre d'annuler le traitement avec la touche ESC
            pTrackCancel = New CancelTracker
            pTrackCancel.CancelOnKeyPress = True
            pTrackCancel.CancelOnClick = False

            'Appel du module qui effectue le traitement
            Call bDessinerGeometrie(gpMxDoc, pTrackCancel, CType(m_GeometrieSelection, IGeometry), False)

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            pMouseCursor = Nothing
            pTrackCancel = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            'Rendre inactive par défaut
            Enabled = False

            'Sortir si aucune géométrie de travail
            If m_GeometrieSelection IsNot Nothing Then
                'Vérifier si au moins une géométrie est présente
                If Not m_GeometrieSelection.IsEmpty Then
                    'Rendre active
                    Enabled = True
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class

