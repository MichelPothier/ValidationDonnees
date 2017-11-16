Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : cmdDessinerGeometrie 
'
'''<summary>
''' Commande qui permet d'effectuer un Zoom sur les géométries en erreur dans la fenêtre graphique active.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 24 mars 2016
'''</remarks>
''' 
Public Class cmdZoomGeometrieErreur
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
        Dim pMouseCursor As IMouseCursor = Nothing      'Interface qui permet de changer l'image du curseur
        Dim pTrackCancel As ITrackCancel = Nothing      'Interface qui permet d'annuler la sélection avec la touche ESC
        Dim pEnvelope As IEnvelope = Nothing            'Interface contenant l'enveloppe de la géométrie.
        Dim pPoint As IPoint = Nothing                  'Point utilisé pour centrer l'enveloppe.

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Permettre d'annuler le traitement avec la touche ESC
            pTrackCancel = New CancelTracker
            pTrackCancel.CancelOnKeyPress = True
            pTrackCancel.CancelOnClick = False

            'Définir l'enveloppe
            pEnvelope = m_GeometrieSelection.Envelope

            'Vérifier s'il faut centrer
            If pEnvelope.Height = 0 And pEnvelope.Width = 0 Then
                'Définir le point pour centrer l'enveloppe
                pPoint = pEnvelope.LowerLeft
                'Définir l'enveloppe
                pEnvelope = gpMxDoc.ActiveView.Extent
                'Centrer l'enveloppe courante
                pEnvelope.CenterAt(pPoint)
            Else
                'Agrandir l'enveloppe de 10% de l'élément en erreur
                pEnvelope.Expand(pEnvelope.Width / 10, pEnvelope.Height / 10, False)
            End If

            'Changer l'enveloppe
            gpMxDoc.ActiveView.Extent = pEnvelope

            'Appel du module qui effectue le traitement pour dessiner les géométries
            Call bDessinerGeometrie(gpMxDoc, pTrackCancel, CType(m_GeometrieSelection, IGeometry), True)

        Catch erreur As Exception
            'Afficher le message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
            pTrackCancel = Nothing
            pEnvelope = Nothing
            pPoint = Nothing
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
            'Message d'erreur
        End Try
    End Sub
End Class
