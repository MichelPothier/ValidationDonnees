﻿Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms

'**
'Nom de la composante : cmdCorrigerFenetre.vb
'
'''<summary>
''' Commande qui permet de corriger les erreurs de précision, d'adjacence et d'attribut présentent dans la fenêtre graphique active.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''Date : 7 novembre 2013
'''</remarks>
''' 
Public Class cmdCorrigerFenetre
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
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

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        'Déclarer les variables de travail
        Dim pMouseCursor As IMouseCursor = Nothing
        Dim bSelect As Boolean = Nothing

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Conserver la sélection de la classe de découpage
            bSelect = m_FeatureLayerDecoupage.Selectable

            'Désactiver la sélection de la classe de découpage
            m_FeatureLayerDecoupage.Selectable = False

            'Corriger seulement les erreurs de précision présentent dans la fenêtre active
            Call CorrigerErreurPrecision(True)

            'Corriger seulement les erreurs d'adjacence présentent dans la fenêtre active
            Call CorrigerErreurAdjacence(True)

            'Corriger seulement les erreurs d'attribut présentent dans la fenêtre active
            Call CorrigerErreurAttribut(True)

            'Redéfinir la sélection de la classe de découpage
            m_FeatureLayerDecoupage.Selectable = bSelect

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si aucune erreur de précision
        If m_ErreurFeaturePrecision Is Nothing Or m_ErreurFeatureAdjacence Is Nothing Or m_ErreurFeatureAttribut Is Nothing Then
            'Désactiver la commande
            Enabled = False
        Else
            'Vérifier si aucune erreur de précision
            If m_ErreurFeaturePrecision.Count = 0 And m_ErreurFeatureAdjacence.Count = 0 And m_ErreurFeatureAttribut.Count = 0 Then
                'Désactiver la commande
                Enabled = False
            Else
                'Activer la commande
                Enabled = True
            End If
        End If
    End Sub
End Class
