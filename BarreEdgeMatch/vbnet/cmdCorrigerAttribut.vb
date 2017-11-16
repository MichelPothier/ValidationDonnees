Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms

'**
'Nom de la composante : cmdCorrigerAttribut.vb
'
'''<summary>
''' Commande qui permet de corriger les erreurs d'attributs aux points d'adjacence sur les limites communes de découpage.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être
'''utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''Date : 21 juillet 2011
'''</remarks>
''' 
Public Class cmdCorrigerAttribut
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

            'Corriger les erreurs d'attribut
            Call CorrigerErreurAttribut()

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
        'Vérifier si aucune erreur d'attribut
        If m_ErreurFeatureAttribut Is Nothing Then
            'Désactiver la commande
            Enabled = False
        Else
            'Vérifier si aucune erreur d'attribut
            If m_ErreurFeatureAttribut.Count = 0 Then
                'Désactiver la commande
                Enabled = False
            Else
                'Activer la commande
                Enabled = True
            End If
        End If
    End Sub
End Class
