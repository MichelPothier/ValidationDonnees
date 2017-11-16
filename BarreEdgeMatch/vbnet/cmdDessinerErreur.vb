Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry

'**
'Nom de la composante : cmdDessinerErreur.vb
'
'''<summary>
''' Commande qui permet de dessiner les erreurs de précision, d'adjacence et d'attribut présentent aux limites communes de découpage.
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
Public Class cmdDessinerErreur
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
        Try
            'Dessiner la géométrie et le texte des erreurs de précision, d'adjacence et d'attribut
            Call DessinerErreurs()

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si aucune erreur
        If m_ErreurFeature Is Nothing Then
            'Désactiver la commande
            Enabled = False
        Else
            'Vérifier si aucune erreur
            If m_ErreurFeature.Count = 0 Then
                'Désactiver la commande
                Enabled = False
            Else
                'Activer la commande
                Enabled = True
            End If
        End If
    End Sub
End Class
