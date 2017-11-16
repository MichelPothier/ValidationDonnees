Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Data.OleDb

'**
'Nom de la composante : cboTypeSelection.vb
'
'''<summary>
''' Commande qui permet de définir le type de sélection utilisé pour sélectionner les éléments de la classe de sélection.
''' 
''' Seuls les éléments sélectionnés sont traités.
''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
''' 
''' Conserver : Les éléments qui respectent la requête sont conservés dans la sélection.
''' Enlever : Les éléments qui respectent la requête sont enlevés dans la sélection.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 14 avril 2015
'''</remarks>
''' 
Public Class cboTypeSelection
    Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox

    'Déclarer les variables de travail
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap
    Dim iDefaut As Integer = Nothing        'Index par défaut

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

                    'Ajouter le type de sélection CONSERVER
                    Me.Add("Conserver", "Conserver")

                    'Ajouter le type de sélection ENLEVER
                    m_TypeSelection = "Enlever"
                    iDefaut = Me.Add(m_TypeSelection, m_TypeSelection)
                    'Sélectionné la valeur par défaut
                    Me.Select(iDefaut)

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overloads Overrides Sub OnSelChange(ByVal cookie As Integer)
        Try
            'Si aucun n'est sélectionné
            If cookie = -1 Then
                'On ne fait rien
                Exit Sub
            End If

            'Définir le type de sélection utilisée 
            m_TypeSelection = CStr(Me.GetItem(cookie).Tag)

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Vérifier si le FeatureLayer de sélection est invalide
        If m_FeatureLayer Is Nothing Then
            'Rendre désactive la commande
            Me.Enabled = False

            'Si le FeatureLayer de sélection est invalide
        Else
            'Rendre active la commande
            Me.Enabled = True
        End If
    End Sub
End Class
