Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Data.OleDb

'**
'Nom de la composante : cboRequete.vb
'
'''<summary>
''' Commande qui permet de définir la requête utilisée pour sélectionner les éléments de la classe de sélection selon les paramètres spécifiés.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 14 avril 2015
'''</remarks>
''' 
Public Class cboRequete
    Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox

    'Déclarer les variables de travail
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

                    'Vider la mémoire
                    m_Requete = Nothing
                    'Créer la requête par défaut
                    m_Requete = New clsRequete

                    'Traiter toutes les requêtes possibles
                    For Each sRequete In m_Requete.Requetes
                        'Ajouter la requête
                        Me.Add(CStr(sRequete), CStr(sRequete))
                    Next

                    'Définir la requête par défaut
                    Me.Value = m_Requete.Nom

                    'Définir les paramètres de la requête par défaut
                    m_Parametres = m_Requete.Requete.Parametres

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

            'Vider la mémoire
            m_Requete = Nothing
            'Définir la requête. 
            m_Requete = New clsRequete(CStr(Me.GetItem(cookie).Tag), m_FeatureLayer)
            'Définir les paramètres de la requête
            m_cboParametres.initParametres()

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overloads Overrides Sub OnFocus(ByVal focus As Boolean)
        Try
            'Si le focus est perdu
            If Not focus Then
                'Traiter tous les items de la collection. 
                For Each item As ESRI.ArcGIS.Desktop.AddIns.ComboBox.Item In Me.items
                    'Vérifier si la valeur est déjà présente
                    If Me.Value = item.Caption Then
                        'Vérifier si la requête est changée
                        If Me.Value <> m_Requete.Nom Then
                            'Vider la mémoire
                            m_Requete = Nothing
                            'Définir la requête
                            m_Requete = New clsRequete(Me.Value, m_FeatureLayer)
                            'Définir les paramètres de la requête
                            m_cboParametres.initParametres()
                        End If
                        'Sortir
                        Exit Sub
                    End If
                Next
                'Remettre l'ancienne requête
                Me.Value = m_Requete.Nom
            End If

        Catch erreur As Exception
            'Afficher un message d'erreur
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
