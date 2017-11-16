Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Data.OleDb
Imports System.Windows.Forms

'**
'Nom de la composante : cboParametres.vb
'
'''<summary>
''' Commande qui permet de définir les paramètres de la requête utilisé pour sélectionner les éléments de la classe de sélection.
'''</summary>
'''
'''<remarks>
''' Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
''' Auteur : Michel Pothier
''' Date : 14 avril 2015
'''</remarks>
''' 
Public Class cboParametres
    Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox

    'Déclarer les variables de travail
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap
    Dim gsNomRequete As String = ""         'Nom de la requête

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

                    m_cboParametres = Me

                    'Vérifier si la requête est valide
                    If m_Requete IsNot Nothing Then
                        'Définir la requête
                        gsNomRequete = m_Requete.Nom
                        'Remplir la liste des paramètres
                        For Each sParametre As String In m_Requete.Requete.ListeParametres
                            'Ajouter la valeur dans la liste des items
                            Me.Add(sParametre, sParametre)
                        Next
                        'Définir la valeur par défaut
                        Me.Value = CStr(m_Requete.Requete.ListeParametres.Item(1))
                    End If

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub

    Protected Overloads Overrides Sub OnFocus(ByVal focus As Boolean)
        Try
            'Si le focus est perdu
            If Not focus Then
                'Définir les paramètres de la requête
                m_Parametres = Me.Value
                'Traiter tous les items de la collection. 
                For Each item As ESRI.ArcGIS.Desktop.AddIns.ComboBox.Item In Me.items
                    'Vérifier si la valeur est déjà présente
                    If Me.Value = item.Caption Then
                        'Sortir
                        Exit Sub
                    End If
                Next
                'Ajouter la valeur dans la liste des items
                Me.Add(Me.Value, Me.Value)
            End If

        Catch erreur As Exception
            'Afficher un message d'erreur
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

            'Définir les paramètres de la requête
            m_Parametres = CStr(Me.GetItem(cookie).Tag)

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

    '''<summary>
    ''' Routine qui permet d'initialiser la liste des paramètres possibles et sélectionner le premier par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub initParametres()
        Try
            'Vérifier si la requête a changé
            If m_Requete.Nom <> gsNomRequete Then
                'Vider la liste des paramètres
                Me.Clear()
                'Remplir la liste des paramètres
                For Each sParametre As String In m_Requete.Requete.ListeParametres
                    'Ajouter la valeur dans la liste des items
                    Me.Add(sParametre, sParametre)
                Next
                'Définir le premier paramètre par défaut
                Me.Value = CStr(m_Requete.Requete.ListeParametres.Item(1))
                m_Parametres = Me.Value
                'Conserver le nom de la requête
                gsNomRequete = m_Requete.Nom
            End If

        Catch erreur As Exception
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        End Try
    End Sub
End Class
