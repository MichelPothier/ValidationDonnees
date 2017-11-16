Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase

Public Class cboAttributGroupe
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

                    'Conserver le ComboBox en mémoire
                    m_cboAttributGroupe = Me

                    'Remplir la liste des attributs
                    Me.initialiser()

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

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
    ''' Routine qui permet d'initialiser la liste des attributs de la classe de sélection .
    ''' 
    '''</summary>
    '''
    Public Sub initialiser()
        'déclarer les variables de travail
        Dim pFields As IFields = Nothing     'Interface contenant un attribut de la classe de sélection.

        Try
            'Vider la liste des attributs de groupe
            Me.Clear()

            'Vérifier si la featureClass est valide
            If m_FeatureLayer IsNot Nothing Then
                'Interface contenant les attributs de la classe de sélection
                pFields = m_FeatureLayer.FeatureClass.Fields

                'Remplir la liste des paramètres
                For i = 0 To pFields.FieldCount - 1
                    'Vérifier si l'attribut est de type entier ou texte
                    If pFields.Field(i).Type = esriFieldType.esriFieldTypeString _
                    Or pFields.Field(i).Type = esriFieldType.esriFieldTypeInteger Then
                        'Ajouter la valeur dans la liste des items
                        Me.Add(pFields.Field(i).Name, pFields.Field(i).Name)
                    End If
                Next
                'Définir la valeur par défaut
                If Me.items.Count > 0 Then Me.Value = CStr(Me.items.Item(0).Tag)
            End If

        Catch erreur As Exception
            'Message d'erreur
            MsgBox("--Message: " & erreur.Message & vbCrLf & "--Source: " & erreur.Source & vbCrLf & "--StackTrace: " & erreur.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            pFields = Nothing
        End Try
    End Sub

    Public Function NomAttribut() As String
        'Extraire le nom de l'attribut
        NomAttribut = Me.Value
    End Function
End Class
