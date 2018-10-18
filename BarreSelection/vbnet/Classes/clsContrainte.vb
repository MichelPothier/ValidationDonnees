Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports Microsoft.VisualBasic.FileIO
Imports System.Windows.Forms
Imports System.Text
Imports System.IO
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.DataManagementTools

'**
'Nom de la composante : clsContrainte.vb
'
'''<summary>
''' Classe générique qui permet de gérer une contrainte d'intégrité quelconque.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 19 avril 2016
'''</remarks>
''' 
Public Class clsContrainte
    '''<summary>Contient le numéro de l'identifiant de la contrainte d'intégrité.</summary>
    Protected giObjectId As Integer = -1
    '''<summary>Contient le groupe de la contrainte d'intégrité.</summary>
    Protected gsGroupe As String = Nothing
    '''<summary>Contient la description de la contrainte d'intégrité.</summary>
    Protected gsDescription As String = Nothing
    '''<summary>Contient le message de la contrainte d'intégrité.</summary>
    Protected gsMessage As String = Nothing
    '''<summary>Contient les requêtes de la contrainte d'intégrité.</summary>
    Protected gsRequetes As String = Nothing
    '''<summary>Contient le nom de la requête principale de la contrainte d'intégrité.</summary>
    Protected gsNomRequete As String = Nothing
    '''<summary>Contient le nom de la table pour laquelle la contrainte d'intégrité est appliquée.</summary>
    Protected gsNomTable As String = Nothing
    '''<summary>Entete d'éxécution du traitement.</summary>
    Protected gsEntete As String = ""
    '''<summary>Identifiant d'éxécution du traitement.</summary>
    Protected gsIdentifiant As String = ""
    '''<summary>Interface contenant le nom de la classe de découpage à traiter.</summary>
    Protected gsNomClasseDecoupage As String = ""
    '''<summary>Nom du répertoire ou de la géodatabase d'erreurs.</summary>
    Protected gsNomRepertoireErreurs As String = ""
    '''<summary>Nom du fichier journal d'éxécution du traitement.</summary>
    Protected gsNomFichierJournal As String = ""
    '''<summary>Contient l'information de la contrainte d'intégrité.</summary>
    Protected gsInformation As String = ""

    '''<summary>Contient le nombre d'éléments traités.</summary>
    Protected giNombreElements As Long = 0

    '''<summary>Interface utilisé pour afficher la barre de progression et pour annuler le traitement.</summary>
    Protected gpTrackCancel As ITrackCancel = Nothing
    '''<summary>RichTextBox utilisé pour écrire le journal d'éxécution du traitement.</summary>
    Protected gqRichTextBox As RichTextBox = Nothing

    '''<summary>Interface contenant la référence spatiale de traitement.</summary>
    Protected gpSpatialReference As ISpatialReference = Nothing
    '''<summary>Interface contenant une requête attributive.</summary>
    Protected gpQueryFilter As IQueryFilter = Nothing

    '''<summary>Interface contenant les classes des éléments à valider.</summary>
    Protected gpGeodatabase As IFeatureWorkspace = Nothing
    '''<summary>Interface contenant un élément de découpage à traiter.</summary>
    Protected gpFeatureDecoupage As IFeature = Nothing
    '''<summary>Interface contenant le FeatureLayer des éléments de découpage à traiter.</summary>
    Protected gpFeatureLayerDecoupage As IFeatureLayer = Nothing
    '''<summary>Interface contenant le FeatureLayer des éléments sélectionnés.</summary>
    Protected gpFeatureLayerSelection As IFeatureLayer = Nothing
    '''<summary>Interface contenant le FeatureLayer des éléments décrivant les erreurs.</summary>
    Protected gpFeatureLayerErreur As IFeatureLayer = Nothing
    '''<summary>Interface contenant les géométries décrivant les erreurs.</summary>
    Protected gpGeometryBagErreur As IGeometryBag = Nothing
    '''<summary>Interface contenant la table des statistiques d'erreurs et de traitement.</summary>
    Protected gpTableStatistiques As ITable = Nothing

#Region "Constructeurs"
    '''<summary>
    ''' Routine qui permet d'instancier une nouvelle contrainte d'intégrité à partir des valeurs nécessaires à cette dernière.
    '''</summary>
    ''' 
    '''<param name="iObjectId"> Contient le numéro d'identifiant de la contrainte.</param>
    '''<param name="sGroupe"> Contient le groupe de la contrainte d'intégrité.</param>
    '''<param name="sDescription"> Contient la description de la contrainte d'intégrité.</param>
    '''<param name="sMessage"> Contient le message de la contrainte d'intégrité.</param>
    '''<param name="sRequetes"> Contient les requêtes de la contrainte d'intégrité.</param>
    '''<param name="sNomRequete"> Contient le nom de la requête principale de la contrainte d'intégrité.</param>
    '''<param name="sNomTable"> Contient le nom de la table pour laquelle la contrainte d'intégrité est appliquée.</param>
    ''' 
    Public Sub New(ByVal iObjectId As Integer, ByVal sGroupe As String, ByVal sDescription As String, ByVal sMessage As String, _
                   ByVal sRequetes As String, ByVal sNomRequete As String, ByVal sNomTable As String)
        Try
            'Définir le numéro d'identifiant de la contrainte.
            giObjectId = iObjectId
            'Définir le groupe de la contrainte d'intégrité.
            gsGroupe = sGroupe
            'Définir la description de la contrainte d'intégrité.
            gsDescription = sDescription
            'Définir le message de la contrainte d'intégrité.
            gsMessage = sMessage
            'Définir les requêtes de la contrainte d'intégrité.
            gsRequetes = sRequetes
            'Définir le nom de la requête principale de la contrainte d'intégrité.
            gsNomRequete = NomRequete
            'Définir le nom de la table pour laquelle la contrainte d'intégrité est appliquée.
            gsNomTable = NomTable

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier une nouvelle contrainte d'intégrité à partir d'une occurence d'une table de contraintes d'intégrité.
    '''</summary>
    ''' 
    '''<param name="pRow"> Contient une occurence d'une table de contraintes d'intégrité.</param>
    '''
    Public Sub New(ByVal pRow As IRow)
        'Déclarer les variables de travail
        Dim iPosAtt As Integer = -1     'Contient la position de l'attribut de la contrainte.

        Try
            'Définir la position de l'attribut
            iPosAtt = pRow.Table.FindField("OBJECTID")
            'Vérifier si l'attribut est présent
            If iPosAtt > -1 Then
                'Définir le numéro d'identifiant de la contrainte
                giObjectId = CInt(pRow.Value(iPosAtt))
            Else
                'Définir le numéro d'identifiant de la contrainte
                giObjectId = -1
            End If

            'Définir la position de l'attribut GROUPE
            iPosAtt = pRow.Table.FindField("GROUPE")
            'Vérifier si l'attribut est présent
            If iPosAtt > -1 Then
                'Définir le groupe de la contrainte d'intégrité
                gsGroupe = pRow.Value(iPosAtt).ToString
            Else
                'Définir le groupe de la contrainte d'intégrité
                gsGroupe = Nothing
            End If

            'Définir la position de l'attribut 
            iPosAtt = pRow.Table.FindField("DESCRIPTION")
            'Vérifier si l'attribut est présent
            If iPosAtt > -1 Then
                'Définir la description de la contrainte d'intégrité.
                gsDescription = pRow.Value(iPosAtt).ToString
            Else
                'Définir la description de la contrainte d'intégrité.
                gsDescription = Nothing
            End If

            'Définir la position de l'attribut 
            iPosAtt = pRow.Table.FindField("MESSAGE")
            'Vérifier si l'attribut est présent
            If iPosAtt > -1 Then
                'Définir le message de la contrainte d'intégrité
                gsMessage = pRow.Value(iPosAtt).ToString
            Else
                'Définir le message de la contrainte d'intégrité
                gsMessage = Nothing
            End If

            'Définir la position de l'attribut 
            iPosAtt = pRow.Table.FindField("REQUETES")
            'Vérifier si l'attribut est présent
            If iPosAtt > -1 Then
                'Définir les requêtes de la contrainte d'intégrité.
                gsRequetes = pRow.Value(iPosAtt).ToString
            Else
                'Définir les requêtes de la contrainte d'intégrité.
                gsRequetes = Nothing
            End If

            'Définir la position de l'attribut 
            iPosAtt = pRow.Table.FindField("NOM_REQUETE")
            'Vérifier si l'attribut est présent
            If iPosAtt > -1 Then
                'Définir le nom de la requête principale de la contrainte d'intégrité.
                gsNomRequete = pRow.Value(iPosAtt).ToString
            Else
                'Définir le nom de la requête principale de la contrainte d'intégrité.
                gsNomRequete = Nothing
            End If

            'Définir la position de l'attribut 
            iPosAtt = pRow.Table.FindField("NOM_TABLE")
            'Vérifier si l'attribut est présent
            If iPosAtt > -1 Then
                'Définir le nom de la table pour laquelle la contrainte d'intégrité est appliquée.
                gsNomTable = pRow.Value(iPosAtt).ToString
            Else
                'Définir le nom de la table pour laquelle la contrainte d'intégrité est appliquée.
                gsNomTable = Nothing
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier une nouvelle contrainte d'intégrité à partir d'une occurence d'une table de contraintes d'intégrité.
    '''</summary>
    ''' 
    '''<param name="oTreeNode"> TreeNode de type [CONTRAINTE].</param>
    '''
    Public Sub New(ByVal oTreeNode As TreeNode)
        'Déclarer les variables de travail
        Dim oNode As TreeNode = Nothing    'TreeNode d'un TreeView de travail.

        Try
            'Vérifier si le TreeNode est de type [CONTRAINTE]
            If oTreeNode.Tag.ToString = "[CONTRAINTE]" Then
                'Définir le TreeNode de travail
                oNode = oTreeNode.Nodes.Item(0)
                'Définir le numéro d'identifiant de la contrainte
                giObjectId = CInt(oNode.Nodes.Item(0).Text)

                'Définir le TreeNode de travail
                oNode = oTreeNode.Nodes.Item(4)
                'Définir le groupe de la contrainte d'intégrité
                gsGroupe = oNode.Nodes.Item(0).Text

                'Définir le TreeNode de travail
                oNode = oTreeNode.Nodes.Item(5)
                'Définir la description de la contrainte d'intégrité.
                gsDescription = oNode.Nodes.Item(0).Text

                'Définir le TreeNode de travail
                oNode = oTreeNode.Nodes.Item(6)
                'Définir le message de la contrainte d'intégrité
                gsMessage = oNode.Nodes.Item(0).Text

                'Définir le TreeNode de travail
                oNode = oTreeNode.Nodes.Item(7)
                'Traiter tous les requêtes
                For i = 0 To oNode.Nodes.Count - 1
                    'Si c'est la première requête
                    If i = 0 Then
                        'Définir les requêtes de la contrainte d'intégrité.
                        gsRequetes = oNode.Nodes.Item(i).Text
                    Else
                        'Définir les requêtes de la contrainte d'intégrité.
                        gsRequetes = gsRequetes & vbCrLf & oNode.Nodes.Item(i).Text
                    End If
                Next

                'Définir le TreeNode de travail
                oNode = oTreeNode.Nodes.Item(8)
                'Définir le nom de la requête principale de la contrainte d'intégrité.
                gsNomRequete = oNode.Nodes.Item(0).Text

                'Définir le TreeNode de travail
                oNode = oTreeNode.Nodes.Item(9)
                'Définir le nom de la table pour laquelle la contrainte d'intégrité est appliquée.
                gsNomTable = oNode.Nodes.Item(0).Text

                'Si le TreeNode n'est pas de type [CONTRAINTE]
            Else
                'Définir le numéro d'identifiant de la contrainte
                giObjectId = -1
                'Définir le groupe de la contrainte d'intégrité
                gsGroupe = Nothing
                'Définir la description de la contrainte d'intégrité.
                gsDescription = Nothing
                'Définir le message de la contrainte d'intégrité
                gsMessage = Nothing
                'Définir les requêtes de la contrainte d'intégrité.
                gsRequetes = Nothing
                'Définir le nom de la requête principale de la contrainte d'intégrité.
                gsNomRequete = Nothing
                'Définir le nom de la table pour laquelle la contrainte d'intégrité est appliquée.
                gsNomTable = Nothing
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oNode = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de vider la mémoire.
    '''</summary>
    Protected Overrides Sub Finalize()
        'Vider la mémoire
        giObjectId = Nothing
        gsGroupe = Nothing
        gsDescription = Nothing
        gsMessage = Nothing
        gsRequetes = Nothing
        gsNomRequete = Nothing
        gsNomTable = Nothing
        gsIdentifiant = Nothing
        gsNomClasseDecoupage = Nothing
        gsNomRepertoireErreurs = Nothing
        gsNomFichierJournal = Nothing
        gpTrackCancel = Nothing
        gpSpatialReference = Nothing
        gpQueryFilter = Nothing
        gpGeodatabase = Nothing
        gpFeatureDecoupage = Nothing
        gpFeatureLayerDecoupage = Nothing
        gpFeatureLayerSelection = Nothing
        gpFeatureLayerErreur = Nothing
        gpGeometryBagErreur = Nothing
        gpTableStatistiques = Nothing
        gsInformation = Nothing
        giNombreElements = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
        'Finaliser
        MyBase.Finalize()
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de définir et retourner l'identifiant d'éxécution du traitement.
    '''</summary>
    ''' 
    Public Property Identifiant() As String
        Get
            Identifiant = gsIdentifiant
        End Get
        Set(ByVal value As String)
            gsIdentifiant = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'entête d'éxécution du traitement.
    '''</summary>
    ''' 
    Public Property Entete() As String
        Get
            Entete = gsEntete
        End Get
        Set(ByVal value As String)
            gsEntete = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom du répertoire ou de la Géodatabase d'erreurs.
    '''</summary>
    ''' 
    Public Property NomRepertoireErreurs() As String
        Get
            NomRepertoireErreurs = gsNomRepertoireErreurs
        End Get
        Set(ByVal value As String)
            gsNomRepertoireErreurs = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom du fichier journal d'éxécution du traitement.
    '''</summary>
    ''' 
    Public Property NomFichierJournal() As String
        Get
            NomFichierJournal = gsNomFichierJournal
        End Get
        Set(ByVal value As String)
            gsNomFichierJournal = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le RichTextBox utilisé pour écrire le journal d'éxécution du traitement.
    '''</summary>
    ''' 
    Public Property RichTextbox() As RichTextBox
        Get
            RichTextbox = gqRichTextBox
        End Get
        Set(ByVal value As RichTextBox)
            gqRichTextBox = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le TrackCancel pour l'exécution de la contrainte d'intégrité.
    '''</summary>
    ''' 
    Public Property TrackCancel() As ITrackCancel
        Get
            TrackCancel = gpTrackCancel
        End Get
        Set(ByVal value As ITrackCancel)
            gpTrackCancel = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le référence spatiale pour l'exécution de la contrainte d'intégrité.
    '''</summary>
    ''' 
    Public Property SpatialReference() As ISpatialReference
        Get
            SpatialReference = gpSpatialReference
        End Get
        Set(ByVal value As ISpatialReference)
            gpSpatialReference = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le Filtre  pour l'exécution de la contrainte d'intégrité.
    '''</summary>
    ''' 
    Public Property QueryFilter() As IQueryFilter
        Get
            QueryFilter = gpQueryFilter
        End Get
        Set(ByVal value As IQueryFilter)
            gpQueryFilter = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner l'identifiant de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public ReadOnly Property ObjectId() As Integer
        Get
            ObjectId = giObjectId
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le groupe de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Property Groupe() As String
        Get
            Groupe = gsGroupe
        End Get
        Set(ByVal value As String)
            gsGroupe = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la description de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Property Description() As String
        Get
            Description = gsDescription
        End Get
        Set(ByVal value As String)
            gsDescription = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le message de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Property Message() As String
        Get
            Message = gsMessage
        End Get
        Set(ByVal value As String)
            gsMessage = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner les requêtes de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Property Requetes() As String
        Get
            Requetes = gsRequetes
        End Get
        Set(ByVal value As String)
            gsRequetes = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre de requêtes présentes dans la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreRequetes() As Integer
        Get
            'Déclarer les variables de travail
            Dim sRequete() As String = Nothing

            'Créer la liste des requêtes
            sRequete = Split(gsRequetes, vbCrLf)

            'Retourner le nombre de requêtes
            NombreRequetes = sRequete.Length

            'Vider la mémoire
            sRequete = Nothing
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre d'erreurs présentes dans la contrainte d'intégrité traité.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreErreurs() As Long
        Get
            'Vérifier si les géométries d'erreurs sont présentes
            If gpGeometryBagErreur Is Nothing Then
                'Retourner le nombre d'erreurs
                NombreErreurs = -1
            Else
                'Déclarer les variables de travail
                Dim pGeometryColl As IGeometryCollection = Nothing      'Interface pour extraire le nombre de géométries en erreur.

                'Créer la liste des requêtes
                pGeometryColl = CType(gpGeometryBagErreur, IGeometryCollection)

                'Retourner le nombre d'erreurs
                NombreErreurs = pGeometryColl.GeometryCount

                'Vider la mémoire
                pGeometryColl = Nothing
            End If
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre d'éléments sélectionnés.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreSelection() As Long
        Get
            'Déclarer les variable de travail
            Dim pFeatureSelection As IFeatureSelection = Nothing    'Interface pour extraire le nombre d'éléments sélectionnés

            'Interface pour extraire le nombre d'éléments sélectionnés
            pFeatureSelection = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Retourner le nombre d'éléments sélectionnés
            NombreSelection = pFeatureSelection.SelectionSet.Count

            'Vider la mémoire
            pFeatureSelection = Nothing
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre d'éléments traités dans la contrainte d'intégrité traité.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreElements() As Long
        Get
            'Retourner le nombre d'éléments
            NombreElements = giNombreElements
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner une requête selon une position dans la liste des requêtes de la contrainte d'intégrité à traiter.
    ''' 
    ''' La position de la requête dans la liste soit se situer entre 0 et le nombre de requêtes présentes moins 1, 
    ''' sinon nothing sera retourné ou rien ne sera modifié.
    '''</summary>
    ''' 
    Public Property Requete(ByVal iPosition As Integer) As String
        Get
            'Définir la valeur par défaut
            Requete = Nothing

            'Vérifier si les requêtes sont spécifiées
            If gsRequetes IsNot Nothing Then
                'Déclarer les variables de travail
                Dim sRequete() As String = Nothing

                'Créer la liste des requêtes
                sRequete = Split(gsRequetes, vbCrLf)

                'Vérifier si la position de la requête est valide
                If iPosition >= 0 And iPosition < sRequete.Length Then
                    'Retourner la requête correspondant à la position demandée
                    Requete = sRequete(iPosition)
                End If

                'Vider la mémoire
                sRequete = Nothing
            End If
        End Get
        Set(ByVal value As String)
            'Vérifier si les requêtes sont spécifiées
            If gsRequetes IsNot Nothing Then
                'Déclarer les variables de travail
                Dim sRequete() As String = Nothing

                'Créer la liste des requêtes
                sRequete = Split(gsRequetes, vbCrLf)

                'Vérifier si la position de la requête est valide
                If iPosition >= 0 And iPosition < sRequete.Length Then
                    'Modifier la requête correspondant à la position demandée
                    sRequete(iPosition) = value
                    'Traiter toutes les requêtes
                    For i = 0 To sRequete.Length - 1
                        'Vérifier si la position est 0
                        If i = 0 Then
                            'Redéfinir la liste des requêtes
                            gsRequetes = sRequete(i)
                        Else
                            'Redéfinir la liste des requêtes
                            gsRequetes = gsRequetes & vbCrLf & sRequete(i)
                        End If
                    Next
                End If

                'Vider la mémoire
                sRequete = Nothing
            End If
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom de la requête principale la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Property NomRequete() As String
        Get
            NomRequete = gsNomRequete
        End Get
        Set(ByVal value As String)
            gsNomRequete = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom de la table pour laquelle la contrainte d'intégrité est traitée.
    '''</summary>
    ''' 
    Public Property NomTable() As String
        Get
            NomTable = gsNomTable
        End Get
        Set(ByVal value As String)
            gsNomTable = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la Géodatabase contenant les classes des éléments à valider.
    '''</summary>
    ''' 
    Public Property Geodatabase() As IFeatureWorkspace
        Get
            Geodatabase = gpGeodatabase
        End Get
        Set(ByVal value As IFeatureWorkspace)
            gpGeodatabase = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom de la classe de découpage à traiter.
    '''</summary>
    ''' 
    Public Property NomClasseDecoupage() As String
        Get
            NomClasseDecoupage = gsNomClasseDecoupage
        End Get
        Set(ByVal value As String)
            'Définir le nom de la classe de découpage
            gsNomClasseDecoupage = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'élément de découpage à traiter.
    '''</summary>
    ''' 
    Public Property FeatureDecoupage() As IFeature
        Get
            'Conserver l'élément de découpage
            FeatureDecoupage = gpFeatureDecoupage
        End Get
        Set(ByVal value As IFeature)
            'Définir l'élément de découpage
            gpFeatureDecoupage = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le FeatureLayer contenant les éléments de découpage à traiter.
    '''</summary>
    ''' 
    Public Property FeatureLayerDecoupage() As IFeatureLayer
        Get
            FeatureLayerDecoupage = gpFeatureLayerDecoupage
        End Get
        Set(ByVal value As IFeatureLayer)
            'Définir le FeatureLayer de découpage
            gpFeatureLayerDecoupage = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le FeatureLayer contenant les éléments sélectionnés.
    '''</summary>
    ''' 
    Public ReadOnly Property FeatureLayerSelection() As IFeatureLayer
        Get
            FeatureLayerSelection = gpFeatureLayerSelection
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le FeatureLayer contenant les éléments décrivant les erreurs.
    '''</summary>
    ''' 
    Public ReadOnly Property FeatureLayerErreur() As IFeatureLayer
        Get
            FeatureLayerErreur = gpFeatureLayerErreur
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le géométrie Bag contenant les géométries décrivant les erreurs.
    '''</summary>
    ''' 
    Public ReadOnly Property GeometryBagErreur() As IGeometryBag
        Get
            GeometryBagErreur = gpGeometryBagErreur
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner la table des statistiques d'erreurs et de traitement.
    '''</summary>
    ''' 
    Public Property TableStatistiques() As ITable
        Get
            TableStatistiques = gpTableStatistiques
        End Get
        Set(ByVal value As ITable)
            'Définir la table des statistiques
            gpTableStatistiques = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner l'information de la contrainte d'intégrité.
    '''</summary>
    ''' 
    Public ReadOnly Property Information() As String
        Get
            Information = gsInformation
        End Get
    End Property
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    ''' Routine qui permet d'exécuter la contrainte d'intégrité et retourner le résultat obtenu.
    '''</summary>
    ''' 
    '''<returns>True si la contrainte est valide, False sinon.</returns>
    ''' 
    Public Function EstValide() As Boolean
        'Déclaration des variables de travail

        'Par défaut, la contrainte est valide
        EstValide = True
        gsInformation = ""

        Try
            'Vérifier si le groupe est invalide
            If gsGroupe.Length <= 0 Or gsGroupe.Length > 50 Then
                'Définir le message d'erreur
                EstValide = False
                gsInformation = "ERREUR : La longueur du groupe est invalide : " & gsGroupe.Length.ToString & vbCrLf

                'Vérifier si la description est invalide
            ElseIf gsDescription.Length <= 0 Or gsDescription.Length > 2000 Then
                'Définir le message d'erreur
                EstValide = False
                gsInformation = "ERREUR : La longueur de la description est invalide : " & gsDescription.Length.ToString & vbCrLf

                'Vérifier si le message est invalide
            ElseIf gsMessage.Length <= 0 Or gsMessage.Length > 2000 Then
                'Définir le message d'erreur
                EstValide = False
                gsInformation = "ERREUR : La longueur du message est invalide : " & gsMessage.Length.ToString & vbCrLf

                'Vérifier si le nom de la requête est invalide
            ElseIf gsNomRequete.Length <= 0 Or gsNomRequete.Length > 50 Then
                'Définir le message d'erreur
                EstValide = False
                gsInformation = "ERREUR : La longueur du nom de la requête est invalide : " & gsNomRequete.Length.ToString & vbCrLf

                'Vérifier si le nom de la table est invalide
            ElseIf gsNomTable.Length <= 0 Or gsNomTable.Length > 50 Then
                'Définir le message d'erreur
                EstValide = False
                gsInformation = "ERREUR : La longueur du nom de la table est invalide : " & gsNomTable.Length.ToString & vbCrLf
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'exécuter la contrainte d'intégrité et retourner le résultat obtenu.
    '''</summary>
    '''
    '''<returns>IGeometry contenant les géométries d'erreurs.</returns>
    ''' 
    Public Function Executer() As IGeometry
        'Déclaration des variables de travail
        Dim oRequete As clsRequete = Nothing                        'Object contenant une requête à exécuter.
        Dim qLayersColl As Collection = Nothing                     'Collection des Layers traités et conservés.
        Dim pGeometry As IGeometry = Nothing                        'Interface contenant les géométries d'erreurs
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim oProcess As Process = Nothing                           'Objet contenant l'information sur la mémoire.
        Dim dDateDebut As DateTime = Nothing                        'Contient la date de début du traitement.
        Dim sTempsTraitement As String = ""                         'Temps de traitement.
        Dim sNomUsager As String = ""                               'Nom de l'usager.

        'Par défaut, un GeometryBag vide est retourné
        Executer = Nothing

        Try
            'Initialiser le résultat
            giNombreElements = 0
            gpGeometryBagErreur = CType(Executer, IGeometryBag)
            sNomUsager = System.Environment.GetEnvironmentVariable("USERNAME")
            If Len(sNomUsager) = 0 Then sNomUsager = System.Environment.GetEnvironmentVariable("LSFUSER")
            If Len(sNomUsager) = 0 Then sNomUsager = "INCONNU"

            'Vérifier si la contrainte est valide
            If EstValide() Then
                'Vérifier si la géodatabase est valide
                If gpGeodatabase IsNot Nothing Then
                    'Définir la date de début
                    dDateDebut = System.DateTime.Now

                    'Interface pour extraire la référence spatiale
                    pSpatialRefFact = New SpatialReferenceEnvironment

                    'Vérifier si la requete est AdjacenceGeometrie, AjustementDecoupage, DecoupageElement, GeometrieValide, NombreIntersection ou RelationSpatiale
                    If Me.gsNomRequete = "AdjacenceGeometrie" Or Me.gsNomRequete = "AjustementDecoupage" Or Me.gsNomRequete = "DecoupageElement" Or Me.gsNomRequete = "GeometrieValide" Or Me.gsNomRequete = "NombreIntersection" Or Me.gsNomRequete = "RelationSpatiale" Or Me.gsNomRequete = "DistanceGeometrie" Then
                        'Définir la référence spatiale GCS_North_American_1983_CSRS:4617
                        gpSpatialReference = pSpatialRefFact.CreateSpatialReference(4617)
                        'Définir la tolérance XY
                        pSpatialRefTol = CType(gpSpatialReference, ISpatialReferenceTolerance)
                        pSpatialRefRes = CType(gpSpatialReference, ISpatialReferenceResolution)
                        pSpatialRefRes.SetDefaultXYResolution()
                        pSpatialRefTol.XYTolerance = 0.00000001
                    Else
                        'Définir la référence spatiale LCC NAD83 CSRS:3979
                        gpSpatialReference = pSpatialRefFact.CreateSpatialReference(3979)
                        'Définir la tolérance XY
                        pSpatialRefTol = CType(gpSpatialReference, ISpatialReferenceTolerance)
                        pSpatialRefRes = CType(gpSpatialReference, ISpatialReferenceResolution)
                        pSpatialRefRes.SetDefaultXYResolution()
                        pSpatialRefTol.XYTolerance = 0.001
                    End If

                    'Afficher l'entête de l'exécution de la contrainte
                    EcrireMessage("--------------------------------------------------------------------------------")
                    EcrireMessage("-" & Me.Entete & ", OID=" & Me.ObjectId & ", " & Me.NomTable & ", " & Me.Groupe)
                    EcrireMessage("-" & Me.Description)
                    EcrireMessage("-" & Me.Message)
                    EcrireMessage("--------------------------------------------------------------------------------")

                    'Créer une nouvelle collection de Layers
                    qLayersColl = New Collection

                    'Traiter toutes les requêtes de la contrainte
                    For Each sCommande In Split(gsRequetes, vbCrLf)
                        'Vérifier si la requête est "ConserverLayer"
                        If sCommande.Contains("ConserverLayer") Then
                            'Exécuter le traitement pour conserver le dernier FeatureLayer d'erreurs ou de sélection
                            Call ConserverLayer(sCommande, qLayersColl)

                            'Vérifier si la requête est "LireMemoire"
                        ElseIf sCommande.Contains("LireMemoire") Then
                            'Exécuter le traitement pour conserver le dernier FeatureLayer d'erreurs ou de sélection
                            Call LireMemoire(sCommande, qLayersColl)

                            'Si on doit exécuter la requête
                        ElseIf sCommande.Length > 0 Then
                            'Définir la requête à exécuter sous forme de commande
                            oRequete = New clsRequete(sCommande, qLayersColl, gpGeodatabase, gpSpatialReference, gpQueryFilter, _
                                                      gsNomClasseDecoupage, gpFeatureLayerDecoupage, gpFeatureDecoupage)

                            'Définir l'objet pour permettre d'annuler le traitement
                            oRequete.TrackCancel = gpTrackCancel
                            'Définir l'object pour l'affichage de l'exécution en interactif
                            oRequete.RichTextbox = gqRichTextBox
                            'Définir le nom du fichier journal
                            oRequete.NomFichierJournal = gsNomFichierJournal

                            'Sélectionner les éléments en erreurs et retourner le GeometryBag d'erreurs
                            pGeometry = oRequete.Executer()

                            'Conserver l'information de la requête dans celle de la contrainte
                            gsInformation = gsInformation & oRequete.Information

                            'Définir le nombre d'éléments traités
                            If giNombreElements = 0 Then giNombreElements = oRequete.NombreElementsDebut

                            'Définir le FeatureLayer d'erreurs
                            gpFeatureLayerErreur = oRequete.Requete.FeatureLayerErreur

                            'Définir le FeatureLayer de sélection
                            gpFeatureLayerSelection = oRequete.Requete.FeatureLayerSelection

                            'Récupération de la mémoire disponible
                            oRequete = Nothing
                            GC.Collect()
                        End If
                    Next

                    'Définir le GeometryBag d'erreurs
                    Executer = pGeometry
                    gpGeometryBagErreur = CType(Executer, IGeometryBag)

                    'Écrire les erreurs au besoin
                    Me.EcrireErreur()

                    'Définir l'objet pour extraire la mémoire utilisée
                    oProcess = Process.GetCurrentProcess()
                    'Afficher le résultat de l'exécution de la requête
                    sTempsTraitement = System.DateTime.Now.Subtract(dDateDebut).ToString
                    EcrireMessage("Temps d'exécution: " & sTempsTraitement & ", (PeakWorkingSet: " & (oProcess.PeakWorkingSet64 / 1024).ToString("N0") & " K)")
                    EcrireMessage("Nombre d'erreurs trouvées: " & Me.NombreErreurs.ToString & "/" & Me.NombreElements)
                    EcrireMessage("")

                    'Écrire les statistiques
                    EcrireStatistiques(gpTableStatistiques, sNomUsager, Me.Identifiant, Me.NomTable, Me.Groupe, _
                                       Me.NombreElements, Me.NombreSelection, Me.NombreErreurs, sTempsTraitement)

                Else
                    'Définir le message d'erreur
                    gsInformation = "ERREUR : La géodatabase est invalide!" & vbCrLf
                End If
            End If

        Catch ex As CancelException
            'Écrire l'erreur
            EcrireMessage("[clsContrainte.Executer] " & ex.Message & vbCrLf)
            'Retourner l'exception
            Throw
        Catch ex As Exception
            'Écrire l'erreur - System.Environment.StackTrace
            EcrireMessage("[clsContrainte.Executer] " & ex.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oRequete = Nothing
            qLayersColl = Nothing
            pGeometry = Nothing
            pSpatialRefFact = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            oProcess = Nothing
            dDateDebut = Nothing
            sTempsTraitement = Nothing
            sNomUsager = Nothing
            'Récupération de la mémoire disponible
            GC.Collect()
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    '''Routine qui permet d'écrire les erreurs sur disque dans une FeatureClass d'une Geodatabase ou dans un FeatureLayer de sélection.
    '''</summary>
    '''
    '''<param name="sAttributDecoupage">Nom de l'attribut de découpage.</param>
    '''<param name="sValeurDecoupage">Nom de la valeur de découpage.</param>
    ''' 
    ''' <remarks>
    ''' Si le nom du répertoire contient [DATE_TIME], il sera remplacé par la date et l'heure courante.
    ''' Si le nom du répertoire contient [sAttributDecoupage], il sera remplacé par sValeurDecoupage.
    '''</remarks>
    ''' 
    Private Sub EcrireErreur(Optional ByVal sAttributDecoupage As String = "", Optional ByVal sValeurDecoupage As String = "")
        'Déclaration des variables de travail
        Dim pWorkspaceErr As IWorkspace = Nothing   'Interface utilisé pour ouvrir une classe spatiale d'erreurs.
        Dim pLayerFile As ILayerFile = Nothing      'Interface qui permet de lire un FeatureLayer sur disque.
        Dim sNomWorkspaceErr As String = Nothing    'Contient le nom du Workspace (Géodatabase) d'erreurs à crééer.
        Dim sNomClasseErr As String = Nothing       'Contient le nom de la classe d'erreurs à créer.

        Try
            'Vérifier si des erreurs sont présentes
            If Me.NombreErreurs > 0 Then
                'Vérifier si un nom de répertoire ou de géodatabase d'erreurs est spécifié
                If gsNomRepertoireErreurs.Length > 0 Then
                    'Définir le nom du répertoire ou de la Géodatabase d'erreurs
                    sNomWorkspaceErr = gsNomRepertoireErreurs
                    'Remplacer la valeur [DATE_TIME] par la date et l'heure actuelle
                    sNomWorkspaceErr = sNomWorkspaceErr.Replace("[DATE_TIME]", DateTime.Now.ToString("yyyyMMdd_HHmmss"))
                    'Remplacer la valeur [sAttributDecoupage] par la valeur de découpage
                    sNomWorkspaceErr = sNomWorkspaceErr.Replace("[" & sAttributDecoupage & "]", sValeurDecoupage)

                    'Créer le répertoire ou la Géodatabase d'erreurs
                    pWorkspaceErr = CreerGeodatabaseErreurs(sNomWorkspaceErr)

                    'Vérifier si on doit écrire les erreurs dans une classe
                    If pWorkspaceErr IsNot Nothing Then
                        'Définir le nom de la classe d'erreurs à créer à partir du nom de la classe, du numéro de contrainte et du groupe traité.
                        sNomClasseErr = Me.NomTable & "_" & Me.ObjectId.ToString & "_" & Me.Groupe
                        'Vérifier si le nom contient un nom d'usager
                        If sNomClasseErr.Contains(".") Then
                            'Enlever le nom de l'usager dans le nom
                            sNomClasseErr = sNomClasseErr.Split(CChar("."))(1)
                        End If
                        'Vérifier si l'identifiant est présent
                        If Me.Identifiant.Length > 0 And sNomWorkspaceErr.Contains(Me.Identifiant) = False Then
                            'Définir le nom de la classe d'erreurs à créer à partir de l'identifiant, du nom de la classe, du numéro de contrainte et du groupe traité.
                            sNomClasseErr = "I" & Me.Identifiant & "_" & sNomClasseErr
                        End If
                        'Afficher le nom de la classe d'erreurs
                        EcrireMessage("  Créer la classe d'erreurs : " & sNomClasseErr)

                        'Écrire la classe d'erreurs sur disque
                        ConvertFeatureClass(gpFeatureLayerErreur.FeatureClass, CType(pWorkspaceErr, IWorkspace2), sNomClasseErr)

                        'Sinon on  écrit seulement les Layers d'erreurs
                    Else
                        'Interface pour ouvrir le FeatureLayer d'erreurs
                        pLayerFile = New LayerFile
                        'Définir le nom du FeatureLayer à partir du nom de la table et le groupe
                        gpFeatureLayerSelection.Name = Me.NomTable & "_" & Me.ObjectId.ToString & "_" & Me.Groupe
                        'Créer le FeatureLayer des erreurs vide sur disque
                        pLayerFile.New(sNomWorkspaceErr & "\" & gpFeatureLayerSelection.Name & ".lyr")
                        'Changer le contenu du FeatureLayer avec les erreurs sélectionnées
                        pLayerFile.ReplaceContents(CreerLayerSelection(gpFeatureLayerSelection, gpFeatureLayerSelection.Name))
                        'Sauver le FeatureLayer sur disque
                        pLayerFile.Save()
                        'Afficher le nom du FeatureLayer contenant la sélection des éléments en erreur
                        EcrireMessage("  Créer le Layer d'erreurs : " & gpFeatureLayerSelection.Name)
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pWorkspaceErr = Nothing
            pLayerFile = Nothing
            sNomWorkspaceErr = Nothing
            sNomClasseErr = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet d'écrire les statistiques d'erreurs et de traitement dans la table des statistiques spécifiée.
    '''</summary>
    '''
    '''<param name="pTableStatistiques">Nom de l'identifiant traité.</param>
    '''<param name="sNomUsager">Nom de l'usager.</param>
    '''<param name="sIdentifiant">Nom de l'identifiant traité.</param>
    '''<param name="sNomTable">Nom de la table traité.</param>
    '''<param name="sContrainte">Nom de la contrainte.</param>
    '''<param name="iNbElements">Nombre d'éléments traités.</param>
    '''<param name="iNbSelection">Nombre d'éléments sélectionnés.</param>
    '''<param name="iNbErreurs">Nombre d'erreurs trouvées.</param>
    '''<param name="sTempsTraitement">Temps de traitement utilisé.</param>
    ''' 
    Private Sub EcrireStatistiques(ByRef pTableStatistiques As ITable, ByVal sNomUsager As String, ByVal sIdentifiant As String, ByVal sNomTable As String, ByVal sContrainte As String, _
                                   ByVal iNbElements As Long, ByVal iNbSelection As Long, ByVal iNbErreurs As Long, ByVal sTempsTraitement As String)
        'Déclaration des variables de travail
        Dim pQueryFilter As IQueryFilter = Nothing  'Interface utilisé pour effectuer une requête attributive.
        Dim pCursor As ICursor = Nothing            'Interface utilisé pour extraire une statistique.
        Dim pRowUpdate As IRow = Nothing            'Interface ESRI contenant un item de la table des statistiques.
        Dim pRowCreate As IRowBuffer = Nothing      'Interface ESRI contenant un item de la table des statistiques.

        Try
            'Vérifier si la table des statistiques est spécifiée
            If pTableStatistiques IsNot Nothing Then
                'Interface utilisé pour effectuer une requête attributive
                pQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = "IDENTIFIANT='" & sIdentifiant & "' AND NOM_TABLE='" & sNomTable & "' AND CONTRAINTE='" & sContrainte & "'"

                'Interface utilisé pour extraire une statistique
                pCursor = pTableStatistiques.Update(pQueryFilter, False)

                'Définir la statistique existante
                pRowUpdate = pCursor.NextRow

                'Vérifier si la statistiques est inexistante
                If pRowUpdate Is Nothing Then
                    'Créer le curseur pour insérer des statistiques
                    pCursor = pTableStatistiques.Insert(True)

                    Try
                        'Créer une nouvelle statistique
                        pRowCreate = pTableStatistiques.CreateRowBuffer
                    Catch ex As Exception
                        'Envoyer le message de non-écriture des statistiques
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Impossible d'écrire la statistique : NbElem=" & iNbElements.ToString & ",NbSelect:" & iNbSelection.ToString & ",NbErr:" & iNbErreurs.ToString & ",Temps:" & sTempsTraitement.Substring(0, 10))
                    End Try

                    'Définir la statistique
                    pRowCreate.Value(1) = sNomUsager
                    pRowCreate.Value(2) = System.DateTime.Now
                    pRowCreate.Value(3) = System.DateTime.Now
                    pRowCreate.Value(4) = sIdentifiant
                    pRowCreate.Value(5) = sNomTable
                    pRowCreate.Value(6) = sContrainte
                    pRowCreate.Value(7) = iNbElements
                    pRowCreate.Value(8) = iNbSelection
                    pRowCreate.Value(9) = iNbErreurs
                    pRowCreate.Value(10) = sTempsTraitement.Substring(0, 10)

                    Try
                        'Insérer la statistique
                        pCursor.InsertRow(pRowCreate)
                    Catch ex As Exception
                        'Envoyer le message de non-écriture des statistiques
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Impossible de créer la statistique : NbElem=" & iNbElements.ToString & ",NbSelect:" & iNbSelection.ToString & ",NbErr:" & iNbErreurs.ToString & ",Temps:" & sTempsTraitement.Substring(0, 10))
                    End Try

                    'Écrire l'information de la statistique
                    pCursor.Flush()

                    'Si la statistique existe déjà
                Else
                    'Définir la statistique
                    pRowUpdate.Value(1) = sNomUsager
                    pRowUpdate.Value(3) = System.DateTime.Now
                    pRowUpdate.Value(7) = iNbElements
                    pRowUpdate.Value(8) = iNbSelection
                    pRowUpdate.Value(9) = iNbErreurs
                    pRowUpdate.Value(10) = sTempsTraitement.Substring(0, 10)

                    Try
                        'Mettre à jour la statistique
                        pCursor.UpdateRow(pRowUpdate)
                    Catch ex As Exception
                        'Envoyer le message de non-écriture des statistiques
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Impossible de modifier la statistique : NbElem=" & iNbElements.ToString & ",NbSelect:" & iNbSelection.ToString & ",NbErr:" & iNbErreurs.ToString & ",Temps:" & sTempsTraitement.Substring(0, 10))
                    End Try
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Release the update cursor to remove the lock on the input data.
            If pCursor IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pCursor)
            'Vider la mémoire
            pQueryFilter = Nothing
            pCursor = Nothing
            pRowUpdate = Nothing
            pRowCreate = Nothing
            'Récupération de la mémoire disponible
            GC.Collect()
        End Try
    End Sub

        '''<summary>
        '''Routine qui permet de créer un répertoire ou une Géodatabase dans lequel les erreurs seront écrites.
        '''</summary>
        '''
        '''<param name="sNomGeodatabaseErreurs">Nom du répertoire ou de la Géodatabase dans lequel les erreurs seront écrites.</param>
        ''' 
        '''<returns>"IWorkspace" correspondants à une Géodatabase, Nothing si c'est un répertoire qui est créé.</returns>
        ''' 
        ''' <remarks>
        ''' Une Géodatabase est créée seulement si le nom contient .mdb ou .gdb, sinon un répertoire sera créé.
        '''</remarks>
        '''
    Private Function CreerGeodatabaseErreurs(ByVal sNomGeodatabaseErreurs As String) As IWorkspace
        'Déclaration des variables de travail
        Dim pFactoryType As Type = Nothing                      'Interface utilisé pour définir le type de Géodatabase.
        Dim pWorkspaceFactory As IWorkspaceFactory = Nothing    'Interface utilise pour créer une Géodatabase.

        'Par défaut, aucune géodatabase d'erreurs n'est créée
        CreerGeodatabaseErreurs = Nothing

        Try
            'Vérifier si un répertoire d'erreurs est spécifié
            If sNomGeodatabaseErreurs.Length > 0 Then
                'Vérifier si le répertoire d'erreurs est une personnel geodatabase
                If sNomGeodatabaseErreurs.Contains(".mdb") Then
                    'Définir le type de workspace : MDB
                    pFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.AccessWorkspaceFactory")
                    'Interface pour ouvrir le Workspace d'erreurs
                    pWorkspaceFactory = CType(Activator.CreateInstance(pFactoryType), IWorkspaceFactory)

                    'Vérifier si le Workspace d'erreurs n'existe pas
                    If Not File.Exists(sNomGeodatabaseErreurs) Then
                        'Afficher le nom de la classe d'erreurs
                        EcrireMessage("  Créer la Géodatabase d'erreurs : " & sNomGeodatabaseErreurs)
                        'Créer le workspace d'erreurs
                        pWorkspaceFactory.Create(IO.Path.GetDirectoryName(sNomGeodatabaseErreurs), IO.Path.GetFileName(sNomGeodatabaseErreurs), Nothing, 0)
                    End If

                    'Ouvrir la géodatabase d'erreurs
                    CreerGeodatabaseErreurs = pWorkspaceFactory.OpenFromFile(sNomGeodatabaseErreurs, 0)

                    'Si le répertoire d'erreurs est une file geodatabase
                ElseIf sNomGeodatabaseErreurs.Contains(".gdb") Then
                    'Définir le type de workspace : GDB
                    pFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory")
                    'Interface pour ouvrir le Workspace d'erreurs
                    pWorkspaceFactory = CType(Activator.CreateInstance(pFactoryType), IWorkspaceFactory)

                    'Vérifier si le Workspace d'erreurs n'existe pas
                    If Not System.IO.Directory.Exists(sNomGeodatabaseErreurs) Then
                        'Afficher le nom de la classe d'erreurs
                        EcrireMessage("  Créer la Géodatabase d'erreurs : " & sNomGeodatabaseErreurs)
                        'Créer le workspace d'erreurs
                        pWorkspaceFactory.Create(IO.Path.GetDirectoryName(sNomGeodatabaseErreurs), IO.Path.GetFileName(sNomGeodatabaseErreurs), Nothing, 0)
                    End If

                    'Ouvrir la géodatabase d'erreurs
                    CreerGeodatabaseErreurs = pWorkspaceFactory.OpenFromFile(sNomGeodatabaseErreurs, 0)

                    'Si le répertoire d'erreurs n'existe pas
                ElseIf Not System.IO.Directory.Exists(sNomGeodatabaseErreurs) Then
                    'Afficher le nom de la classe d'erreurs
                    EcrireMessage("  Créer le Répertoire d'erreurs : " & sNomGeodatabaseErreurs)
                    'Créer le répertoire d'erreurs
                    System.IO.Directory.CreateDirectory(sNomGeodatabaseErreurs)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pFactoryType = Nothing
            pWorkspaceFactory = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de créer un FeatureLayer de sélection.
    '''</summary>
    '''
    '''<param name="pFeatureLayer">FeatureLayer contenant la sélection d'éléments.</param>
    '''<param name="sNomLayer">Nom du Layer à créer.</param>
    ''' 
    '''<returns>"IFeatureLayer" dont seuls les éléments d'un SelectionSet sont présents dans le Layer.</returns>
    ''' 
    ''' <remarks>
    ''' Si aucun élément n'est sélectionné, rien n'est effectué.
    '''</remarks>
    Private Function CreerLayerSelection(ByVal pFeatureLayer As IFeatureLayer, ByVal sNomLayer As String) As IFeatureLayer
        'Déclarer les variables de travail
        Dim pFeatureSelection As IFeatureSelection = Nothing    'Interface qui permet de traiter la sélection du Layer
        Dim pSelectionSet As ISelectionSet = Nothing            'Interface contenant la sélection du Layer

        Dim pGeoFeatureLayer As IGeoFeatureLayer = Nothing      'Interface contenant les paramètres d'affichage d'un Layer
        Dim pDisplayString As IDisplayString = Nothing          'Interface utilisé pour extraire l'information du display
        Dim pFLDef As IFeatureLayerDefinition = Nothing         'Interface utilisé pour créer un Layer de sélection

        Dim pNewFeatureLayer As IFeatureLayer = Nothing         'Interface contenant le nouveau FeatureLayer
        Dim pNewGeoFeatureLayer As IGeoFeatureLayer = Nothing   'Interface contenant les nouveaux paramètres d'affichage d'un Layer
        Dim pNewDisplayString As IDisplayString = Nothing       'Interface utilisé pour remettre l'information du display
        Dim pNewFLDef As IFeatureLayerDefinition = Nothing      'Interface utilisé pour ajouter une recherche attributive

        'Initialiser lae FeatureLayer de retour
        CreerLayerSelection = Nothing

        Try
            'Interface pour traiter la sélection
            pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
            pSelectionSet = pFeatureSelection.SelectionSet

            'Vérifier si au moins un élément est sélectionné
            If pSelectionSet.Count > 0 Then
                'Interface pour créer un FeatureLayer selon la sélection
                pFLDef = CType(pFeatureLayer, IFeatureLayerDefinition)
                pNewFeatureLayer = pFLDef.CreateSelectionLayer(sNomLayer, True, "", "")
                pNewFeatureLayer.DisplayField = pFeatureLayer.DisplayField
                pNewFLDef = CType(pNewFeatureLayer, IFeatureLayerDefinition)
                pNewFLDef.DefinitionExpression = pFLDef.DefinitionExpression

                'Remettre l'information du display
                pDisplayString = CType(pFeatureLayer, IDisplayString)
                pNewDisplayString = CType(pNewFeatureLayer, IDisplayString)
                pNewDisplayString.ExpressionProperties = pDisplayString.ExpressionProperties

                'Conserver la représentation graphique
                pGeoFeatureLayer = CType(pFeatureLayer, IGeoFeatureLayer)
                pNewGeoFeatureLayer = CType(pNewFeatureLayer, IGeoFeatureLayer)
                pNewGeoFeatureLayer.Renderer = pGeoFeatureLayer.Renderer
                pNewGeoFeatureLayer.AnnotationProperties = pGeoFeatureLayer.AnnotationProperties
                pNewGeoFeatureLayer.DisplayAnnotation = pGeoFeatureLayer.DisplayAnnotation

                'Remettre la sélection dans le nouveau layer
                pFeatureSelection = CType(pNewFeatureLayer, IFeatureSelection)
                pFeatureSelection.SelectionSet = pSelectionSet

                'Ajouter le Layer visible
                pNewFeatureLayer.Visible = False

                'Définir le FeatureLayer de retour
                CreerLayerSelection = pNewFeatureLayer
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureSelection = Nothing
            pSelectionSet = Nothing
            pGeoFeatureLayer = Nothing
            pDisplayString = Nothing
            pFLDef = Nothing
            pNewFeatureLayer = Nothing
            pNewGeoFeatureLayer = Nothing
            pNewDisplayString = Nothing
            pNewFLDef = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de convertir une FeatureClass d'un format vers un autre format.
    '''
    '''<param name="pSourceClass">Interface contenant la featureClass de départ.</param>
    '''<param name="pTargetWorkspace">Interface contenant la géodatabase d'arrivée.</param>
    '''<param name="sTargetName">Nom de la featureClass d'arrivée.</param>
    ''' 
    ''' <return>IEnumInvalidObject contenant les éléments invalides qui n'ont pas été transférés.</return>
    ''' 
    '''</summary>
    '''
    Private Function ConvertFeatureClass(ByVal pSourceClass As IFeatureClass, ByVal pTargetWorkspace As IWorkspace2, ByVal sTargetName As String) As IEnumInvalidObject
        'Définir les variables de travail
        Dim pSourceDataset As IDataset = Nothing                    'Interface pour extraire le nom complet de la classe de départ.
        Dim pSourceFeatureClassName As IFeatureClassName = Nothing  'Interface pour ouvrir la classe de départ.
        Dim pSourceFields As IFields = Nothing                      'Interface contenant les attribut de la classe de départ.
        Dim pShapeField As IField = Nothing                         'Interface contenant l'attribut de la géométrie.
        Dim pTargetWorkspaceDataset As IDataset = Nothing           'Interface pour extraire le nom complet de la géodatabase d'arrivée.
        Dim pTargetWorkspaceName As IWorkspaceName = Nothing        'Interface pour ouvrir la géodatabase d'arrivée.
        Dim pTargetFeatureClassName As IFeatureClassName = Nothing  'Interface pour ouvrir la classe d'arrivée.
        Dim pTargetDatasetName As IDatasetName = Nothing            'Interface pour définir le Worspace d'arrivée.
        Dim pFieldChecker As IFieldChecker = Nothing                'Interface pour valider les attributs de la classe d'arrivée.
        Dim pTargetFields As IFields = Nothing                      'Interface contenant les attributs de la classe d'arrivée.
        Dim pEnumFieldError As IEnumFieldError = Nothing            'Interface contenant les attributs de la classe d'arrivée en erreur.
        Dim pGeometryDef As IGeometryDef = Nothing                  'Interface contenant la définition de l'attribut géométrie de la classe de départ.
        Dim pGeometryDefClone As IClone = Nothing                   'Interface utilisé pour cloner l'attribut géométrie de la classe de départ.
        Dim pTargetGeometryDefClone As IClone = Nothing             'Interface contenant l'attribut géométrie de la classe d'arrivée.
        Dim pTargetGeometryDef As IGeometryDef = Nothing            'Interface contenant la définition de la classe d'arrivée
        Dim pFeatureDataConverter As IFeatureDataConverter = Nothing 'Interface utilisé pour convertir une classe de départ vers une classe d'arrivée.
        Dim pFeatureWorkspace As IFeatureWorkspace = Nothing        'Interface pour ouvrir la FeatureClass existante
        Dim pDataset As IDataset = Nothing                          'Interface pour détruire la FeatureClass existante

        'Définir le résultat par défaut
        ConvertFeatureClass = Nothing

        Try
            'Interface pour extraire le nom complet de la classe de départ
            pSourceDataset = CType(pSourceClass, IDataset)
            'Interface pour ouvrir la classe de départ
            pSourceFeatureClassName = CType(pSourceDataset.FullName, IFeatureClassName)
            'Interface contenant les attribut de la classe de départ
            pSourceFields = pSourceClass.Fields
            'Interface contenant l'attribut de la géométrie
            pShapeField = pSourceFields.Field(pSourceFields.FindField(pSourceClass.ShapeFieldName))

            'Le nom de la FeatureClass ne doit pas dépasser 52
            If sTargetName.Length > 52 Then sTargetName = sTargetName.Substring(0, 52)
            'Interface pour extraire le nom complet de la géodatabase d'arrivée.
            pTargetWorkspaceDataset = CType(pTargetWorkspace, IDataset)
            'Interface pour ouvrir la géodatabase d'arrivée.
            pTargetWorkspaceName = CType(pTargetWorkspaceDataset.FullName, IWorkspaceName)
            'Interface pour ouvrir la classe d'arrivée.
            pTargetFeatureClassName = New FeatureClassNameClass()
            'Interface pour définir le Worspace d'arrivée.
            pTargetDatasetName = CType(pTargetFeatureClassName, IDatasetName)
            'Définir le nom de la Géodatabase d'arrivée
            pTargetDatasetName.Name = sTargetName
            'Définir le workspaceName d'arrivée
            pTargetDatasetName.WorkspaceName = pTargetWorkspaceName

            'Interface pour valider les attributs de la classe d'arrivée
            pFieldChecker = New FieldCheckerClass With _
                                                { _
                                                .InputWorkspace = pSourceDataset.Workspace, _
                                                .ValidateWorkspace = CType(pTargetWorkspace, IWorkspace) _
                                                }

            'Valider et définir les attributs de la classe d'arrivée
            pFieldChecker.Validate(pSourceClass.Fields, pEnumFieldError, pTargetFields)

            'Interface contenant la définition de la géométrie
            pGeometryDef = pShapeField.GeometryDef
            'Interface utilisé pour cloner l'attribut géométrie de la classe de départ
            pGeometryDefClone = CType(pGeometryDef, IClone)
            'Cloner l'attribut géométrie de la classe de départ pour la classe d'arrivée
            pTargetGeometryDefClone = pGeometryDefClone.Clone()
            'Interface contenant la définition de la classe d'arrivée
            pTargetGeometryDef = CType(pTargetGeometryDefClone, IGeometryDef)

            'Vérifier si la classe est déja présente
            If pTargetWorkspace.NameExists(esriDatasetType.esriDTFeatureClass, sTargetName) Then
                'Interface pour ouvrir la FeatureClass
                pFeatureWorkspace = CType(pTargetWorkspace, IFeatureWorkspace)
                'Ouvrir la FeatureClass
                pDataset = CType(pFeatureWorkspace.OpenFeatureClass(sTargetName), IDataset)
                'Détruire la FeatureClass
                pDataset.Delete()
            End If

            'Instancier un objet pour convertir une classe
            pFeatureDataConverter = New FeatureDataConverterClass()
            'Convertir la FeatureClass en mémoire sur disque
            ConvertFeatureClass = pFeatureDataConverter.ConvertFeatureClass(pSourceFeatureClassName, Nothing, Nothing, _
                                                                            pTargetFeatureClassName, pTargetGeometryDef, pTargetFields, _
                                                                            Nothing, 1000, 0)

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pSourceDataset = Nothing
            pSourceFeatureClassName = Nothing
            pSourceFields = Nothing
            pShapeField = Nothing
            pTargetWorkspaceDataset = Nothing
            pTargetWorkspaceName = Nothing
            pTargetFeatureClassName = Nothing
            pTargetDatasetName = Nothing
            pFieldChecker = Nothing
            pTargetFields = Nothing
            pEnumFieldError = Nothing
            pGeometryDef = Nothing
            pGeometryDefClone = Nothing
            pTargetGeometryDefClone = Nothing
            pTargetGeometryDef = Nothing
            pFeatureWorkspace = Nothing
            pDataset = Nothing
            pFeatureDataConverter = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'écrire le message d'exécution dans un RichTextBox, un fichier journal et/ou dans la console.
    '''</summary>
    ''' 
    '''<param name="sMessage"> Message à écrire dans un RichTextBox, un fichier journal et/ou dans la console.</param>
    '''
    Private Sub EcrireMessage(ByVal sMessage As String)
        Try
            'Conserver le message dans l'information du traitement
            gsInformation = gsInformation & sMessage & vbCrLf

            'Vérifier si le RichTextBox est présent
            If gqRichTextBox IsNot Nothing Then
                'Écrire le message dans le RichTextBox
                gqRichTextBox.AppendText(sMessage & vbCrLf)
            End If

            'Vérifier si le nom du fichier journal est présent
            If gsNomFichierJournal.Length > 0 Then
                'Écrire le message dans le RichTextBox
                File.AppendAllText(gsNomFichierJournal, sMessage & vbCrLf)
            End If

            'Écrire dans la console
            Console.WriteLine(sMessage)

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de conserver dans une collection le dernier FeatureLayer d'erreur ou de sélection traité.
    '''</summary>
    ''' 
    '''<param name="sRequete"> Requête correspondant à la commande pour conserver un FeatureLayer résultant.</param>
    '''<param name="qLayersColl"> Collection des FeatureLayer à conserver.</param>
    ''' 
    Private Sub ConserverLayer(ByVal sRequete As String, ByRef qLayersColl As Collection)
        'Déclaration des variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface pour sélectionner les éléments.
        Dim dDateDebut As DateTime = Nothing        'Contient la date de début du traitement.
        Dim sNomAncienLayer As String = ""          'Contient le nom de l'ancien Layer.
        Dim sNomNouveauLayer As String = ""         'Contient le nom du nouveau Layer.

        Try
            'Traiter une requête de la contrainte
            EcrireMessage(sRequete)
            'Afficher la date de début
            dDateDebut = System.DateTime.Now
            EcrireMessage("  Début: " & dDateDebut.ToString())

            'Définir le nom du Layer à conserver
            sNomAncienLayer = Split(sRequete, " ")(1)
            sNomAncienLayer = sNomAncienLayer.Replace(Chr(34), "")
            'Définir le nom de la copie du Layer
            sNomNouveauLayer = Split(sRequete, " ")(2)
            sNomNouveauLayer = sNomNouveauLayer.Replace(Chr(34), "")

            'Vérifier si on veut conserver le Layer d'erreur
            If sNomAncienLayer = "[ERREUR]" And gpFeatureLayerErreur IsNot Nothing Then
                'Modifier le nom du FeatureLayer
                gpFeatureLayerErreur.Name = sNomNouveauLayer

                'Ajouter le FeatureLayer dans la liste des Layers traités
                qLayersColl.Add(gpFeatureLayerErreur, sNomNouveauLayer)

                'Interface pour extraire les éléments sélectionnés
                pFeatureSel = CType(gpFeatureLayerErreur, IFeatureSelection)

                'Sélectionner tous les éléments
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)

                'Sinon on conserve le Layer de sélection
            Else
                'Modifier le nom du FeatureLayer
                gpFeatureLayerSelection.Name = sNomNouveauLayer

                'Ajouter le FeatureLayer dans la liste des Layers traités
                qLayersColl.Add(gpFeatureLayerSelection, sNomNouveauLayer)

                'Interface pour extraire les éléments sélectionnés
                pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            End If

            'Afficher les statistiques des éléments et des erreurs trouvées.
            EcrireMessage("  " & pFeatureSel.SelectionSet.Count.ToString & " éléments traités, " & pFeatureSel.SelectionSet.Count.ToString & " éléments sélectionnés.")
            'Afficher la date de fin
            EcrireMessage("  Fin: " & System.DateTime.Now.ToString() & " (Temps exécution: " & System.DateTime.Now.Subtract(dDateDebut).ToString & ")")

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de lire les éléments en mémoire dans un FeatureLayer et de le conserver dans une collection.
    '''</summary>
    ''' 
    '''<param name="sRequete"> Requête correspondant à la commande pour lire en mémoire un FeatureLayer résultant.</param>
    '''<param name="qLayersColl"> Collection des FeatureLayer à conserver.</param>
    ''' 
    Private Sub LireMemoire(ByVal sRequete As String, ByRef qLayersColl As Collection)
        'Déclaration des variables de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant les éléments en mémoire.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface pour sélectionner les éléments.
        Dim dDateDebut As DateTime = Nothing        'Contient la date de début du traitement.
        Dim sNomAncienLayer As String = ""          'Contient le nom de l'ancien Layer.
        Dim sNomNouveauLayer As String = ""         'Contient le nom du nouveau Layer.

        Try
            'Traiter une requête de la contrainte
            EcrireMessage(sRequete)
            'Afficher la date de début
            dDateDebut = System.DateTime.Now
            EcrireMessage("  Début: " & dDateDebut.ToString())

            'Définir le nom du Layer à conserver
            sNomAncienLayer = Split(sRequete, " ")(1)
            sNomAncienLayer = sNomAncienLayer.Replace(Chr(34), "")
            'Définir le nom de la copie du Layer
            sNomNouveauLayer = Split(sRequete, " ")(2)
            sNomNouveauLayer = sNomNouveauLayer.Replace(Chr(34), "")

            'Vérifier si le FeatureLayer de sélection est présent dans la collection des Layers conservés
            If qLayersColl.Contains(sNomAncienLayer) Then
                'Définir le FeatureLayer de sélection présent dans la liste des Layers traités
                pFeatureLayer = CType(qLayersColl.Item(sNomAncienLayer), IFeatureLayer)

                'Si le FeatureLayer est absent de la liste des Layers
            Else
                'Définir le FeatureLayer de sélection
                pFeatureLayer = CreerFeatureLayer(sNomAncienLayer, gpGeodatabase, gpSpatialReference, gpQueryFilter)

                'Ajouter le FeatureLayer dans la liste des Layers traités
                qLayersColl.Add(pFeatureLayer, sNomAncienLayer)
            End If

            'Créer un FeatureLayer contenant les éléments en mémoire
            gpFeatureLayerSelection = CreerLayerSelectionMemoire(pFeatureLayer, sNomNouveauLayer)

            'Ajouter le FeatureLayer dans la liste des Layers traités
            qLayersColl.Add(gpFeatureLayerSelection, sNomNouveauLayer)

            'Interface pour extraire les éléments sélectionnés
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Sélectionner tous les éléments
            pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)

            'Afficher les statistiques des éléments et des erreurs trouvées.
            EcrireMessage("  " & pFeatureSel.SelectionSet.Count.ToString & " éléments traités, " & pFeatureSel.SelectionSet.Count.ToString & " éléments sélectionnés.")
            'Afficher la date de fin
            EcrireMessage("  Fin: " & System.DateTime.Now.ToString() & " (Temps exécution: " & System.DateTime.Now.Subtract(dDateDebut).ToString & ")")

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
            pFeatureSel = Nothing
            dDateDebut = Nothing
            sNomAncienLayer = Nothing
            sNomNouveauLayer = Nothing
        End Try
    End Sub

    '''<summary>
    '''Fonction qui permet de créer un nouveau FeatureLayer contenant une FeatureClass en mémoire des éléments sélectionnés.
    '''</summary>
    '''
    '''<param name="pFeatureLayer">Interface ESRI contenant le FeatureLayer de sélection.</param>
    '''<param name="sNomLayer">Nom du nouveau Layer à Créer.</param>
    '''
    '''<returns>"IFeatureLayer" contenant seulement les éléments sélectionnés en mémoire.</returns>
    '''
    Protected Function CreerLayerSelectionMemoire(ByVal pFeatureLayer As IFeatureLayer, ByVal sNomLayer As String) As IFeatureLayer
        'Déclarer les variables de travail
        Dim pFeatureClass As IFeatureClass = Nothing    'Interface contenant la classe en mémoire.
        Dim pResult As IGeoProcessorResult = Nothing    'Interface contenant le résultat du Géotraitement.
        Dim pGPUtilities As IGPUtilities = Nothing      'Interface pour extraire le résultat obtenu du Geotraitement.
        Dim CopyFeatures As CopyFeatures = Nothing      'Objet contenant le Geotraitement à exécuter.
        Dim pGeoprocessor As ESRI.ArcGIS.Geoprocessor.Geoprocessor = Nothing  'Objet utilisé pour exécuter le Geotraitement.

        'Définir la valeur de retour par défaut
        CreerLayerSelectionMemoire = pFeatureLayer

        Try
            'Définir les objets du Géoprocessing
            pGPUtilities = New GPUtilities
            pGeoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor()
            CopyFeatures = New CopyFeatures()

            'Définir les paramètres du Géotraitement
            CopyFeatures.in_features = pFeatureLayer
            CopyFeatures.out_feature_class = "in_memory\" & sNomLayer
            pGeoprocessor.OverwriteOutput = True
            pGeoprocessor.AddOutputsToMap = False

            'Créer la FeatureClass en mémoire
            pResult = CType(pGeoprocessor.Execute(CopyFeatures, Nothing), IGeoProcessorResult)

            'Vérifier si le résultat est valide
            If pResult.Status = esriJobStatus.esriJobSucceeded Then
                'Définir la FeatureClass en mémoire
                pGPUtilities.DecodeFeatureLayer(pResult.GetOutput(0), pFeatureClass, Nothing)

                'Vérifier si la classe est valide
                If pFeatureClass IsNot Nothing Then
                    'Créer le nouveau FeatureLayer contenant la Featureclass en mémoire
                    CreerLayerSelectionMemoire = New FeatureLayer
                    CreerLayerSelectionMemoire.FeatureClass = pFeatureClass
                    CreerLayerSelectionMemoire.Name = sNomLayer
                    CreerLayerSelectionMemoire.Visible = False
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureClass = Nothing
            pResult = Nothing
            pGPUtilities = Nothing
            CopyFeatures = Nothing
            pGeoprocessor = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de créer un FeatureLayer contenant ou non une condition attributive.
    '''</summary>
    '''
    '''<param name="sNomClasse">Nom de la classe contenant ou non une condition attributive (ex:BDG_COURBE_NIVEAU_1:ELEVATION=100).</param>
    '''<param name="pFeatureWorkspace">Interface utilisé pour ouvrir une classe dans la Géodatabase.</param>
    '''<param name="pSpatialReference">Interface contenant la référence spatiale de traitement.</param>
    '''<param name="pQueryFilter">Interface utilisé pour sélectionner seulement les éléments du découpage.</param> 
    '''<param name="sNomProprietaireDefaut"> Contient le nom du propriétaire des tables par défaut pour les Géodatabase Enterprise.</param>
    ''' 
    '''<returns>IFeatureLayer contenant la classe à traiter avec sa requête attributive si spécifiée.</returns>
    ''' 
    Function CreerFeatureLayer(ByVal sNomClasse As String, ByVal pFeatureWorkspace As IFeatureWorkspace,
                               Optional ByVal pSpatialReference As ISpatialReference = Nothing,
                               Optional ByVal pQueryFilter As IQueryFilter = Nothing,
                               Optional ByVal sNomProprietaireDefaut As String = "BDG_DBA") As IFeatureLayer
        'Déclaration des variables de travail
        Dim pWorkspace2 As IWorkspace2 = Nothing                    'Interface pour vérifier si la table existe.
        Dim pWorkspace As IWorkspace = Nothing                      'Interface pour vérifier le type de Géodatabase.
        Dim sNomTable As String = ""        'Contient le nom de la table.
        Dim pFeatureClass As IFeatureClass = Nothing                'Interface contenant la classe de la Géodatabase.
        Dim pFeatureLayerDef As IFeatureLayerDefinition = Nothing   'Interface qui permet d'ajouter une condition attributive.
        Dim pFeatureSel As IFeatureSelection = Nothing              'Interface pour sélectionner les éléments.
        Dim sNomGeodatabase As String = ""  'Contient le nom de la Géodatabase.
        Dim sNomLayer As String = ""        'Contient le nom du Layer.
        Dim sFiltre As String = ""          'Contient le filtre sur le Layer. 
        Dim sLayer() As String = Nothing    'Contient la liste des paramètres d'un Layer (0:Nom de la classe, 1:Filtre).

        'Définir la valeur par défaut, une Layer vide
        CreerFeatureLayer = New FeatureLayer

        Try
            'Vérifier si le nom de la classe est présent
            If sNomClasse.Length > 0 Then
                'Extraire le nom du Layer et sa condition si elle est présente
                sLayer = Split(sNomClasse, ":")
                sNomLayer = sLayer(0)
                'Vérifier si le nom de la classe possède une condition attributive
                If sLayer.Length > 1 Then
                    'Définir le filtre
                    sFiltre = sLayer(1)
                End If

                'Interface pour vérifier le type de Géodatabase.
                pWorkspace = CType(pFeatureWorkspace, IWorkspace)
                'Vérifier si la Géodatabase est de type "Enterprise" 
                If pWorkspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                    'Vérifier si le nom de la table contient le nom du propriétaire
                    If Not sNomTable.Contains(".") Then
                        'Définir le nom de la table avec le nom du propriétaire
                        sNomTable = sNomProprietaireDefaut & "." & sNomTable
                    End If
                End If

                'Définir le nom de la table
                sNomTable = sNomLayer
                'Interface pour vérifier si la table existe
                pWorkspace2 = CType(pFeatureWorkspace, IWorkspace2)
                'Vérifier si la table existe
                If pWorkspace2.NameExists(esriDatasetType.esriDTFeatureClass, sNomTable) Then
                    'Ouvrir la FeatureClass contenue dans la Géodatabase
                    pFeatureClass = pFeatureWorkspace.OpenFeatureClass(sNomTable)
                    'Si la table n'existe pas
                Else
                    'Retourner une erreur
                    Throw New Exception("ERREUR : Le nom de la classe " & sNomTable & " est invalide ou absente de la Géodatabase!")
                End If

                'Si le nom de la classe est absent
            Else
                'Retourner l'erreur
                Throw New Exception("ERREUR : Le nom de la classe est absent!")
            End If

            'Définir le nom du FeatureLayer
            CreerFeatureLayer.Name = sNomLayer & "_" & System.DateTime.Now.ToString

            'Mettre non visible par défaut
            CreerFeatureLayer.Visible = False

            'Définir la FeatureClass liée au FeatureLayer
            CreerFeatureLayer.FeatureClass = pFeatureClass

            'Interface pour ajouter une condition attributive
            pFeatureLayerDef = CType(CreerFeatureLayer, IFeatureLayerDefinition)

            'Vérifier si le nom de la classe possède une condition attributive
            If sFiltre.Length > 0 Then
                'Vérifier si une requête additionnelle est nécessaire
                If pQueryFilter IsNot Nothing Then
                    'Ajouter la condition attributive et celle additionnelle
                    pFeatureLayerDef.DefinitionExpression = pQueryFilter.WhereClause & " AND (" & sFiltre & ")"

                    'Si on ajoute seulement la contifion attributive
                Else
                    'Ajouter la condition attributive seulement
                    pFeatureLayerDef.DefinitionExpression = sFiltre
                End If

                'Vérifier si une requête additionnelle est nécessaire
            ElseIf pQueryFilter IsNot Nothing Then
                'Ajouter la condition attributive additionnelle seulement
                pFeatureLayerDef.DefinitionExpression = pQueryFilter.WhereClause
            End If

            'Interface pour sélectionner les élément
            pFeatureSel = CType(CreerFeatureLayer, IFeatureSelection)

            'Sélectionner les éléments
            pFeatureSel.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)

            'Définir la référence spatiale du FeatureLayer au besoin
            If pSpatialReference IsNot Nothing Then CreerFeatureLayer.SpatialReference = pSpatialReference

        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pWorkspace2 = Nothing
            pWorkspace = Nothing
            pFeatureClass = Nothing
            pFeatureLayerDef = Nothing
            pFeatureSel = Nothing
            sNomGeodatabase = Nothing
            sNomLayer = Nothing
            sNomTable = Nothing
            sFiltre = Nothing
            sLayer = Nothing
        End Try
    End Function
#End Region
End Class
