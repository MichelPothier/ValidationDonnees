Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms

'**
'Nom de la composante : clsTableContraintes.vb
'
'''<summary>
''' Classe générique qui permet de gérer une table de contraintes d'intégrité à appliquer.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 02 mai 2016
'''</remarks>
''' 
Public Class clsTableContraintes
    '''<summary>Interface contenant le nom de la géodatabase contenant les classes à traiter.</summary>
    Protected gsNomGeodatabaseClasses As String = ""
    '''<summary>Contient le nom de la table des contraintes d'intégrité à traiter.</summary>
    Protected gsNomTableContraintes As String = ""
    '''<summary>Contient la valeur du filtre appliqué sur la table des contraintes d'intégrité à traiter.</summary>
    Protected gsFiltreContraintes As String = ""
    '''<summary>Interface contenant le nom de la classe de découpage à traiter.</summary>
    Protected gsNomClasseDecoupage As String = ""
    '''<summary>Interface contenant le nom de l'attribut de découpage à traiter.</summary>
    Protected gsNomAttributDecoupage As String = ""
    '''<summary>Nom du fichier contenant le rapport d'erreurs.</summary>
    Protected gsNomRapportErreurs As String = ""
    '''<summary>Nom du répertoire contenant les erreurs.</summary>
    Protected gsNomRepertoireErreurs As String = ""
    '''<summary>Nom du fichier journal d'éxécution du traitement.</summary>
    Protected gsNomFichierJournal As String = ""
    '''<summary>Contient l'information de traitement.</summary>
    Protected gsInformation As String = ""
    '''<summary>Courriel dans lesquels sera retourné le rapport d'erreurs.</summary>
    Protected gsCourriel As String = ""
    '''<summary>Contient le nom de la table des statistiques d'erreurs et de traitement.</summary>
    Protected gsNomTableStatistiques As String = ""
    '''<summary>Commande utilisée pour exécuter le traitement de validation des contraintes d'intégrité.</summary>
    Protected gsCommande As String = ""

    '''<summary>Contient le code d'erreur d'exécution des contraintes.</summary>
    Protected giCodeErreur As Integer = 0
    '''<summary>Contient le nombre total de découpage traitées.</summary>
    Protected giNombreTotalDecoupages As Integer = 0
    '''<summary>Contient le nombre total de contraintes traitées.</summary>
    Protected giNombreTotalContraintes As Integer = 0
    '''<summary>Contient le nombre total d'érreurs traitées.</summary>
    Protected giNombreTotalErreurs As Long = 0
    '''<summary>Contient le nombre total d'éléments traités.</summary>
    Protected giNombreTotalElements As Long = 0

    '''<summary>Indique si les erreurs seront ajoutées à la Map active.</summary>
    Protected gbAjouterErreursMap As Boolean = False
    '''<summary>Interface contenant la Map active lorsqu'appelé via ArcMap.</summary>
    Protected gpMap As IMap = Nothing
    '''<summary>Interface utilisé pour afficher la barre de progression et pour annuler le traitement.</summary>
    Protected gpTrackCancel As ITrackCancel = Nothing
    '''<summary>RichTextBox utilisé pour écrire le journal d'éxécution du traitement.</summary>
    Protected gqRichTextBox As RichTextBox = Nothing
    '''<summary>Interface contenant une référence spatiale.</summary>
    Protected gpSpatialReference As ISpatialReference = Nothing
    '''<summary>Interface contenant la géodatabase des classes à traiter.</summary>
    Protected gpGeodatabaseClasses As IFeatureWorkspace = Nothing
    '''<summary>Interface contenant la géodatabase de la table des contraintes d'intégrité à traiter.</summary>
    Protected gpGeodatabaseContraintes As IFeatureWorkspace = Nothing
    '''<summary>Interface contenant la table des contraintes d'intégrité à appliquer.</summary>
    Protected gpTableContraintes As IStandaloneTable = Nothing
    '''<summary>Interface contenant le FeatureLayer des éléments de découpage à traiter.</summary>
    Protected gpFeatureLayerDecoupage As IFeatureLayer = Nothing
    '''<summary>Interface contenant la table des statistiques d'erreurs et de traitement.</summary>
    Protected gpTableStatistiques As ITable = Nothing
    '''<summary>Interface contenant la géodatabase de la table des statistiques à traiter.</summary>
    Protected gpGeodatabaseStatistiques As IFeatureWorkspace = Nothing

#Region "Constructeurs"
    '''<summary>
    ''' Routine qui permet d'instancier un objet qui permet de gérer une table des contraintes d'intégrité à appliquer à partir d'un nom de table.
    '''</summary>
    ''' 
    '''<param name="sNomGeodatabaseClasses"> Contient le nom de la Géodatabase des classes à traiter.</param>
    '''<param name="sNomTableContraintes"> Contient le nom de la table des contraintes d'intégrité à traiter.</param>
    '''<param name="sFiltreContraintes"> Contient le nom de la table des contraintes d'intégrité à traiter.</param>
    ''' 
    Public Sub New(ByVal sNomGeodatabaseClasses As String, ByVal sNomTableContraintes As String, Optional ByVal sFiltreContraintes As String = "")
        'Déclarer les variables de travail

        Try
            'Conserver le nom de la Géodatabases des classes à traiter
            gsNomGeodatabaseClasses = sNomGeodatabaseClasses

            'Conserver le nom de la table des contraintes d'intégrité
            gsNomTableContraintes = sNomTableContraintes

            'Conserver le filtre sur les contraintes
            gsFiltreContraintes = sFiltreContraintes

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier un objet qui permet de gérer une table des contraintes d'intégrité à appliquer.
    '''</summary>
    ''' 
    '''<param name="pTableContraintes"> Interface contenant la table des contraintes d'intégrité à appliquer.</param>
    '''
    Public Sub New(ByVal pGeodatabaseClasses As IFeatureWorkspace, ByVal pTableContraintes As IStandaloneTable)
        'Déclarer les variables de travail
        Dim pWorkspace As IWorkspace = Nothing      'Interface utilisé pour extraire le nom de laa Géodatabase.
        Dim pDatasetName As IDatasetName = Nothing  'Interface utilisé pour extraire le nom complet de la table.
        Dim pTableDef As ITableDefinition = Nothing 'Interface qui permet d'appliquer un filtre sur les contraintes.
        Dim pTableSel As ITableSelection = Nothing  'Interface qui permet d'extraire le nombre de contraintes sélectionnées.

        Try
            'Conserver la Géodatabase des classes à traiter
            gpGeodatabaseClasses = pGeodatabaseClasses
            'Interface pour extraire le nom de la Géodatabase
            pWorkspace = CType(gpGeodatabaseClasses, IWorkspace)
            'Définir le nom de la Géodatabase des classes
            gsNomGeodatabaseClasses = pWorkspace.PathName

            'Conserver la table des contraintes d'intégrité à traiter.
            gpTableContraintes = pTableContraintes
            'Interface pour extraire le nom complet de la table des contraintes d'intégrité.
            pDatasetName = CType(pTableContraintes, IDatasetName)
            'Définir le nom complet de la table des contraintes d'intégrité.
            gsNomTableContraintes = pDatasetName.WorkspaceName.BrowseName & "\" & pDatasetName.Name

            'Interface pour appliquer un filtre
            pTableDef = CType(gpTableContraintes, ITableDefinition)
            'Définir le filtre sur les contraintes
            gsFiltreContraintes = pTableDef.DefinitionExpression

            'Interface pour extraire le nombre de contraintes sélectionnées
            pTableSel = CType(gpTableContraintes, ITableSelection)
            'Vérifier si aucune contrainte sélectionnées
            If pTableSel.SelectionSet.Count = 0 Then
                'Sélectionner toutes les contraintes
                pTableSel.SelectRows(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If
            'Conserver le nombre total de contraintes à traiter
            giNombreTotalContraintes = pTableSel.SelectionSet.Count

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pWorkspace = Nothing
            pDatasetName = Nothing
            pTableDef = Nothing
            pTableSel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier un objet qui permet de gérer une table des contraintes d'intégrité à appliquer 
    ''' à partir d'une noeud de TreeNode.
    '''</summary>
    ''' 
    '''<param name="sNomGeodatabaseClasses"> Nom de la Géodatabase des classes à traiter.</param>
    '''<param name="oTreeNode"> TreeNode de type [CONTRAINTE].</param>
    '''<param name="pMap"> Map active dans laquelle la table est présente.</param>
    '''<param name="bAjouterErreursMap"> Indique si les erreurs seront ajoutées dans la Map active.</param>
    '''
    Public Sub New(ByVal sNomGeodatabaseClasses As String, ByVal oTreeNode As TreeNode, ByVal pMap As IMap, ByVal bAjouterErreursMap As Boolean)
        'Déclarer les variables de travail

        Try
            'Conserver le nom de la Géodatabases des classes à traiter
            gsNomGeodatabaseClasses = sNomGeodatabaseClasses

            'Conserver la Map active
            gpMap = pMap

            'Conserver si on doit ajouter les erreurs dans la map active
            gbAjouterErreursMap = bAjouterErreursMap

            'Sélectionner les contraintes de la table à partir du Noeud du TreeNode spécifié et une Map
            SelectionnerContraintes(oTreeNode, pMap)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de vider la mémoire.
    '''</summary>
    Protected Overrides Sub Finalize()
        'Vider la mémoire
        gsNomGeodatabaseClasses = Nothing
        gsNomTableContraintes = Nothing
        gsFiltreContraintes = Nothing
        gsNomClasseDecoupage = Nothing
        gsNomAttributDecoupage = Nothing
        gsNomRapportErreurs = Nothing
        gsNomRepertoireErreurs = Nothing
        gsNomFichierJournal = Nothing
        gsInformation = Nothing
        gsCourriel = Nothing
        gsNomTableStatistiques = Nothing

        gpTrackCancel = Nothing
        gpSpatialReference = Nothing
        gpGeodatabaseClasses = Nothing
        gpGeodatabaseContraintes = Nothing
        gpTableContraintes = Nothing
        gpFeatureLayerDecoupage = Nothing
        gpTableStatistiques = Nothing
        gpGeodatabaseStatistiques = Nothing
        gqRichTextBox = Nothing

        giNombreTotalDecoupages = Nothing
        giNombreTotalContraintes = Nothing
        giNombreTotalErreurs = Nothing
        giNombreTotalElements = Nothing

        'Récupération de la mémoire disponible
        GC.Collect()
        'Finaliser
        MyBase.Finalize()
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de définir et retourner la commande utilisée pour exécuter le traitement de validation des contraintes d'intégrité.
    '''</summary>
    ''' 
    Public Property Commande() As String
        Get
            'Vérifier si la commande a été spécifiée
            If gsCommande.Length > 0 Then
                'Retourner la commande spécifiée
                Commande = gsCommande

                'Si la commande n'a pas été spécifiée
            Else
                'Retourner la commande construite
                Commande = "ValiderContrainte.exe """ & gsNomGeodatabaseClasses & """ """ & gsNomTableContraintes _
                           & """ """ & gsFiltreContraintes & """ """ & gsNomClasseDecoupage & """ """ & gsNomAttributDecoupage _
                           & """ """ & gsNomRepertoireErreurs & """ """ & gsNomRapportErreurs & """ """ & gsNomFichierJournal _
                           & """ """ & gsCourriel & """"
            End If
        End Get
        Set(ByVal value As String)
            'Conserver la commande exécutée
            gsCommande = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nom de la table des contraintes d'intégrité à appliquer.
    '''</summary>
    ''' 
    Public ReadOnly Property NomTableContraintes() As String
        Get
            NomTableContraintes = gsNomTableContraintes
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner la table des contraintes d'intégrité à appliquer.
    '''</summary>
    ''' 
    Public ReadOnly Property TableContraintes() As IStandaloneTable
        Get
            TableContraintes = gpTableContraintes
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet et retourner la Géodatabase contenant la table des contraintes d'intégrité.
    '''</summary>
    ''' 
    Public ReadOnly Property GeodatabaseContraintes() As IFeatureWorkspace
        Get
            GeodatabaseContraintes = gpGeodatabaseContraintes
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom du répertoire d'erreurs.
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
    ''' Propriété qui permet de définir et retourner le nom du fichier contenant le rapport d'erreurs.
    '''</summary>
    ''' 
    Public Property NomRapportErreurs() As String
        Get
            NomRapportErreurs = gsNomRapportErreurs
        End Get
        Set(ByVal value As String)
            'Conserver le nom du Rapport d'erreurs
            gsNomRapportErreurs = value

            'Vérifier si le rapport d'erreurs est présent
            If File.Exists(gsNomRapportErreurs) Then
                'Détruire le fichier
                File.Delete(gsNomRapportErreurs)
            End If
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
            'Conserver le nom du fichier journal
            gsNomFichierJournal = value

            'Vérifier si le fichier journal est présent
            If File.Exists(gsNomFichierJournal) Then
                'Détruire le fichier
                File.Delete(gsNomFichierJournal)
            End If
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
    ''' Propriété qui permet de définir et retourner le nom de la géodatabase contenant les classes à traiter.
    '''</summary>
    ''' 
    Public Property NomGeodatabaseClasses() As String
        Get
            NomGeodatabaseClasses = gsNomGeodatabaseClasses
        End Get
        Set(ByVal value As String)
            gsNomGeodatabaseClasses = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la Géodatabase contenant les classes des éléments à valider.
    '''</summary>
    ''' 
    Public Property GeodatabaseClasses() As IFeatureWorkspace
        Get
            GeodatabaseClasses = gpGeodatabaseClasses
        End Get
        Set(ByVal value As IFeatureWorkspace)
            gpGeodatabaseClasses = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le filtre à appliquer sur la table des contraintes d'intégrité.
    '''</summary>
    ''' 
    Public Property FiltreContraintes() As String
        Get
            FiltreContraintes = gsFiltreContraintes
        End Get
        Set(ByVal value As String)
            gsFiltreContraintes = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le FeatureLayer contenant les éléments de découpage à traiter.
    '''</summary>
    ''' 
    Public Property FeatureLayerDecoupage() As IFeatureLayer
        Get
            FeatureLayerDecoupage = gpFeatureLayerDecoupage
        End Get
        Set(ByVal value As IFeatureLayer)
            gpFeatureLayerDecoupage = value
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
    ''' Propriété qui permet de définir et retourner le nom de l'attribut de découpage à traiter.
    '''</summary>
    ''' 
    Public Property NomAttributDecoupage() As String
        Get
            NomAttributDecoupage = gsNomAttributDecoupage
        End Get
        Set(ByVal value As String)
            gsNomAttributDecoupage = value
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

    '''<summary>
    ''' Propriété qui permet de définir et retourner les courriels dans lesquels le rapport d'erreurs pourra être envoyé.
    '''</summary>
    ''' 
    Public Property Courriel() As String
        Get
            Courriel = gsCourriel
        End Get
        Set(ByVal value As String)
            gsCourriel = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom de la table des statistiques d'erreurs et de traitement.
    '''</summary>
    ''' 
    Public Property NomTableStatistiques() As String
        Get
            NomTableStatistiques = gsNomTableStatistiques
        End Get
        Set(ByVal value As String)
            gsNomTableStatistiques = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le code d'erreur d'exécution des contrainte.
    '''  0: Aucune erreur d'exécution.
    '''  1: Erreur d'exécution survenue dans une requête mais sans arrêt du programme.
    ''' -1: Erreur d'exécution survenue avec arrêt du programme.
    '''</summary>
    ''' 
    Public ReadOnly Property CodeErreur() As Integer
        Get
            'Retourner lecode d'erreur d'exécution.
            CodeErreur = giCodeErreur
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre total de découpage traités.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreTotalDecoupage() As Integer
        Get
            'Retourner le nombre total de découpage traités.
            NombreTotalDecoupage = giNombreTotalDecoupages
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre total de contraintes traitées.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreTotalContraintes() As Integer
        Get
            'Retourner le nombre total de contraintes traitées.
            NombreTotalContraintes = giNombreTotalContraintes
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre total d'erreurs trouvées.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreTotalErreurs() As Long
        Get
            'Retourner le nombre total d'erreurs trouvées
            NombreTotalErreurs = giNombreTotalErreurs
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre total d'éléments traités.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreTotalElements() As Long
        Get
            'Retourner le nombre total d'éléments traités
            NombreTotalElements = giNombreTotalElements
        End Get
    End Property
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    ''' Routine qui permet de vérifier si on peut exécuter les contraintes d'intégrité et retourner le résultat obtenu.
    '''</summary>
    ''' 
    '''<returns>True si la contrainte est valide, False sinon.</returns>
    ''' 
    Public Function EstValide() As Boolean
        'Déclaration des variables de travail
        Dim pTableSel As ITableSelection = Nothing      'Interface contenant les contraintes sélectionnées.

        'Par défaut, la contrainte est valide
        EstValide = True
        gsInformation = "Le traitement des contraintes d'intégrité est valide !"

        Try

            'Vérifier si la géodatabase des classes à traiter est invalide
            If gpGeodatabaseClasses Is Nothing Then
                'Définir le message d'erreur
                EstValide = False
                giCodeErreur = 2
                gsInformation = "ERREUR : La géodatabase est invalide : " & gsNomGeodatabaseClasses

                'Vérifier si le standaloneTable est invalide
            ElseIf gpTableContraintes Is Nothing Then
                'Définir le message d'erreur
                EstValide = False
                giCodeErreur = 2
                gsInformation = "ERREUR : La table des contraintes d'intégrité est invalide : " & gsNomTableContraintes

                'Vérifier si la table est invalide
            ElseIf gpTableContraintes.Table Is Nothing Then
                'Définir le message d'erreur
                EstValide = False
                giCodeErreur = 2
                gsInformation = "ERREUR : La table des contraintes d'intégrité est invalide : " & gpTableContraintes.Name

                'Si la table est valide
            Else
                'Interface pour vérifier la sélection des contraintes
                pTableSel = CType(gpTableContraintes, ITableSelection)

                'Vérifier si aucune sélection
                If pTableSel.SelectionSet.Count = 0 Then
                    'Sélectionner toutes les contraintes
                    pTableSel.SelectRows(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                End If

                'Vérifier si aucune contrainte à traiter
                If pTableSel.SelectionSet.Count = 0 Then
                    'Définir le message d'erreur
                    EstValide = False
                    giCodeErreur = 2
                    gsInformation = "ERREUR : Aucune contrainte d'intégrité à traiter : " & gsNomTableContraintes
                End If

                'Vérifier si la table des contraintes contient l'attribut "REQUETES"
                If gpTableContraintes.Table.FindField("REQUETES") = -1 Then
                    'Définir le message d'erreur
                    EstValide = False
                    giCodeErreur = 2
                    gsInformation = "ERREUR : La table des contraintes ne contient pas l'attribut 'REQUETES' : " & gsNomTableContraintes
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pTableSel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'exécuter la validation des contraintes d'intégrité sélectionnées de la table 
    ''' en fonction des découpages à traiter et retourner le résultat obtenu.
    '''</summary>
    '''
    '''<returns>Boolean qui indique si le traitement a été exécuté avec succès.</returns>
    ''' 
    Public Function Executer() As Boolean
        'Déclaration des variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface qui permet d'extraire les découpage sélectionnées.
        Dim pCursor As ICursor = Nothing                'Interface utilisé pour extraire les découpages.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface utilisé pour extraire les découpages.
        Dim pFeatureDecoupage As IFeature = Nothing     'Interface contenant un élément de découpage.
        Dim pQueryFilter As IQueryFilter = Nothing      'Interface qui permet de traiter seulement les éléments du découpage.
        Dim pQueryFilterDef As IQueryFilterDefinition = Nothing 'Interface utilisé pour indiquer que l'on veut trier.
        Dim oContrainte As clsContrainte = Nothing      'Objet utilisé pour traiter une contrainte d'intégrité.
        Dim dDateDebut As DateTime = Nothing            'Contient la date de début du traitement.
        Dim oProcess As Process = Nothing               'Contient un processus pour envoyer un courriel.
        Dim sCmd As String = Nothing            'Contient la commande pour envoyer un courriel.
        Dim iPosAtt As Integer = 0              'Position de l'attribut de découpage.
        Dim iNbDecoupage As Integer = 0         'Compteur de découpage.
        Dim iNbTotalDecoupage As Integer = 0    'Nombre total de découpage.

        'Par défaut, le traitement ne s'est pas exécuté avec succès
        Executer = False

        Try
            'Initialiser le résultat
            giNombreTotalElements = 0
            giNombreTotalErreurs = 0

            'Définir la date de début
            dDateDebut = System.DateTime.Now
            'Vérifier si le nom du répertoire d'erreurs est présent
            If gsNomRepertoireErreurs.Length > 0 Then
                'Redéfinir les noms de fichier contenant le mot [DATE_TIME]
                gsNomRepertoireErreurs = gsNomRepertoireErreurs.Replace("[DATE_TIME]", dDateDebut.ToString("yyyyMMdd_HHmmss"))
                'Vérifier si les répertoires existent
                If Not IO.Directory.Exists(IO.Path.GetDirectoryName(gsNomRepertoireErreurs)) Then
                    'Créer le répertoire
                    IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(gsNomRepertoireErreurs))
                End If
            End If
            'Vérifier si le nom du rapport d'erreurs est présent
            If gsNomRapportErreurs.Length > 0 Then
                'Redéfinir les noms de fichier contenant le mot [DATE_TIME]
                gsNomRapportErreurs = gsNomRapportErreurs.Replace("[DATE_TIME]", dDateDebut.ToString("yyyyMMdd_HHmmss"))
                'Vérifier si les répertoires existent
                If Not IO.Directory.Exists(IO.Path.GetDirectoryName(gsNomRapportErreurs)) Then
                    'Créer le répertoire
                    IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(gsNomRapportErreurs))
                End If

            End If
            'Vérifier si le nom du fichier journal est présent
            If gsNomFichierJournal.Length > 0 Then
                'Redéfinir les noms de fichier contenant le mot [DATE_TIME]
                gsNomFichierJournal = gsNomFichierJournal.Replace("[DATE_TIME]", dDateDebut.ToString("yyyyMMdd_HHmmss"))
                'Vérifier si les répertoires existent
                If Not IO.Directory.Exists(IO.Path.GetDirectoryName(gsNomFichierJournal)) Then
                    'Créer le répertoire
                    IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(gsNomFichierJournal))
                End If
            End If

            'Afficher les paramètres d'exécution du traitement
            EcrireMessage("")
            EcrireMessage(Me.Commande)
            EcrireMessage("--------------------------------------------------------------------------------")
            EcrireMessage("-Version : " & IO.File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString)
            EcrireMessage("-Usager  : " & System.Environment.GetEnvironmentVariable("USERNAME"))
            EcrireMessage("-Date    : " & dDateDebut.ToString)
            EcrireMessage("")
            EcrireMessage("-Paramètres :")
            EcrireMessage(" ------------------------")
            EcrireMessage(" Géodatabase des classes : " & gsNomGeodatabaseClasses)
            EcrireMessage(" Table des contraintes   : " & gsNomTableContraintes)
            EcrireMessage(" Filtre des contraintes  : " & gsFiltreContraintes)
            EcrireMessage(" Classe de découpage     : " & gsNomClasseDecoupage)
            EcrireMessage(" Attribut de découpage   : " & gsNomAttributDecoupage)
            EcrireMessage(" Répertoire d'erreurs    : " & gsNomRepertoireErreurs)
            EcrireMessage(" Rapport d'erreurs       : " & gsNomRapportErreurs)
            EcrireMessage(" Journal d'exécution     : " & gsNomFichierJournal)
            EcrireMessage(" Courriel                : " & gsCourriel)
            EcrireMessage(" Table des statistiques  : " & gsNomTableStatistiques)
            EcrireMessage("--------------------------------------------------------------------------------")
            EcrireMessage("")

            'Écrire les statistiques d'utilisation
            Call EcrireStatistiqueUtilisation("clsTableContraintes.Executer " & Me.NomGeodatabaseClasses)

            'Définir la Géodatabase des classes à traiter
            gpGeodatabaseClasses = CType(DefinirGeodatabase(gsNomGeodatabaseClasses), IFeatureWorkspace)

            'Définir le Standalone de la table des contraintes d'intégrité
            DefinirTableContraintes(gsNomTableContraintes, gsFiltreContraintes)

            'Définir la classe de découpage
            DefinirClasseDecoupage(gsNomClasseDecoupage)

            'Définir la table des statistiques d'erreurs et de traitement.
            DefinirTableStatistiques(gsNomTableStatistiques)

            'Vérifier si la contrainte est valide
            If EstValide() Then
                'Vérifier si aucun découpage à traiter
                If Me.NombreTotalDecoupage = 0 Then
                    'Valides les contraintes sélectionnés de la table sans découpage
                    ValiderContraintes()

                    'Si plusieurs découpages à traiter
                Else
                    'Définir la position de l'attribut de découpage
                    iPosAtt = gpFeatureLayerDecoupage.FeatureClass.FindField(gsNomAttributDecoupage)

                    'Interface pour traiter tous les découpages sélectionnés
                    pFeatureSel = CType(gpFeatureLayerDecoupage, IFeatureSelection)

                    'Conserver le nombre total de découpage  à traiter
                    iNbTotalDecoupage = pFeatureSel.SelectionSet.Count

                    'Créer une nouvelle requete vide
                    pQueryFilterDef = New QueryFilter
                    'Indiquer la méthode pour trier
                    pQueryFilterDef.PostfixClause = "ORDER BY " & gsNomAttributDecoupage

                    'Extraire les découpages à traiter
                    pFeatureSel.SelectionSet.Search(CType(pQueryFilterDef, IQueryFilter), False, pCursor)
                    pFeatureCursor = CType(pCursor, IFeatureCursor)

                    'Extraire le premier découpage sélectionné
                    pFeatureDecoupage = pFeatureCursor.NextFeature()

                    'Traiter tous les découpages sélectionnés
                    Do Until pFeatureDecoupage Is Nothing
                        'Compter les découpages
                        iNbDecoupage = iNbDecoupage + 1

                        'Interface pour définir la requête pour extraire les éléments du découpage
                        pQueryFilter = New QueryFilter
                        'Définir la requête pour extraire les éléments du découpage
                        pQueryFilter.WhereClause = gsNomAttributDecoupage & "='" & pFeatureDecoupage.Value(iPosAtt).ToString & "'"

                        'Valides les contraintes sélectionnés de la table en fonction d'un découpage
                        ValiderContraintes(pFeatureDecoupage, pQueryFilter, _
                                           iNbDecoupage.ToString & "/" & iNbTotalDecoupage.ToString, _
                                           pFeatureDecoupage.Value(iPosAtt).ToString)

                        'Extraire le prochain découpage sélectionné
                        pFeatureDecoupage = pFeatureCursor.NextFeature()
                    Loop
                End If

                'Afficher le résultat de l'exécution de la requête
                EcrireMessage("")
                EcrireMessage("-Statistiques sur le traitement exécuté:")
                EcrireMessage(" ---------------------------------------")
                EcrireMessage("  Nombre total de découpage traités    : " & giNombreTotalDecoupages.ToString)
                EcrireMessage("  Nombre total de contraintes traitées : " & giNombreTotalContraintes.ToString)
                EcrireMessage("  Nombre total d'éléments traités      : " & giNombreTotalElements.ToString)
                EcrireMessage("  Nombre total d'erreurs trouvées      : " & giNombreTotalErreurs.ToString)
                EcrireMessage("")
                EcrireMessage("-Temps total d'exécution : " & System.DateTime.Now.Subtract(dDateDebut).ToString)

                'Indique le succès du traitement
                Executer = True

                'Si l'exécution est invalide
            Else
                'Écrire le message d'erreur
                EcrireMessage(gsInformation)
            End If

            'Vérifier si on doit envoyer un courriel
            If gsNomRapportErreurs.Length > 0 And gsCourriel.Length > 0 Then
                'Définir le nom de la commande pour envoyer le courriel
                sCmd = "\\DFSCITSH\CITS\EnvCits\applications\cits\pro\py\EnvoyerCourriel.py -to " & gsCourriel & " -subj ""Rapport d'intégrité des contraintes spatiales: " & IO.Path.GetFileName(gsNomRapportErreurs) & """ -f " & gsNomRapportErreurs & " -encodage utf-8"
                'Afficher l'envoit du message
                EcrireMessage(" ")
                EcrireMessage(Environment.GetEnvironmentVariable("python27") & " " & sCmd)
                'Envoyer le courriel
                oProcess = Process.Start(Environment.GetEnvironmentVariable("python27"), sCmd)
            End If

        Catch ex As Exception
            'Écrire l'erreur
            EcrireMessage("[clsTableContraintes.Executer] " & ex.StackTrace)
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeatureDecoupage = Nothing
            pQueryFilter = Nothing
            oContrainte = Nothing
            dDateDebut = Nothing
            oProcess = Nothing
            sCmd = Nothing
            iPosAtt = Nothing
            iNbDecoupage = Nothing
            iNbTotalDecoupage = Nothing
            pQueryFilterDef = Nothing
            GC.Collect()
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet d'exécuter la validation des contraintes d'intégrité sélectionnées en fonction d'un découpage ou non.
    '''</summary>
    '''
    '''<param name="pFeatureDecoupage"> Élément de découpage à traiter.</param>
    '''<param name="pQueryFilter"> Filtre permettant d'extraire les éléments correspondant au découpage à traiter.</param>
    '''<param name="sIdentifiant"> Identifiant du traitement de découpage.</param>
    '''<param name="sNomDecoupage"> Nom de l'élément de découpage à traiter.</param>
    ''' 
    Private Sub ValiderContraintes(Optional ByVal pFeatureDecoupage As IFeature = Nothing, _
                                   Optional ByVal pQueryFilter As IQueryFilter = Nothing, _
                                   Optional ByVal sIdentifiant As String = "",
                                   Optional ByVal sNomDecoupage As String = "")
        'Déclaration des variables de travail
        Dim pTableSel As ITableSelection = Nothing      'Interface qui permet d'extraire les contraintes sélectionnées.
        Dim pRow As IRow = Nothing                      'Interface contenant une contrainte d'intégrité.
        Dim oContrainte As clsContrainte = Nothing      'Objet utilisé pour traiter une contrainte d'intégrité.
        Dim pTableSort As ITableSort = Nothing          'Interface qui permet de trier les contraintes.
        Dim pEnumIDs As IEnumIDs = Nothing              'Interface qui permet d'extraire les OIDs.
        Dim sNomRepertoireErreurs As String = ""        'Contient le nom du répertoire ou de la Géodatabase d'erreurs spécifique.
        Dim iOid As Integer = Nothing           'Contient le OID de la contrainte à traiter.
        Dim iNbContrainte As Integer = 0        'Compteur de contraintes.
        Dim iNbTotalContrainte As Integer = 0   'Nombre total de contraintes.

        Try
            'Redéfinir le nom du fichier contenant le mot [DATASET_NAME] de l'attribut de découpage
            sNomRepertoireErreurs = gsNomRepertoireErreurs.Replace("[" & gsNomAttributDecoupage & "]", sNomDecoupage)

            'Interface pour traiter toutes les contraintes d'intégrités sélectionnées
            pTableSel = CType(gpTableContraintes, ITableSelection)

            'Conserver le nombre total de contraintes
            iNbTotalContrainte = pTableSel.SelectionSet.Count

            'Conserver le nombre total de contraintes à traiter
            giNombreTotalContraintes = pTableSel.SelectionSet.Count

            'Interface pour trier de façon ascendante les contraintes en fonction du nom de la table et du groupe
            pTableSort = New TableSortClass()
            pTableSort.Fields = "NOM_TABLE,GROUPE"
            pTableSort.SelectionSet = pTableSel.SelectionSet
            pTableSort.Ascending("NOM_TABLE,GROUPE") = True
            pTableSort.Sort(Nothing)

            'Interface pour extraire les OIds triés
            pEnumIDs = pTableSort.IDs

            'Extraire le premier Oid des contraintes sélectionnées
            pEnumIDs.Reset()
            iOid = pEnumIDs.Next()

            'Traiter toutes les contraintes d'intégrités sélectionnées
            Do Until iOid = -1
                'Définir la contrainte à traiter
                pRow = gpTableContraintes.Table.GetRow(iOid)

                'Compteur de contraintes utilisé pour définir l'identifiant
                iNbContrainte = iNbContrainte + 1

                'Définir la contrainte d'intégrité
                oContrainte = New clsContrainte(pRow)
                oContrainte.TrackCancel = gpTrackCancel
                oContrainte.RichTextbox = gqRichTextBox
                oContrainte.NomFichierJournal = gsNomFichierJournal
                oContrainte.NomRepertoireErreurs = sNomRepertoireErreurs
                oContrainte.Geodatabase = gpGeodatabaseClasses
                oContrainte.SpatialReference = gpSpatialReference
                oContrainte.TableStatistiques = gpTableStatistiques
                'Vérifier si l'élément de découpage est absent
                If pFeatureDecoupage Is Nothing Then
                    'Définir le nom de la classe de découpage à traiter
                    oContrainte.NomClasseDecoupage = gsNomClasseDecoupage
                    'Définir l'entête du traitement sans découpage
                    oContrainte.Entete = iNbContrainte.ToString & "/" & iNbTotalContrainte.ToString
                    'Définir l'identifiant du traitement sans découpage
                    oContrainte.Identifiant = "CANADA"
                Else
                    'Définir l'élément de découpage à traiter
                    oContrainte.FeatureLayerDecoupage = gpFeatureLayerDecoupage
                    'Définir l'élément de découpage à traiter
                    oContrainte.FeatureDecoupage = pFeatureDecoupage
                    'Définir le filtre correspondant à l'élément de découpage
                    oContrainte.QueryFilter = pQueryFilter
                    'Définir l'entête du traitement avec un découpage
                    oContrainte.Entete = sIdentifiant & "-" & sNomDecoupage & ", " & iNbContrainte.ToString & "/" & iNbTotalContrainte.ToString
                    'Définir l'identifiant du traitement avec un découpage
                    oContrainte.Identifiant = sNomDecoupage
                End If

                'Exécuter la contrainte d'intégrité
                oContrainte.Executer()

                'Conserver le nombre total d'éléments traités
                giNombreTotalElements = giNombreTotalElements + oContrainte.NombreElements

                'Vérifier la présence d'erreurs
                If oContrainte.NombreErreurs > 0 Then
                    'Conserver le nombre total d'erreurs trouvées
                    giNombreTotalErreurs = giNombreTotalErreurs + oContrainte.NombreErreurs

                    'Écrire le message du résultat obtenu de la contrainte dans le rapport d'erreurs
                    If gsNomRapportErreurs.Length > 0 Then File.AppendAllText(gsNomRapportErreurs, oContrainte.Information)

                    'Vérifier si on doit ajouter les erreurs dans la map active
                    If gbAjouterErreursMap And gpMap IsNot Nothing Then
                        'Vérifier si le nom du FeatureLayer d'erreur ne contient pas le nom du découpage
                        If sNomDecoupage.Length > 0 And Not oContrainte.FeatureLayerErreur.Name.Contains(sNomDecoupage) Then
                            'Ajouter le nom du découpage dans le nom du FeatureLayer d'erreur
                            oContrainte.FeatureLayerErreur.Name = sNomDecoupage & ":" & oContrainte.FeatureLayerErreur.Name
                        End If
                        'Ajouter Le FeatureLayer d'erreurs dans la Map active
                        gpMap.AddLayer(oContrainte.FeatureLayerErreur)
                        'Ajouter Le FeatureLayer de sélection dans la Map active
                        'gpMap.AddLayer(oContrainte.FeatureLayerSelection)
                    End If

                    'Vérifier si une erreur d'exécution est survenue
                ElseIf oContrainte.NombreErreurs = -1 Then
                    'Écrire le message du résultat obtenu de la contrainte dans le rapport d'erreurs
                    If gsNomRapportErreurs.Length > 0 Then File.AppendAllText(gsNomRapportErreurs, oContrainte.Information)
                    'Définir le code d'erreur d'exécution sans arrêt du programme
                    giCodeErreur = 1
                End If

                'Récupération de la mémoire disponible
                oContrainte = Nothing
                GC.Collect()

                'Vérifier si un Cancel a été effectué
                If gpTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                'Extraire le prochain Oid des contraintes sélectionnées
                iOid = pEnumIDs.Next()
            Loop

        Catch ex As CancelException
            'Retourner l'exception
            Throw
        Catch ex As Exception
            'Écrire l'erreur
            EcrireMessage("[clsTableContraintes.Executer] " & ex.Message)
            'Définir le code d'erreur d'exécution avec arrêt du programme
            giCodeErreur = -1
        Finally
            'Vider la mémoire
            pTableSel = Nothing
            pRow = Nothing
            oContrainte = Nothing
            pTableSort = Nothing
            pEnumIDs = Nothing
            sNomRepertoireErreurs = Nothing
            iOid = Nothing
            iNbContrainte = Nothing
            iNbTotalContrainte = Nothing
            'Récupération de la mémoire disponible
            GC.Collect()
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de définir le nom de la classe de découpage.
    '''</summary>
    '''
    '''<param name="sNomClasseDecoupage"> Nom de la classe de découpage à traiter.</param>
    ''' 
    '''<remarks>
    ''' Si le nom de la classe de découpage correspond à Layer sur disque, le FeatureLayer de découpage sera défini.
    '''</remarks>
    ''' 
    Private Sub DefinirClasseDecoupage(ByVal sNomClasseDecoupage As String)
        'Déclaration des variables de travail
        Dim pLayerFile As ILayerFile = Nothing          'Interface qui permet de lire un FeatureLayer sur disque.
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface contenant les éléments sélectionnés
        Dim pGeodatabase As IFeatureWorkspace = Nothing 'Interface contenant la Géodatabase de la classe de découpage.
        Dim sRepArcCatalog As String = ""               'Nom du répertoire contenant les connexions des Géodatabase .sde.
        Dim sNomGeodatabase As String = ""              'Nom de la Géodatabase de la classe de découpage.

        Try
            'Extraire le nom du répertoire contenant les connexions des Géodatabase .sde.
            sRepArcCatalog = IO.Directory.GetDirectories(Environment.GetEnvironmentVariable("APPDATA"), "ArcCatalog", IO.SearchOption.AllDirectories)(0)

            'Redéfinir le nom complet de la Géodatabase .sde
            gsNomClasseDecoupage = gsNomClasseDecoupage.ToLower.Replace("database connections", sRepArcCatalog)

            'Vérifier si le nom de la classe de découpage est un Layer de découpage
            If sNomClasseDecoupage.Contains(".lyr") Then
                'Interface pour ouvrir le FeatureLayer de découpage
                pLayerFile = New LayerFile

                'Vérifier si le FeatureLayer de découpage est valide
                If pLayerFile.IsLayerFile(sNomClasseDecoupage) Then
                    'Ouvrir le FeatureLayer de découpage sur disque
                    pLayerFile.Open(sNomClasseDecoupage)

                    'Définir le FeatureLayer de découpage
                    gpFeatureLayerDecoupage = CType(pLayerFile.Layer, IFeatureLayer)

                    'Vérifier si la classe de découpage est valide
                    If gpFeatureLayerDecoupage.FeatureClass Is Nothing Then
                        'Retourner l'erreur
                        Err.Raise(-1, , "ERREUR : Classe de découpage invalide !")
                    End If

                    'Interface pour vérifier si des éléments sont sélectionnés
                    pFeatureSel = CType(gpFeatureLayerDecoupage, IFeatureSelection)

                    'Vérifier si des éléments sont sélectionnés
                    If pFeatureSel.SelectionSet.Count = 0 Then
                        'Retourner l'erreur
                        Err.Raise(-1, , "ERREUR : Aucun élément de découpage n'est sélectionné !")
                    End If

                    'Conserver le nombre totale de découpage à traiter
                    giNombreTotalDecoupages = pFeatureSel.SelectionSet.Count

                    'Si le Layer de découpage est invalide
                Else
                    'Retourner l'erreur
                    Err.Raise(-1, , "ERREUR : Layer de découpage invalide !")
                End If

                'Vérifier si le nom de la classe de découpagecontient une requête
            ElseIf sNomClasseDecoupage.Contains(":") Then
                'Définir la Géodatabase par défaut
                pGeodatabase = gpGeodatabaseClasses

                'Vérifier si la classe de découpage contient le nom de la Géodatabase
                If sNomClasseDecoupage.Contains("\") Then
                    'Définir le nom de la Géodatabase de la classe de découpage
                    sNomGeodatabase = sNomClasseDecoupage.Split(CChar("\"))(0)
                    'Définir le nom de la Géodatabase de la classe de découpage
                    sNomClasseDecoupage = sNomClasseDecoupage.Split(CChar("\"))(1)
                    'Définir la géodatabase de la classe de découpage
                    pGeodatabase = CType(DefinirGeodatabase(sNomGeodatabase), IFeatureWorkspace)
                End If

                'Définir le FeatureLayer de découpage
                gpFeatureLayerDecoupage = CreerFeatureLayer(sNomClasseDecoupage, pGeodatabase)

                'Vérifier si la classe de découpage est valide
                If gpFeatureLayerDecoupage.FeatureClass Is Nothing Then
                    'Retourner l'erreur
                    Err.Raise(-1, , "ERREUR : Classe de découpage invalide !")
                End If

                'Interface pour vérifier si des éléments sont sélectionnés
                pFeatureSel = CType(gpFeatureLayerDecoupage, IFeatureSelection)

                'Vérifier si des éléments sont sélectionnés
                If pFeatureSel.SelectionSet.Count = 0 Then
                    'Retourner l'erreur
                    Err.Raise(-1, , "ERREUR : Aucun élément de découpage n'est sélectionné !")
                End If

                'Conserver le nombre totale de découpage à traiter
                giNombreTotalDecoupages = pFeatureSel.SelectionSet.Count
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pLayerFile = Nothing
            pFeatureSel = Nothing
            pGeodatabase = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de définir la table contenant les contraintes d'intégrité dans un StandaloneTable.
    '''</summary>
    ''' 
    '''<param name="sNomTableContraintes"> Nom de la tables contenant les contraintes d'intégrité.</param>
    '''<param name="pMap"> Map active dans laquelle la table des contraintes est présentes.</param>
    ''' 
    Private Sub DefinirTableContraintes(ByVal sNomTableContraintes As String, ByVal pMap As IMap)
        'Déclaration des variables de travail
        Dim pWorkspace As IWorkspace = Nothing      'Interface contenant une Géodatabase.
        Dim pDataset As IDataset = Nothing          'Interface contenant le nom de la table des contraintes.
        Dim pTableDef As ITableDefinition = Nothing 'Interface qui permet d'appliquer un filtre sur les contraintes.
        Dim pTableSel As ITableSelection = Nothing  'Interface qui permet d'extraire le nombre de contraintes sélectionnées.
        Dim pStandaloneColl As IStandaloneTableCollection = Nothing 'Collection des tables dans la Map active.
        Dim sNomTable As String = ""                'Contient le nom de la table.

        Try
            'Initialiser les variables
            gsNomTableContraintes = sNomTableContraintes
            gsFiltreContraintes = ""
            gpTableContraintes = Nothing

            'Interface pour axtraire la table des contraintes
            pStandaloneColl = CType(pMap, IStandaloneTableCollection)

            'Traiter toutes les tables
            For i = 0 To pStandaloneColl.StandaloneTableCount - 1
                'Interface pour extraire le nom de la table des contraintes
                pDataset = CType(pStandaloneColl.StandaloneTable(i), IDataset)
                'Vérifier la présence du Path de la Géodatabase
                If pDataset.Workspace.PathName.Length = 0 Then
                    'Définir le nom complet de la table
                    sNomTable = pDataset.Name
                Else
                    'Définir le nom complet de la table
                    sNomTable = pDataset.Workspace.PathName & "\" & pDataset.Name
                End If

                'Vérifier si le nom correspond à la table
                If sNomTable.ToLower() = sNomTableContraintes.ToLower() Then
                    'Conserver la table des contraintes
                    gpTableContraintes = CType(pDataset, IStandaloneTable)

                    'Interface pour appliquer un filtre
                    pTableDef = CType(gpTableContraintes, ITableDefinition)

                    'Conserver le filtre su rla table des contraintes
                    gsFiltreContraintes = pTableDef.DefinitionExpression

                    'Interface pour extraire le nombre de contraintes sélectionnées
                    pTableSel = CType(gpTableContraintes, ITableSelection)
                    'Vérifier si aucune contrainte sélectionnées
                    If pTableSel.SelectionSet.Count = 0 Then
                        'Sélectionner toutes les contraintes
                        pTableSel.SelectRows(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                    End If
                    'Conserver le nombre total de contraintes à traiter
                    giNombreTotalContraintes = pTableSel.SelectionSet.Count

                    'Sortir de la boucle
                    Exit For
                End If
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pWorkspace = Nothing
            pDataset = Nothing
            pTableDef = Nothing
            pTableSel = Nothing
            pStandaloneColl = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de définir la table contenant les contraintes d'intégrité dans un StandaloneTable.
    '''</summary>
    ''' 
    '''<param name="sNomTableContraintes"> Nom de la tables contenant les contraintes d'intégrité.</param>
    '''<param name="sFiltreContraintes"> Filtre à appliquer sur les contraintes d'intégrité.</param>
    '''<param name="sNomProprietaireDefaut"> Contient le nom du propriétaire des tables par défaut pour les Géodatabase Enterprise.</param>
    ''' 
    Private Sub DefinirTableContraintes(ByVal sNomTableContraintes As String, Optional ByVal sFiltreContraintes As String = "",
                                        Optional ByVal sNomProprietaireDefaut As String = "BDG_DBA")
        'Déclaration des variables de travail
        Dim pWorkspace2 As IWorkspace2 = Nothing    'Interface pour vérifier si la table existe.
        Dim pWorkspace As IWorkspace = Nothing      'Interface pour vérifier le type de Géodatabase.
        Dim pTable As ITable = Nothing              'Interface contenant la table des contraintes.
        Dim pTableDef As ITableDefinition = Nothing 'Interface qui permet d'appliquer un filtre sur les contraintes.
        Dim pTableSel As ITableSelection = Nothing  'Interface qui permet d'extraire le nombre de contraintes sélectionnées.
        Dim sNomTable As String = ""                'Contient le nom de la table des contraintes.
        Dim sNomGeodatabase As String = ""          'Contient le nom de la Géodatabase.
        Dim sRepArcCatalog As String = ""           'Nom du répertoire contenant les connexions des Géodatabase .sde.

        Try
            'Extraire le nom du répertoire contenant les connexions des Géodatabase .sde.
            sRepArcCatalog = IO.Directory.GetDirectories(Environment.GetEnvironmentVariable("APPDATA"), "ArcCatalog", IO.SearchOption.AllDirectories)(0)

            'Redéfinir le nom complet de la Géodatabase .sde
            sNomTableContraintes = sNomTableContraintes.ToLower.Replace("database connections", sRepArcCatalog)

            'Définir le nom de la table des contraintes sans le nom de la Géodatabase
            sNomTable = System.IO.Path.GetFileName(sNomTableContraintes)

            'Définir le nom de la géodatabase sans celui du nom de la table
            sNomGeodatabase = sNomTableContraintes.Replace("\" & sNomTable, "")
            sNomGeodatabase = sNomGeodatabase.Replace(sNomTable, "")

            'Vérifier si le nom de la Géodatabase de la table est absent
            If sNomGeodatabase.Length = 0 Then
                'Définir la Géodatabase de la table des contraintes à partir de celle des classes
                gpGeodatabaseContraintes = gpGeodatabaseClasses
                'Interface pour extraire le nom de la Géodatabase
                pWorkspace = CType(gpGeodatabaseContraintes, IWorkspace)
                'Définir le nom de la Géodatabase
                sNomGeodatabase = pWorkspace.PathName

                'Si le nom de la Géodatabase est présent
            Else
                'Ouvrir la géodatabase de la table des contraintes
                gpGeodatabaseContraintes = CType(DefinirGeodatabase(sNomGeodatabase), IFeatureWorkspace)

                'Vérifier si la Géodatabase est valide
                If gpGeodatabaseContraintes Is Nothing Then
                    'Retourner l'erreur
                    Err.Raise(-1, , "ERREUR : Le nom de la Géodatabase est invalide : " & sNomGeodatabase)
                End If
            End If

            'Interface pour vérifier le type de Géodatabase.
            pWorkspace = CType(gpGeodatabaseContraintes, IWorkspace)
            'Vérifier si la Géodatabase est de type "Enterprise" 
            If pWorkspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                'Vérifier si le nom de la table contient le nom du propriétaire
                If Not sNomTable.Contains(".") Then
                    'Définir le nom de la table avec le nom du propriétaire
                    sNomTable = sNomProprietaireDefaut & "." & sNomTable
                End If
            End If

            'Interface pour vérifier si la table existe
            pWorkspace2 = CType(gpGeodatabaseContraintes, IWorkspace2)
            'Vérifier si la table existe
            If pWorkspace2.NameExists(esriDatasetType.esriDTTable, sNomTable) Then
                'Ouvrir la table des contraintes
                pTable = gpGeodatabaseContraintes.OpenTable(sNomTable)
                'Si la table n'existe pas
            Else
                'Retourner une erreur
                Throw New Exception("ERREUR : Incapable d'ouvrir la table des contraintes : " & sNomTable)
            End If

            'Créer le StandaloneTable de la table des contraintes d'intégrité
            gpTableContraintes = New StandaloneTable

            'Définir la table dans le StandaloneTable des contraintes d'intégrité 
            gpTableContraintes.Table = pTable

            'Conserver le filtre
            gsFiltreContraintes = sFiltreContraintes
            'Vérifier si un filtre doit être appliqué
            If sFiltreContraintes.Length > 0 Then
                'Interface pour appliquer un filtre
                pTableDef = CType(gpTableContraintes, ITableDefinition)
                'Ajouter le filtre
                pTableDef.DefinitionExpression = sFiltreContraintes
            End If

            'Interface pour extraire le nombre de contraintes sélectionnées
            pTableSel = CType(gpTableContraintes, ITableSelection)
            'Vérifier si aucune contrainte sélectionnées
            If pTableSel.SelectionSet.Count = 0 Then
                'Sélectionner toutes les contraintes
                pTableSel.SelectRows(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If
            'Conserver le nombre total de contraintes à traiter
            giNombreTotalContraintes = pTableSel.SelectionSet.Count

            'Définir le nom complet de la table des contraintes d'intégrité
            gsNomTableContraintes = sNomGeodatabase & "\" & gpTableContraintes.Name

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pWorkspace2 = Nothing
            pWorkspace = Nothing
            pTable = Nothing
            pTableDef = Nothing
            pTableSel = Nothing
            sNomTable = Nothing
            sNomGeodatabase = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de définir la table contenant les statistiques d'erreurs et de traitement.
    '''</summary>
    ''' 
    '''<param name="sNomTableStatistiques"> Nom de la table contenant les statistiques d'erreurs et de traitement.</param>
    '''<param name="sNomProprietaireDefaut"> Contient le nom du propriétaire des tables par défaut pour les Géodatabase Enterprise.</param>
    ''' 
    Private Sub DefinirTableStatistiques(ByVal sNomTableStatistiques As String, Optional ByVal sNomProprietaireDefaut As String = "BDG_DBA")
        'Déclaration des variables de travail
        Dim pWorkspace2 As IWorkspace2 = Nothing    'Interface pour vérifier si la table existe.
        Dim pWorkspace As IWorkspace = Nothing      'Interface pour vérifier le type de Géodatabase.
        Dim pTable As ITable = Nothing              'Interface contenant la table des statistiques.
        Dim sNomTable As String = ""                'Contient le nom de la table des statistiques.
        Dim sNomGeodatabase As String = ""          'Contient le nom de la Géodatabase.
        Dim sRepArcCatalog As String = ""           'Nom du répertoire contenant les connexions des Géodatabase .sde.

        Try
            'Vérifier si le nom de la table des statistiques est présent
            If sNomTableStatistiques.Length > 0 Then
                'Extraire le nom du répertoire contenant les connexions des Géodatabase .sde.
                sRepArcCatalog = IO.Directory.GetDirectories(Environment.GetEnvironmentVariable("APPDATA"), "ArcCatalog", IO.SearchOption.AllDirectories)(0)

                'Redéfinir le nom complet de la Géodatabase .sde
                sNomTableStatistiques = sNomTableStatistiques.ToLower.Replace("database connections", sRepArcCatalog)

                'Définir le nom de la table des statistiques sans le nom de la Géodatabase
                sNomTable = System.IO.Path.GetFileName(sNomTableStatistiques)

                'Définir le nom de la géodatabase sans celui du nom de la table
                sNomGeodatabase = sNomTableStatistiques.Replace("\" & sNomTable, "")
                sNomGeodatabase = sNomGeodatabase.Replace(sNomTable, "")

                'Vérifier si le nom de la Géodatabase de la table est absent
                If sNomGeodatabase.Length = 0 Then
                    'Définir la Géodatabase de la table des statistiques à partir de celle des classes
                    gpGeodatabaseStatistiques = gpGeodatabaseClasses
                    'Interface pour extraire le nom de la Géodatabase
                    pWorkspace = CType(gpGeodatabaseStatistiques, IWorkspace)
                    'Définir le nom de la Géodatabase
                    sNomGeodatabase = pWorkspace.PathName

                    'Si le nom de la Géodatabase est présent
                Else
                    'Ouvrir la géodatabase de la table des statistiques
                    gpGeodatabaseStatistiques = CType(DefinirGeodatabase(sNomGeodatabase), IFeatureWorkspace)

                    'Vérifier si la Géodatabase est valide
                    If gpGeodatabaseStatistiques Is Nothing Then
                        'Retourner l'erreur
                        Err.Raise(-1, , "ERREUR : Le nom de la Géodatabase est invalide : " & sNomGeodatabase)
                    End If
                End If

                'Interface pour vérifier le type de Géodatabase.
                pWorkspace = CType(gpGeodatabaseStatistiques, IWorkspace)
                'Vérifier si la Géodatabase est de type "Enterprise" 
                If pWorkspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                    'Vérifier si le nom de la table contient le nom du propriétaire
                    If Not sNomTable.Contains(".") Then
                        'Définir le nom de la table avec le nom du propriétaire
                        sNomTable = sNomProprietaireDefaut & "." & sNomTable
                    End If
                End If

                'Interface pour vérifier si la table existe
                pWorkspace2 = CType(gpGeodatabaseStatistiques, IWorkspace2)
                'Vérifier si la table existe
                If pWorkspace2.NameExists(esriDatasetType.esriDTTable, sNomTable) Then
                    'Ouvrir la table des statistiques
                    pTable = gpGeodatabaseStatistiques.OpenTable(sNomTable)
                    'Si la table n'existe pas
                Else
                    'Retourner l'erreur
                    Throw New Exception("ERREUR : Incapable d'ouvrir la table des statistiques : " & sNomTable)
                End If

                'Définir la table des statistiques
                gpTableStatistiques = pTable
                'Définir le nom complet de la table des statistiques
                gsNomTableStatistiques = sNomGeodatabase & "\" & sNomTable
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pWorkspace2 = Nothing
            pWorkspace = Nothing
            pTable = Nothing
            sNomTable = Nothing
            sNomGeodatabase = Nothing
        End Try
    End Sub

    '''<summary>
    '''Fonction qui permet d'ouvrir et retourner la Geodatabase à partir d'un nom de Géodatabase.
    '''</summary>
    '''
    '''<param name="sNomGeodatabase"> Nom de la géodatabase à ouvrir.</param>
    ''' 
    Private Function DefinirGeodatabase(ByRef sNomGeodatabase As String) As IWorkspace
        'Déclaration des variables de travail
        Dim pFactoryType As Type = Nothing                      'Interface utilisé pour définir le type de géodatabase.
        Dim pWorkspaceFactory As IWorkspaceFactory2 = Nothing   'Interface utilisé pour ouvrir la géodatabase.
        Dim sRepArcCatalog As String = ""                       'Nom du répertoire contenant les connexions des Géodatabase .sde.

        'Par défaut, aucune Géodatabase n'est retournée
        DefinirGeodatabase = Nothing

        Try
            'Valider le paramètre de la Geodatabase
            If sNomGeodatabase.Length > 0 Then
                'Extraire le nom du répertoire contenant les connexions des Géodatabase .sde.
                sRepArcCatalog = IO.Directory.GetDirectories(Environment.GetEnvironmentVariable("APPDATA"), "ArcCatalog", IO.SearchOption.AllDirectories)(0)

                'Redéfinir le nom complet de la Géodatabase .sde
                sNomGeodatabase = sNomGeodatabase.ToLower.Replace("database connections", sRepArcCatalog)

                'Vérifier si le nom est une geodatabase SDE PRO prédéfinie
                If sNomGeodatabase = "bdrs_pro_bdg_dba" Then
                    'Définir le type de workspace : SDE
                    pFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.SdeWorkspaceFactory")

                    'Interface pour ouvrir le Workspace
                    pWorkspaceFactory = CType(Activator.CreateInstance(pFactoryType), IWorkspaceFactory2)
                    Try
                        'Ouvrir le workspace de la Géodatabase
                        DefinirGeodatabase = pWorkspaceFactory.OpenFromString("INSTANCE=sde:oracle11g:bdrs_pro;USER=BDG_DBA;PASSWORD=123bdg_dba;VERSION=sde.DEFAULT", 0)
                    Catch ex As Exception
                        'Retourner l'erreur
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Incapable d'ouvrir la Géodatabase : " & sNomGeodatabase)
                    End Try

                    'Vérifier si le nom est une geodatabase SDE TST prédéfinie
                ElseIf sNomGeodatabase = "bdrs_tst_bdg_dba" Then
                    'Définir le type de workspace : SDE
                    pFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.SdeWorkspaceFactory")
                    'Interface pour ouvrir le Workspace
                    pWorkspaceFactory = CType(Activator.CreateInstance(pFactoryType), IWorkspaceFactory2)
                    Try
                        'Ouvrir le workspace de la Géodatabase
                        DefinirGeodatabase = pWorkspaceFactory.OpenFromString("INSTANCE=sde:oracle11g:bdrs_tst;USER=BDG_DBA;PASSWORD=tst;VERSION=sde.DEFAULT", 0)
                    Catch ex As Exception
                        'Retourner l'erreur
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Incapable d'ouvrir la Géodatabase : " & sNomGeodatabase)
                    End Try

                    'Vérifier si le nom est une geodatabase SDE
                ElseIf IO.Path.GetExtension(sNomGeodatabase) = ".sde" Then
                    'Définir le type de workspace : SDE
                    pFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.SdeWorkspaceFactory")
                    'Interface pour ouvrir le Workspace
                    pWorkspaceFactory = CType(Activator.CreateInstance(pFactoryType), IWorkspaceFactory2)
                    Try
                        'Ouvrir le workspace de la Géodatabase
                        DefinirGeodatabase = pWorkspaceFactory.OpenFromFile(sNomGeodatabase, 0)
                    Catch ex As Exception
                        'Retourner l'erreur
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Incapable d'ouvrir la Géodatabase : " & sNomGeodatabase)
                    End Try

                    'Si la Geodatabse est une File Geodatabase
                ElseIf sNomGeodatabase.Contains(".gdb") Then
                    'Définir le type de workspace : SDE
                    pFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory")
                    'Interface pour ouvrir le Workspace
                    pWorkspaceFactory = CType(Activator.CreateInstance(pFactoryType), IWorkspaceFactory2)
                    Try
                        'Ouvrir le workspace de la Géodatabase
                        DefinirGeodatabase = pWorkspaceFactory.OpenFromFile(sNomGeodatabase, 0)
                    Catch ex As Exception
                        'Retourner l'erreur
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Incapable d'ouvrir la Géodatabase : " & sNomGeodatabase)
                    End Try

                    'Si la Geodatabse est une personnelle Geodatabase
                ElseIf sNomGeodatabase.Contains(".mdb") Then
                    'Définir le type de workspace : SDE
                    pFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.AccessWorkspaceFactory")
                    'Interface pour ouvrir le Workspace
                    pWorkspaceFactory = CType(Activator.CreateInstance(pFactoryType), IWorkspaceFactory2)
                    Try
                        'Ouvrir le workspace de la Géodatabase
                        DefinirGeodatabase = pWorkspaceFactory.OpenFromFile(sNomGeodatabase, 0)
                    Catch ex As Exception
                        'Retourner l'erreur
                        Throw New Exception(ex.Message & vbCrLf & "ERREUR : Incapable d'ouvrir la Géodatabase : " & sNomGeodatabase)
                    End Try

                    'Sinon
                Else
                    'Retourner l'erreur
                    Err.Raise(-1, , "ERREUR : Le nom de la Géodatabase ne correspond pas à une Geodatabase !")
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFactoryType = Nothing
            pWorkspaceFactory = Nothing
        End Try
    End Function

    '''<summary>
    '''Routine qui permet de sélectionner les contraintes de la table à partir d'un Noeud du TreeNode.
    '''</summary>
    '''
    '''<param name="oTreeNode"> TreeNode utilisé pour sélectionner les contraintes de la table.</param>
    '''<param name="pMap"> Map active dans laquelle la table des contraintes est présente.</param>
    ''' 
    Private Sub SelectionnerContraintes(ByVal oTreeNode As TreeNode, ByVal pMap As IMap)
        'Déclaration des variables de travail
        Dim pTableSel As ITableSelection = Nothing  'Interface qui permet de sélectionner les contraintes à traiter.

        Try
            'Vérifier si le TreeNode est de type [TABLE]
            If oTreeNode.Tag.ToString = "[TABLE]" Then
                'Définir le Standalone de la table des contraintes d'intégrité
                DefinirTableContraintes(oTreeNode.Text, pMap)

                'Vérifier si la table des contraintes est valide
                If gpTableContraintes IsNot Nothing Then
                    'Interface qui permet de sélectionner les contraintes à traiter.
                    pTableSel = CType(gpTableContraintes, ITableSelection)

                    'Vider la sélection
                    pTableSel.Clear()

                    'Sélectionner les contraintes à partir du noeud de type [TABLE]
                    SelectionnerNoeudTable(oTreeNode, pTableSel)
                End If

                'Si le TreeNode est de type [TRIER]
            ElseIf oTreeNode.Tag.ToString.Contains("[TRIER]") Then
                'Définir le Standalone de la table des contraintes d'intégrité
                DefinirTableContraintes(oTreeNode.Parent.Text, pMap)

                'Vérifier si la table des contraintes est valide
                If gpTableContraintes IsNot Nothing Then
                    'Vérifier si aucun filtre n'est présent
                    If gsFiltreContraintes.Length = 0 Then
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = oTreeNode.Tag.ToString.Replace("[TRIER]:", "") & "='" & oTreeNode.Text & "'"

                        'Si un filtre est déjà présent
                    Else
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = gsFiltreContraintes & " AND " & oTreeNode.Tag.ToString.Replace("[TRIER]:", "") & "='" & oTreeNode.Text & "'"
                    End If

                    'Interface qui permet de sélectionner les contraintes à traiter.
                    pTableSel = CType(gpTableContraintes, ITableSelection)

                    'Vider la sélection
                    pTableSel.Clear()

                    'Sélectionner les contraintes à partir du noeud de type [TRIER]
                    SelectionnerNoeudTrier(oTreeNode, pTableSel)
                End If

                'Si le TreeNode est de type [CONTRAINTE]
            ElseIf oTreeNode.Tag.ToString.Contains("[CONTRAINTE]") Then
                'Définir le Standalone de la table des contraintes d'intégrité
                DefinirTableContraintes(oTreeNode.Parent.Parent.Text, pMap)

                'Vérifier si la table des contraintes est valide
                If gpTableContraintes IsNot Nothing Then
                    'Vérifier si aucun filtre n'est présent
                    If gsFiltreContraintes.Length = 0 Then
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = gpTableContraintes.Table.OIDFieldName & "=" & oTreeNode.FirstNode.FirstNode.Text

                        'Si un filtre est déjà présent
                    Else
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = gsFiltreContraintes & " AND " & gpTableContraintes.Table.OIDFieldName & "=" & oTreeNode.FirstNode.FirstNode.Text
                    End If

                    'Interface qui permet de sélectionner les contraintes à traiter.
                    pTableSel = CType(gpTableContraintes, ITableSelection)

                    'Vider la sélection
                    pTableSel.Clear()

                    'Sélectionner les contraintes à partir du noeud de type [CONTRAINTE]
                    SelectionnerNoeudContrainte(oTreeNode, pTableSel)
                End If

                'Si le TreeNode est de type [REQUETES]
            ElseIf oTreeNode.Tag.ToString.Contains("[REQUETES]") Then
                'Définir le Standalone de la table des contraintes d'intégrité
                DefinirTableContraintes(oTreeNode.Parent.Parent.Parent.Text, pMap)

                'Vérifier si la table des contraintes est valide
                If gpTableContraintes IsNot Nothing Then
                    'Vérifier si aucun filtre n'est présent
                    If gsFiltreContraintes.Length = 0 Then
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = gpTableContraintes.Table.OIDFieldName & "=" & oTreeNode.Parent.FirstNode.FirstNode.Text

                        'Si un filtre est déjà présent
                    Else
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = gsFiltreContraintes & " AND " & gpTableContraintes.Table.OIDFieldName & "=" & oTreeNode.Parent.FirstNode.FirstNode.Text
                    End If

                    'Interface qui permet de sélectionner les contraintes à traiter.
                    pTableSel = CType(gpTableContraintes, ITableSelection)

                    'Vider la sélection
                    pTableSel.Clear()

                    'Sélectionner les contraintes à partir du noeud de type [CONTRAINTE]
                    SelectionnerNoeudContrainte(oTreeNode.Parent, pTableSel)
                End If

                'Si le TreeNode est de type [REQUETE]
            ElseIf oTreeNode.Tag.ToString.Contains("[REQUETE]") Then
                'Définir le Standalone de la table des contraintes d'intégrité
                DefinirTableContraintes(oTreeNode.Parent.Parent.Parent.Parent.Text, pMap)

                'Vérifier si la table des contraintes est valide
                If gpTableContraintes IsNot Nothing Then
                    'Vérifier si aucun filtre n'est présent
                    If gsFiltreContraintes.Length = 0 Then
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = gpTableContraintes.Table.OIDFieldName & "=" & oTreeNode.Parent.Parent.FirstNode.FirstNode.Text

                        'Si un filtre est déjà présent
                    Else
                        'Définir le filtre sur les contraintes
                        gsFiltreContraintes = gsFiltreContraintes & " AND " & gpTableContraintes.Table.OIDFieldName & "=" & oTreeNode.Parent.Parent.FirstNode.FirstNode.Text
                    End If

                    'Interface qui permet de sélectionner les contraintes à traiter.
                    pTableSel = CType(gpTableContraintes, ITableSelection)

                    'Vider la sélection
                    pTableSel.Clear()

                    'Sélectionner les contraintes à partir du noeud de type [CONTRAINTE]
                    SelectionnerNoeudContrainte(oTreeNode.Parent.Parent, pTableSel)
                End If

                'Sinon
            Else
                'Retourner l'erreur
                Err.Raise(-1, , "ERREUR : Le type de TreeNode est invalide : " & oTreeNode.Tag.ToString)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pTableSel = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de sélectionner les contraintes de la table à partir d'un Noeud de type [TABLE].
    '''</summary>
    '''
    '''<param name="oTreeNode"> TreeNode utilisé pour sélectionner les contraintes de la table.</param>
    ''' 
    Private Sub SelectionnerNoeudTable(ByVal oTreeNode As TreeNode, ByRef pTableSel As ITableSelection)
        'Déclaration des variables de travail
        Dim oNode As TreeNode = Nothing  'Objet contenant un noeud de type [TRIER]

        Try
            'Traiter tous les noeudsde type [TRIER]
            For Each oNode In oTreeNode.Nodes
                'Sélectionner les contraintes à partir du noeud de type [TRIER]
                SelectionnerNoeudTrier(oNode, pTableSel)
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oNode = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de sélectionner les contraintes de la table à partir d'un Noeud de type [TRIER].
    '''</summary>
    '''
    '''<param name="oTreeNode"> TreeNode utilisé pour sélectionner les contraintes de la table.</param>
    ''' 
    Private Sub SelectionnerNoeudTrier(ByVal oTreeNode As TreeNode, ByRef pTableSel As ITableSelection)
        'Déclaration des variables de travail
        Dim oNode As TreeNode = Nothing  'Objet contenant un noeud de type [TRIER]

        Try
            'Traiter tous les noeudsde type [TRIER]
            For Each oNode In oTreeNode.Nodes
                'Sélectionner les contraintes à partir du noeud de type [CONTRAINTE]
                SelectionnerNoeudContrainte(oNode, pTableSel)
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oNode = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de sélectionner une contrainte de la table à partir d'un Noeud de type [CONTRAINTE].
    '''</summary>
    '''
    '''<param name="oTreeNode"> TreeNode utilisé pour sélectionner les contraintes de la table.</param>
    ''' 
    Private Sub SelectionnerNoeudContrainte(ByVal oTreeNode As TreeNode, ByRef pTableSel As ITableSelection)
        'Déclaration des variables de travail
        Dim oNodeOid As TreeNode = Nothing  'Objet contenant un noeud de type [VALEUR] contenant le OBJECTID

        Try
            'Définir le noeud de la contrainte contenant la valeur de l'attribut OBJECTID
            oNodeOid = oTreeNode.FirstNode.FirstNode

            'Sélectionner la contrainte à partir du noeud de type [VALEUR] contenant le OBJECTID
            pTableSel.SelectionSet.Add(CInt(oNodeOid.Text))

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oNodeOid = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'écrire les statistiques d'utilisation d'un usager.
    ''' 
    '''<param name="sCommande"> Commande à écrire dans le fichier de statistique d'utilisation.</param>
    '''<param name="sNomRepertoire"> Nom du répertoire dans lequel le fichier de statistique est présent.</param>
    ''' 
    '''</summary>
    '''
    Protected Sub EcrireStatistiqueUtilisation(ByVal sCommande As String, Optional ByVal sNomRepertoire As String = "S:\Developpement\geo\")
        'Déclarer les variables de travail
        Dim oStreamWriter As StreamWriter = Nothing     'Objet utilisé pour écrire dans un fichier text.
        Dim sNomFichier As String = ""                  'Nom complet du fichier de statistique d'utilisation.
        Dim sNomUsager As String = ""                   'Nom de l'usager.

        Try
            'Définir le nom de l'usager
            sNomUsager = Environment.GetEnvironmentVariable("USERNAME")

            'Définir le nom complet du fichier
            sNomFichier = sNomRepertoire & sNomUsager & ".txt"

            'Vérifier si le fichier existe
            If File.Exists(sNomFichier) Then
                'Définir l'objet pour écrire à la fin du fichier
                oStreamWriter = File.AppendText(sNomFichier)

                'Si le fichier n'existe pas
            Else
                'Définir l'objet pour écrire dans un nouveau fichier créé
                oStreamWriter = File.CreateText(sNomFichier)

                'Écrire l'entête du fichier
                oStreamWriter.WriteLine("Date, 	 Env, 	 Usager, 	 UsagerBD, 	 UsagerSIB, 	 Outil")
            End If

            'Écrire la commande utilisée
            oStreamWriter.WriteLine(DateTime.Now.ToString & "," & vbTab & System.IO.Path.GetFileName(System.Environment.GetCommandLineArgs()(0)) & "," & vbTab & sNomUsager & "," & vbTab & "NONE," & vbTab & "NONE," & vbTab & sCommande)

            'Fermer le fichier
            oStreamWriter.Close()

        Catch ex As Exception
            'Retourner l'erreur
            'Throw ex
        Finally
            'Vider la mémoire
            oStreamWriter = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'écrire le message d'exécution dans un RichTextBox, un fichier journal/Rapport et/ou dans la console.
    '''</summary>
    ''' 
    '''<param name="sMessage"> Message à écrire dans un RichTextBox, un fichier journal et/ou dans la console.</param>
    '''
    Private Sub EcrireMessage(ByVal sMessage As String)
        Try
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

            'Vérifier si on doit écrire dans le rapport d'erreurs
            If gsNomRapportErreurs.Length > 0 Then
                'Écrire le message dans le rapport d'erreurs
                File.AppendAllText(gsNomRapportErreurs, sMessage & vbCrLf)
            End If

            'Écrire dans la console
            Console.WriteLine(sMessage)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

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
    Function CreerFeatureLayer(ByVal sNomClasse As String, ByVal pFeatureWorkspace As IFeatureWorkspace, _
                               Optional ByVal pSpatialReference As ISpatialReference = Nothing, _
                               Optional ByVal pQueryFilter As IQueryFilter = Nothing,
                               Optional ByVal sNomProprietaireDefaut As String = "BDG_DBA") As IFeatureLayer
        'Déclaration des variables de travail
        Dim pWorkspace2 As IWorkspace2 = Nothing                    'Interface pour vérifier si la table existe.
        Dim pWorkspace As IWorkspace = Nothing                      'Interface pour vérifier le type de Géodatabase.
        Dim pFeatureClass As IFeatureClass = Nothing                'Interface contenant la classe de la Géodatabase.
        Dim pFeatureLayerDef As IFeatureLayerDefinition = Nothing   'Interface qui permet d'ajouter une condition attributive.
        Dim pFeatureSel As IFeatureSelection = Nothing              'Interface pour sélectionner les éléments.
        Dim sNomGeodatabase As String = ""  'Contient le nom de la Géodatabase.
        Dim sNomLayer As String = ""        'Contient le nom du Layer.
        Dim sNomTable As String = ""        'Contient le nom de la table.
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

                'Définir le nom de la table
                sNomTable = sNomLayer
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
                'Retourner une erreur
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
            sFiltre = Nothing
            sLayer = Nothing
        End Try
    End Function
#End Region
End Class
