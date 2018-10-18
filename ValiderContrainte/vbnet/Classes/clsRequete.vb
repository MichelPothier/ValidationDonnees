Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Windows.Forms

'**
'Nom de la composante : clsRequete.vb
'
'''<summary>
''' Classe générique qui permet de traiter une requête quelconque de contrainte d'intégrité.
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 15 avril 2015
'''</remarks>
''' 
Public Class clsRequete
    'Déclarer les variables globales
    '''<summary>Nom de l'attribut à traiter.</summary>
    Protected Const C_PARAMETRES_DEFAUT As String = "DATASET_NAME ^(\d\d\d[A-P](0[1-9]|1[0-6]))$"

    '''<summary>Interface utilisé pour afficher la barre de progression et pour annuler le traitement.</summary>
    Protected gpTrackCancel As ITrackCancel = Nothing
    '''<summary>RichTextBox utilisé pour écrire le journal d'éxécution du traitement.</summary>
    Protected gqRichTextBox As RichTextBox = Nothing
    '''<summary>Nom du fichier journal d'éxécution du traitement.</summary>
    Protected gsNomFichierJournal As String = ""
    '''<summary>Interface contenant le nom de la classe de découpage à traiter.</summary>
    Protected gsNomClasseDecoupage As String = ""
    '''<summary>Contient l'information de la contrainte d'intégrité.</summary>
    Protected gsInformation As String = ""

    '''<summary>Objet contenant la requête à traiter.</summary>
    Protected goRequete As intRequete = Nothing
    '''<summary>Interface contenant les géométries décrivant les erreurs.</summary>
    Protected gpGeometryBagErreur As IGeometryBag = Nothing
    '''<summary>Interface contenant un élément de découpage à traiter.</summary>
    Protected gpFeatureDecoupage As IFeature = Nothing
    '''<summary>Interface contenant le FeatureLayer des éléments de découpage à traiter.</summary>
    Protected gpFeatureLayerDecoupage As IFeatureLayer = Nothing
    '''<summary>Interface contenant la Géodatabase des classes à traiter.</summary>
    Protected gpFeatureWorkspace As IFeatureWorkspace = Nothing
    '''<summary>Interface contenant la référence spatiale à utiliser pour le traitement.</summary>
    Protected gpSpatialReference As ISpatialReference = Nothing
    '''<summary>Interface contenant la requête attributive à utiliser pour le traitement.</summary>
    Protected gpQueryFilter As IQueryFilter = Nothing
    '''<summary>Collection contenant la liste des Layers traités.</summary>
    Protected gqListeLayers As Collection = Nothing

    '''<summary>Contient le nombre d'éléments sélectionnés au début du traitement.</summary>
    Protected giNombreElementsDebut As Long = 0
    '''<summary>Contient le nombre d'éléments sélectionnés à la fin du traitement.</summary>
    Protected giNombreElementsFin As Long = 0

#Region "Constructeurs"
    '''<summary>
    ''' Routine qui permet d'instancier la requête à traiter selon une commande sous forme de texte.
    '''</summary>
    '''
    '''<param name="sCommande"> Commande qui correspond à une requête à traiter.</param>
    '''<param name="qListeLayers"> Collection des Layers traités et conservés.</param>
    '''<param name="pFeatureWorkspace"> Interface contenant la géodatabase.</param>
    '''<param name="pSpatialReference"> Interface contenant la référence spatiale.</param>
    '''<param name="pQueryFilter"> Interface contenant un filtre de recherche.</param>
    '''<param name="sNomClasseDecoupage"> Nom de la classe de découpage.</param>
    '''<param name="pFeatureLayerDecoupage"> Layer de la classe de découpage.</param>
    '''<param name="pFeatureDecoupage"> Élément de découpage.</param>
    ''' 
    Public Sub New(ByVal sCommande As String, ByRef qListeLayers As Collection, ByVal pFeatureWorkspace As IFeatureWorkspace, _
                   Optional ByVal pSpatialReference As ISpatialReference = Nothing, Optional ByVal pQueryFilter As IQueryFilter = Nothing, _
                   Optional ByVal sNomClasseDecoupage As String = "", Optional ByVal pFeatureLayerDecoupage As IFeatureLayer = Nothing, _
                   Optional ByVal pFeatureDecoupage As IFeature = Nothing)
        'Déclaration des variables de travail
        Dim sChamps() As String = Nothing           'Liste des champs d'une commande pour une requête de contrainte d'intégrité.
        Dim sNomRequete As String = ""              'Nom de la requête à traiter.
        Dim sNomClasseSelection As String = ""      'Nom de la classe de sélection à traiter.
        Dim sParametres As String = ""              'Paramètres de la requête à traiter.
        Dim sNomClassesRelation As String = ""      'Nom des classes en relation à traiter.
        Dim sTypeSelection As String = "Enlever"    'Type de sélection à effectuer dans la requête à traiter.
        Dim sReferenceSpatiale As String = ""       'Nom de la référence spatiale et de la précision utilisées pour le traitement.
        Dim iNoTypeSel As Integer = 4               'Numéro du champs pour le type de sélection.

        Try
            'Requête invalide par défaut
            goRequete = Nothing

            'Conserver la liste des Layers traités
            gqListeLayers = qListeLayers
            'Conserver la Géodatabase
            gpFeatureWorkspace = pFeatureWorkspace
            'Conserver la référence spatiale
            gpSpatialReference = pSpatialReference
            'Conserver la requête attributive
            gpQueryFilter = pQueryFilter
            'Conserver le nom de la classe de découpage
            gsNomClasseDecoupage = sNomClasseDecoupage
            'Conserver le Layer de découpage
            gpFeatureLayerDecoupage = pFeatureLayerDecoupage
            'Conserver l'élément de découpage
            gpFeatureDecoupage = pFeatureDecoupage

            'Lire les champs de la commande
            Using rdr As New StringReader(sCommande)
                Using parser As New TextFieldParser(rdr)
                    parser.TextFieldType = FileIO.FieldType.Delimited
                    parser.Delimiters = New String() {" "}
                    parser.HasFieldsEnclosedInQuotes = True
                    sChamps = parser.ReadFields()
                End Using
            End Using

            'Définir le nom de la requête
            sNomRequete = sChamps(0)

            'Définir le nom de la classe de sélection
            sNomClasseSelection = sChamps(1)

            'Définir les paramètres de la requête
            sParametres = sChamps(2)

            'Vérifier si des classes en relation sont présentes
            If sChamps.Length > 3 Then
                'Vérifier si la valeur correspond au type de sélection Enlever
                If sChamps(3).ToUpper = "ENLEVER" Or sChamps(3).ToUpper = "CONSERVER" Then
                    'Définir le numéro du champs du type de sélection à effectuer dans la requête à traiter.
                    iNoTypeSel = 3

                    'Si des classes en relations sont présentes
                Else
                    'Définir la collection des FeatureLayer en relation
                    sNomClassesRelation = sChamps(3)
                End If

                'Vérifier si le type de sélection est spécifié
                If sChamps.Length > iNoTypeSel Then
                    'Vérifier si le type de sélection est Conserver
                    If sChamps(iNoTypeSel).ToUpper = "CONSERVER" Then
                        'Définir le type de sélection à effectuer dans la requête à traiter.
                        sTypeSelection = "Conserver"
                    End If

                    'Vérifier si la référence spatiale et la précision sont spécifiées
                    If sChamps.Length > iNoTypeSel + 1 Then
                        'Définir la la référence spatiale et la précision
                        sReferenceSpatiale = sChamps(iNoTypeSel + 1)
                    End If
                End If
            End If

            'Créer la requête, définir le FeatureLayer de sélection et les paramètres
            CreerRequeteCommande(sNomRequete, sNomClasseSelection, sParametres, sNomClassesRelation, sTypeSelection, sReferenceSpatiale)

            'Vérifier si l'élément de découpage est absent
            If pFeatureDecoupage Is Nothing Then
                'Définir le nom de la classe de découpage à traiter
                Me.NomClasseDecoupage = sNomClasseDecoupage
            Else
                'Définir le FeatureLayer de découpage à traiter
                Me.FeatureLayerDecoupage = pFeatureLayerDecoupage
                'Définir l'élément de découpage à traiter
                Me.FeatureDecoupage = pFeatureDecoupage
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            sChamps = Nothing
            sNomRequete = Nothing
            sNomClasseSelection = Nothing
            sParametres = Nothing
            sNomClassesRelation = Nothing
            sTypeSelection = Nothing
            sReferenceSpatiale = Nothing
            iNoTypeSel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la requête selon son nom et ses paramêtres.
    ''' 
    '''<param name="sNom"> Nom de la requête à traiter.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(Optional ByVal sNom As String = "ValeurAttribut",
                   Optional ByRef pFeatureLayerSelection As IFeatureLayer = Nothing,
                   Optional ByRef sParametres As String = C_PARAMETRES_DEFAUT)
        Try
            'Créer la requête
            goRequete = CreerRequete(sNom, pFeatureLayerSelection, sParametres)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de vider la mémoire.
    '''</summary>
    Protected Overrides Sub Finalize()
        'Vider la mémoire
        goRequete = Nothing
        gpGeometryBagErreur = Nothing
        gsNomFichierJournal = Nothing
        gsNomClasseDecoupage = Nothing
        gpFeatureWorkspace = Nothing
        gpSpatialReference = Nothing
        gpQueryFilter = Nothing
        gqListeLayers = Nothing
        gpFeatureDecoupage = Nothing
        gpFeatureLayerDecoupage = Nothing
        gpTrackCancel = Nothing
        giNombreElementsDebut = Nothing
        giNombreElementsFin = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
        'Finaliser
        MyBase.Finalize()
    End Sub
#End Region

#Region "Propriétés"
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
    ''' Propriété qui permet de définir et retourner le TrackCancel pour l'exécution de la requête.
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
    ''' Propriété qui permet de retourner le nom de la requête à traiter.
    '''</summary>
    ''' 
    Public ReadOnly Property Nom() As String
        Get
            'Vérifier si la requête est valide
            If goRequete IsNot Nothing Then
                Nom = goRequete.Nom
            Else
                Nom = ""
            End If
        End Get
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
            'Vérifier si le nom de la classe est présent
            If gsNomClasseDecoupage.Length > 0 Then
                'Définir le FeatureLayer de découpage
                gpFeatureLayerDecoupage = CreerFeatureLayer(gsNomClasseDecoupage, gpFeatureWorkspace, gpSpatialReference, gpQueryFilter)
                'Définir le FeatureLayer de découpage
                goRequete.FeatureLayerDecoupage = gpFeatureLayerDecoupage
                'Définir la limite du découpage
                goRequete.DefinirLimiteLayerDecoupage(gpFeatureLayerDecoupage)
            End If
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'élément de découpage à traiter.
    '''</summary>
    ''' 
    Public Property FeatureDecoupage() As IFeature
        Get
            'Retourner l'élément de découpage
            FeatureDecoupage = gpFeatureDecoupage
        End Get
        Set(ByVal value As IFeature)
            'Définir l'élément de découpage
            gpFeatureDecoupage = value
            'Définir l'élément de découpage
            goRequete.FeatureDecoupage = gpFeatureDecoupage
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le FeatureLayer contenant les éléments de découpage à traiter.
    '''</summary>
    ''' 
    Public Property FeatureLayerDecoupage() As IFeatureLayer
        Get
            'Retourner le FeatureLayer de découpage
            FeatureLayerDecoupage = gpFeatureLayerDecoupage
        End Get
        Set(ByVal value As IFeatureLayer)
            'Définir le FeatureLayer de découpage
            gpFeatureLayerDecoupage = value
            'Définir le FeatureLayer de découpage dans la requête même
            goRequete.FeatureLayerDecoupage = gpFeatureLayerDecoupage
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre d'éléments sélectionnés au début du traitement.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreElementsDebut() As Long
        Get
            NombreElementsDebut = giNombreElementsDebut
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le nombre d'éléments sélectionnés à la fin du traitement.
    '''</summary>
    ''' 
    Public ReadOnly Property NombreElementsFin() As Long
        Get
            NombreElementsFin = giNombreElementsFin
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner la requête à traiter.
    '''</summary>
    ''' 
    Public ReadOnly Property Requete() As intRequete
        Get
            Requete = goRequete
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
    ''' Fonction qui permet de retourner la liste des requêtes possibles.
    '''</summary>
    ''' 
    Public Function Requetes() As Collection
        Try
            'Définir la liste des paramètres par défaut
            Requetes = New Collection

            'Ajouter une requête possible
            Requetes.Add("AdjacenceGeometrie")

            'Ajouter une requête possible
            Requetes.Add("AjustementDecoupage")

            'Ajouter une requête possible
            Requetes.Add("AngleGeometrie")

            'Ajouter une requête possible
            Requetes.Add("ComparerElement")

            'Ajouter une requête possible
            Requetes.Add("CreerGeometrie")

            'Ajouter une requête possible
            Requetes.Add("DecoupageElement")

            'Ajouter une requête possible
            Requetes.Add("DistanceGeometrie")

            'Ajouter une requête possible
            Requetes.Add("EquidistanceCourbe")

            'Ajouter une requête possible
            Requetes.Add("GeometrieValide")

            'Ajouter une requête possible
            Requetes.Add("LongueurGeometrie")

            'Ajouter une requête possible
            Requetes.Add("NombreGeometrie")

            'Ajouter une requête possible
            Requetes.Add("NombreIntersection")

            'Ajouter une requête possible
            Requetes.Add("ProximiteGeometrie")

            'Ajouter une requête possible
            Requetes.Add("RelationSpatiale")

            'Ajouter une requête possible
            Requetes.Add("SensEcoulement")

            'Ajouter une requête possible
            Requetes.Add("SuperficieGeometrie")

            'Ajouter une requête possible
            Requetes.Add("ValeurAttribut")

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'exécuter la requête et retourner le résultat obtenu.
    '''</summary>
    '''
    '''<returns>IGeometry contenant les géométries d'erreurs.</returns>
    ''' 
    Public Function Executer() As IGeometry
        'Déclaration des variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface contenant les éléments sélectionnés.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant les géométrie d'erreurs.
        Dim dDateDebut As DateTime = Nothing            'Contient la date de début du traitement.
        Dim oProcess As Process = Nothing               'Objet contenant l'information sur la mémoire.

        'Par défaut, un GeometryBag vide est retourné
        Executer = New GeometryBag

        Try
            'Initialiser le traitement
            giNombreElementsDebut = 0
            giNombreElementsFin = 0

            'Afficher la date de début
            dDateDebut = System.DateTime.Now
            'Afficher la requête sous forme de commande
            EcrireMessage(goRequete.Commande)
            'Afficher la date de début
            EcrireMessage("  Début: " & dDateDebut.ToString())

            'Interface pour extraire les éléments sélectionnés
            pFeatureSel = CType(goRequete.FeatureLayerSelection, IFeatureSelection)
            'Conserver le nombre d'éléments à traiter
            giNombreElementsDebut = pFeatureSel.SelectionSet.Count

            'Si au moins un élément est sélectionné
            If giNombreElementsDebut > 0 Then
                'Sélectionner les éléments en erreurs et retourner le GeometryBag d'erreurs
                pGeometry = goRequete.Selectionner(gpTrackCancel, goRequete.EnleverSelection)
            Else
                'Retourner un GeometryBag vide
                pGeometry = New GeometryBag
                'Définir la référence spatiale
                pGeometry.SpatialReference = goRequete.SpatialReference
            End If

            'Définir l'objet pour extraire la mémoire utilisée
            oProcess = Process.GetCurrentProcess()
            'Conserver le nombre d'éléments sélectionnés
            giNombreElementsFin = pFeatureSel.SelectionSet.Count
            'Afficher le nombre d'erreur
            EcrireMessage(" -" & giNombreElementsDebut.ToString & " éléments traités, " & giNombreElementsFin.ToString & " éléments sélectionnés, (WorkingSet: " & (oProcess.WorkingSet64 / 1024).ToString("N0") & " K)")
            'Afficher la date de fin
            EcrireMessage("  Fin: " & System.DateTime.Now.ToString() & " (Temps exécution: " & System.DateTime.Now.Subtract(dDateDebut).ToString & ")")

            'Conserver les géométries d'erreurs
            gpGeometryBagErreur = CType(pGeometry, IGeometryBag)

            'Indique le succès du traitement
            Executer = pGeometry

        Catch ex As CancelException
            'Écrire l'erreur
            EcrireMessage("[clsRequete.Executer] " & ex.Message & vbCrLf)
            'Retourner l'exception
            Throw
        Catch ex As Exception
            'Retourner l'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pGeometry = Nothing
            dDateDebut = Nothing
            oProcess = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de créer la requête selon tous les paramétres d'une ligne de commande. 
    '''</summary>
    ''' 
    '''<param name="sNomRequete"> Nom de la requête à traiter.</param>
    '''<param name="sNomClasseSelection"> Nom dela classe de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    '''<param name="sNomClassesRelation"> Nom des classes en relation à traiter.</param>
    '''<param name="sTypeSelection"> Type de sélection utilisé pour le traitement.</param>
    '''<param name="sReferenceSpatiale"> Référence spatiale et précision utilisées pour le traitement.</param>
    '''
    Private Sub CreerRequeteCommande(ByVal sNomRequete As String, ByVal sNomClasseSelection As String, ByVal sParametres As String, _
                                     ByVal sNomClassesRelation As String, ByVal sTypeSelection As String, ByVal sReferenceSpatiale As String)
        'Déclarer les variables de travail
        Dim pFeatureLayerSelection As IFeatureLayer = Nothing       'FeatureLayer de sélection.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pSpatialReference As ISpatialReference = Nothing        'Interface contenant la référence spatiale.
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
        Dim sParams() As String             'Liste des paramètres de la réference spatiale.
        Dim sCode As String = ""            'Contient le code de la référence spatiale.
        Dim sPrecision As String = ""       'Contient la précision de la référence spatiale.

        Try
            'Vérifier si la référence spatiale est présente
            If sReferenceSpatiale.Length > 0 Then
                'Définir les paramètres de la référence spatiale
                sParams = sReferenceSpatiale.Split(CChar(","))

                'Définir le code de la référence spatiale
                sCode = sParams(0).Split(CChar(":"))(0)

                'Vérifier si la précision est spécifiée
                If sParams.Length > 1 Then
                    'Définir la précision de la référence spatiale
                    sPrecision = sParams(1)
                End If

                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment

                'Créer la référence spatiale
                pSpatialReference = pSpatialRefFact.CreateSpatialReference(CInt(sCode))

                'Initialiser la résolution
                pSpatialRefRes = CType(pSpatialReference, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                'Interface pour définir la tolérance XY
                pSpatialRefTol = CType(pSpatialReference, ISpatialReferenceTolerance)
                pSpatialRefTol.XYTolerance = pSpatialRefRes.XYResolution(True) * 2

                'Vérifier si la précison est présente
                If sPrecision.Length > 0 Then
                    'Vérifier la référence spatiale est géographique
                    pSpatialRefTol.XYTolerance = CDbl(sPrecision)
                End If

                'Conserver la référence spatiale
                gpSpatialReference = pSpatialReference
            End If

            'Interface pour définir la tolérance XY
            pSpatialRefTol = CType(gpSpatialReference, ISpatialReferenceTolerance)
            'Valider la précision
            If Not pSpatialRefTol.XYToleranceValid = esriSRToleranceEnum.esriSRToleranceOK Then
                'Retourner une erreur
                Throw New Exception("ERREUR : La tolérance XY de référence spatiale est invalide!")
            End If
            'Définition de la référence spatiale et de la précision
            sReferenceSpatiale = gpSpatialReference.FactoryCode.ToString & ":" & gpSpatialReference.Name & "," & pSpatialRefTol.XYTolerance.ToString

            'Vérifier si le FeatureLayer de sélection est présent dans la collection des Layers conservés
            If gqListeLayers.Contains(sNomClasseSelection) Then
                'Définir le FeatureLayer de sélection présent dans la liste des Layers traités
                pFeatureLayerSelection = CType(gqListeLayers.Item(sNomClasseSelection), IFeatureLayer)

                'Si le FeatureLayer est absent de la liste des Layers
            Else
                'Définir le FeatureLayer de sélection
                pFeatureLayerSelection = CreerFeatureLayer(sNomClasseSelection, gpFeatureWorkspace, gpSpatialReference, gpQueryFilter)

                'Ajouter le FeatureLayer dans la collection des Layers traités
                gqListeLayers.Add(pFeatureLayerSelection, sNomClasseSelection)
            End If

            'Vérifier les requêtes
            goRequete = CreerRequete(sNomRequete, pFeatureLayerSelection, sParametres)

            'Définir le type de sélection
            goRequete.EnleverSelection = (sTypeSelection.ToUpper = "ENLEVER")

            'Vérifier si des classes en relation sont présentes
            If sNomClassesRelation.Length > 0 Then
                'Définir la collection des FeatureLayer en relation
                goRequete.FeatureLayersRelation = CreerFeatureLayerCollection(sNomClassesRelation, gpFeatureWorkspace, gqListeLayers, gpSpatialReference, gpQueryFilter)

                'Définir la collection des RasterLayer en relation
                goRequete.RasterLayersRelation = CreerRasterLayerCollection(sNomClassesRelation, gpSpatialReference, gpQueryFilter)

                'Définir la commande complète
                goRequete.Commande = sNomRequete & " """ & sNomClasseSelection & """ """ & sParametres & """ """ & sNomClassesRelation & """ " & sTypeSelection & " """ & sReferenceSpatiale & """"
            Else
                'Définir la commande complète
                goRequete.Commande = sNomRequete & " """ & sNomClasseSelection & """ """ & sParametres & """ " & sTypeSelection & " """ & sReferenceSpatiale & """"
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayerSelection = Nothing
            pSpatialRefFact = Nothing
            pSpatialReference = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
            sParams = Nothing
            sCode = Nothing
            sPrecision = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de créer la requête selon son nom de requête, un FeatureLayer de sélection et ses paramêtres. 
    '''</summary>
    ''' 
    '''<param name="sNom"> Nom de la requête à traiter.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    '''
    Private Function CreerRequete(Optional ByVal sNom As String = "ValeurAttribut",
                                 Optional ByRef pFeatureLayerSelection As IFeatureLayer = Nothing,
                                 Optional ByRef sParametres As String = C_PARAMETRES_DEFAUT) As intRequete
        'La requête est invalide par défaut
        CreerRequete = Nothing

        Try
            'Vérifier les requêtes
            If sNom = "AdjacenceGeometrie" Then
                'Définir la requête
                CreerRequete = New clsAdjacenceGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "AjustementDecoupage" Then
                'Définir la requête
                CreerRequete = New clsAjustementDecoupage(pFeatureLayerSelection, sParametres, gpFeatureDecoupage)

            ElseIf sNom = "AngleGeometrie" Then
                'Définir la requête
                CreerRequete = New clsAngleGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "ComparerElement" Then
                'Définir la requête
                CreerRequete = New clsComparerElement(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "CreerGeometrie" Then
                'Définir la requête
                CreerRequete = New clsCreerGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "DecoupageElement" Then
                'Définir la requête
                CreerRequete = New clsDecoupageElement(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "DistanceGeometrie" Then
                'Définir la requête
                CreerRequete = New clsDistanceGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "EquidistanceCourbe" Then
                'Définir la requête
                CreerRequete = New clsEquidistanceCourbe(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "GeometrieValide" Then
                'Définir la requête
                CreerRequete = New clsGeometrieValide(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "LongueurGeometrie" Then
                'Définir la requête
                CreerRequete = New clsLongueurGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "NombreGeometrie" Then
                'Définir la requête
                CreerRequete = New clsNombreGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "NombreIntersection" Then
                'Définir la requête
                CreerRequete = New clsNombreIntersection(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "ProximiteGeometrie" Then
                'Définir la requête
                CreerRequete = New clsProximiteGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "RelationSpatiale" Then
                'Définir la requête
                CreerRequete = New clsRelationSpatiale(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "SuperficieGeometrie" Then
                'Définir la requête
                CreerRequete = New clsSuperficieGeometrie(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "SensEcoulement" Then
                'Définir la requête
                CreerRequete = New clsSensEcoulement(pFeatureLayerSelection, sParametres)

            ElseIf sNom = "ValeurAttribut" Then
                'Définir la requête
                CreerRequete = New clsValeurAttribut(pFeatureLayerSelection, sParametres)
            Else
                'Retourner l'erreur
                Err.Raise(-1, , "ERREUR : Le nom de la requête '" & sNom & "' est invalide!")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
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
    Private Function CreerFeatureLayer(ByVal sNomClasse As String, ByVal pFeatureWorkspace As IFeatureWorkspace,
                                       Optional ByVal pSpatialReference As ISpatialReference = Nothing,
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
                'Vérifier si le nom de la classe correspond au motclé [DECOUPAGE]
                If sNomClasse = "[DECOUPAGE]" Then
                    'Vérifier si le Layer de découpage est absent
                    If gpFeatureLayerDecoupage Is Nothing Then
                        'définir le nom de la classe de découpage
                        sNomClasse = gsNomClasseDecoupage

                        'Vérifier si le Layer de découpage est présent
                    Else
                        'Définir le nom de la classe
                        sNomClasse = gpFeatureLayerDecoupage.FeatureClass.AliasName
                    End If
                End If

                'Vérifier si le nom de la classe contient le nom de la Géodatabase
                If sNomClasse.Contains("\") Then
                    'Définir le nom de la classe
                    sNomLayer = System.IO.Path.GetFileName(sNomClasse)

                    'Définir le nom de la géodatabase sans celui du nom de la classe
                    sNomGeodatabase = sNomClasse.Replace("\" & sNomLayer, "")
                    sNomGeodatabase = sNomGeodatabase.Replace(sNomLayer, "")

                    'Définir la Géodatabase
                    pFeatureWorkspace = CType(DefinirGeodatabase(sNomGeodatabase), IFeatureWorkspace)

                    'Si le nom de la classe ne contient pas le nom de la Géodatabase
                Else
                    'Extraire le nom du Layer et sa condition si elle est présente
                    sLayer = Split(sNomClasse, ":")
                    sNomLayer = sLayer(0)
                    'Vérifier si le nom de la classe possède une condition attributive
                    If sLayer.Length > 1 Then
                        'Définir le filtre
                        sFiltre = sLayer(1)
                    End If
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
                'Ajouter la condition attributive seulement
                pFeatureLayerDef.DefinitionExpression = sFiltre

                '    'Vérifier si une requête additionnelle est nécessaire
                '    If pQueryFilter IsNot Nothing Then
                '        'Ajouter la condition attributive et celle additionnelle
                '        pFeatureLayerDef.DefinitionExpression = pQueryFilter.WhereClause & " AND (" & sFiltre & ")"

                '        'Si on ajoute seulement la contifion attributive
                '    Else
                '        'Ajouter la condition attributive seulement
                '        pFeatureLayerDef.DefinitionExpression = sFiltre
                '    End If

                '    'Vérifier si une requête additionnelle est nécessaire
                'ElseIf pQueryFilter IsNot Nothing Then
                '    'Ajouter la condition attributive additionnelle seulement
                '    pFeatureLayerDef.DefinitionExpression = pQueryFilter.WhereClause
            End If

            'Interface pour sélectionner les élément
            pFeatureSel = CType(CreerFeatureLayer, IFeatureSelection)

            'Sélectionner les éléments
            pFeatureSel.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)

            'Définir la référence spatiale du FeatureLayer au besoin
            If pSpatialReference IsNot Nothing Then CreerFeatureLayer.SpatialReference = pSpatialReference

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
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
            sNomTable = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de créer une collection de FeatureLayer en relation contenant ou non une condition attributive.
    '''</summary>
    '''
    '''<param name="sNomClasses">Nom des classes en relation contenant ou non une condition attributive (ex:BDG_COURBE_NIVEAU_1:ELEVATION=100).</param>
    '''<param name="pFeatureWorkspace">Interface utilisé pour ouvrir une classe dans la Géodatabase.</param>
    '''<param name="qListeLayers">Collection des Layers traités et conservés.</param>
    '''<param name="pSpatialReference">Interface contenant la référence spatiale de traitement.</param>
    '''<param name="pQueryFilter">Interface utilisé pour sélectionner seulement les éléments du découpage.</param> 
    ''' 
    Private Function CreerFeatureLayerCollection(ByVal sNomClasses As String, _
                                                 ByVal pFeatureWorkspace As IFeatureWorkspace, _
                                                 ByVal qListeLayers As Collection, _
                                                 Optional ByVal pSpatialReference As ISpatialReference = Nothing, _
                                                 Optional ByVal pQueryFilter As IQueryFilter = Nothing) As Collection
        'Déclaration des variables de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant la classe de la Géodatabase.
        Dim pLayerFile As ILayerFile = Nothing          'Interface qui permet de lire un FeatureLayer sur disque.
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface contenant les éléments sélectionnés

        'Définir la valeur par défaut, une collection vide
        CreerFeatureLayerCollection = New Collection

        Try
            'Définir la liste des Layers en relation
            For Each sNomClasse In Split(sNomClasses, ";")
                'Vérifier si le nom n'est pas une image en format .tif
                If Not sNomClasse.Contains(".tif") Then
                    'Vérifier si le nom est un FeatureLayer sur disque
                    If sNomClasse.Contains(".lyr") Then
                        'Interface pour ouvrir le FeatureLayer de découpage
                        pLayerFile = New LayerFile

                        'Vérifier si le FeatureLayer est valide
                        If pLayerFile.IsLayerFile(sNomClasse) Then
                            'Ouvrir le FeatureLayer de découpage sur disque
                            pLayerFile.Open(sNomClasse)

                            'Définir le FeatureLayer de découpage
                            pFeatureLayer = CType(pLayerFile.Layer, IFeatureLayer)

                            'Vérifier si la classe est valide
                            If pFeatureLayer.FeatureClass Is Nothing Then
                                'Retourner l'erreur
                                Err.Raise(-1, , "ERREUR : Classe invalide !")
                            End If

                            'Interface pour vérifier si des éléments sont sélectionnés
                            pFeatureSel = CType(pFeatureLayer, IFeatureSelection)

                            'Vérifier si des éléments sont sélectionnés
                            If pFeatureSel.SelectionSet.Count = 0 Then
                                'Sélectionner tous les éléments
                                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                            End If

                            'Si le Layer de découpage est invalide
                        Else
                            'Retourner l'erreur
                            Err.Raise(-1, , "ERREUR : FeatureLayer sur disque invalide : " & sNomClasse)
                        End If

                        'si le nom est une classe
                    Else
                        'Vérifier si le FeatureLayer de sélection est présent dans la liste des Layers
                        If qListeLayers.Contains(sNomClasse) Then
                            'Définir le FeatureLayer présent dans la liste
                            pFeatureLayer = CType(qListeLayers.Item(sNomClasse), IFeatureLayer)

                            'si le FeatureLayer de sélection est absent de la liste des Layers
                        Else
                            'Définir le FeatureLayer en relation
                            pFeatureLayer = CreerFeatureLayer(sNomClasse, pFeatureWorkspace, pSpatialReference, pQueryFilter)
                        End If
                    End If

                    'Ajouter le FeatureLayer dans la collection
                    CreerFeatureLayerCollection.Add(pFeatureLayer, pFeatureLayer.Name)
                End If
            Next

            'Si aucun FeatureLayers en relation on vide la mémoire
            If CreerFeatureLayerCollection.Count = 0 Then CreerFeatureLayerCollection = Nothing

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
            pLayerFile = Nothing
            pFeatureSel = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de créer une collection de RasterLayer en relation.
    '''</summary>
    '''
    '''<param name="sNomImages">Nom complet des images ou du catalogue d'images en relation (ex:BDG_MNE_CATALOG).</param>
    '''<param name="pSpatialReference">Interface contenant la référence spatiale de traitement.</param>
    '''<param name="pQueryFilter">Interface utilisé pour sélectionner seulement les éléments du découpage.</param> 
    ''' 
    Private Function CreerRasterLayerCollection(ByVal sNomImages As String, _
                                                Optional ByVal pSpatialReference As ISpatialReference = Nothing, _
                                                Optional ByVal pQueryFilter As IQueryFilter = Nothing) As Collection
        'Déclaration des variables de travail
        Dim pRasterLayer As IRasterLayer = Nothing        'Interface contenant un MNE en relation.

        'Définir la valeur par défaut, une collection vide
        CreerRasterLayerCollection = New Collection

        Try
            'Définir la liste des Layers en relation
            For Each sNomImage In Split(sNomImages, ";")
                'Vérifier si le nom est une image en format .tif
                If sNomImage.Contains(".tif") Then
                    'Définir le FeatureLayer en relation
                    pRasterLayer = CreerRasterLayer(sNomImage, pSpatialReference)

                    'Ajouter le FeatureLayer dans la collection
                    CreerRasterLayerCollection.Add(pRasterLayer, pRasterLayer.Name)
                End If
            Next

            'Si aucun RasterLayers en relation on vide la mémoire
            If CreerRasterLayerCollection.Count = 0 Then CreerRasterLayerCollection = Nothing

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pRasterLayer = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de créer un RasterLayer à partir d'un nom complet d'image contenant un MNE.
    '''</summary>
    '''
    '''<param name="sNomImage">Nom complet de l'image contenant un MNE.</param>
    '''<param name="pSpatialReference">Interface contenant la référence spatiale de traitement.</param>
    ''' 
    Private Function CreerRasterLayer(ByVal sNomImage As String, _
                                      Optional ByVal pSpatialReference As ISpatialReference = Nothing) As RasterLayer
        'Déclaration des variables de travail

        'Définir la valeur par défaut, une Layer vide
        CreerRasterLayer = New RasterLayer

        Try
            'Ajouter le Raster dans le Layer
            CreerRasterLayer.CreateFromFilePath(sNomImage)

            'Définir le nom du RasterLayer
            CreerRasterLayer.Name = sNomImage & "_" & System.DateTime.Now.ToString

            'Mettre non visible par défaut
            CreerRasterLayer.Visible = False

            'Vérifier si la référence spatiale est spécifié
            If pSpatialReference IsNot Nothing Then
                'Définir la référence spatiale du RasterLayer
                CreerRasterLayer.SpatialReference = pSpatialReference
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Function

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
    ''' Routine qui permet d'écrire le message d'exécution dans un RichTextBox, un fichier journal et/ou dans la console.
    '''</summary>
    ''' 
    '''<param name="sMessage"> Message à écrire dans un RichTextBox, un fichier journal et/ou dans la console.</param>
    '''
    Private Sub EcrireMessage(ByVal sMessage As String)
        Try
            'Conserver le message dans l'information du traitement
            gsInformation = gsInformation & sMessage & vbCrLf

            'Écrire dans la console
            Console.WriteLine(sMessage)

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

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

#End Region
End Class
