Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.GeoDatabaseUI

'**
'Nom de la composante : CancelException.vb
'
'''<summary>
''' Classe contenant une exception permettant d'annulé l'exécution d'un traitement.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 24 mars 2016
'''</remarks>
''' 
Public Class CancelException
    Inherits Exception

    Public Sub New()
    End Sub

    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(message As String, inner As Exception)
        MyBase.New(message, inner)
    End Sub
End Class

'**
'Nom de la composante : clsValeurAttribut.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont la valeur d’attribut respecte ou non l’expression régulière spécifiée.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 15 avril 2015
'''</remarks>
''' 
Public Class clsValeurAttribut
    Implements intRequete

    'Déclarer les variables globales
    '''<summary>Nom de l'attribut à traiter.</summary>
    Protected Const C_NOM_ATTRIBUT_DEFAUT As String = "DATASET_NAME"
    '''<summary>Expression régulière à traiter.</summary>
    Protected Const C_EXPRESSION_REGULIERE_DEFAUT As String = "^(\d\d\d[A-P](0[1-9]|1[0-6]))$"

    '''<summary>Interface contenant l'application.</summary>
    Protected gpApplication As IApplication = Nothing
    '''<summary>Interface contenant le document ArcMap à traiter.</summary>
    Protected gpMxDocument As IMxDocument = Nothing
    '''<summary>Interface contenant la Map à traiter.</summary>
    Protected gpMap As IMap = Nothing
    '''<summary>Interface contenant l'envelope des éléments sélectionnés à traiter.</summary>
    Protected gpEnvelope As IEnvelope = Nothing
    '''<summary>Interface contenant une référence spatiale.</summary>
    Protected gpSpatialReference As ISpatialReference = Nothing
    '''<summary>Interface contenant le FeatureLayer de sélection à traiter.</summary>
    Protected gpFeatureLayerSelection As IFeatureLayer = Nothing
    '''<summary>Interface contenant la collection des FeatureLayers en relation.</summary>
    Protected gpFeatureLayersRelation As Collection = Nothing
    '''<summary>Interface contenant la collection des RasterLayers en relation.</summary>
    Protected gpRasterLayersRelation As Collection = Nothing
    '''<summary>Interface contenant le FeatureLayer de la classe de découpage.</summary>
    Protected gpFeatureLayerDecoupage As IFeatureLayer = Nothing
    '''<summary>Interface contenant l'élément de découpage de la classe de découpage.</summary>
    Protected gpFeatureDecoupage As IFeature = Nothing
    '''<summary>Interface contenant le polygone de découpage.</summary>
    Protected gpPolygoneDecoupage As IPolygon = Nothing
    '''<summary>Interface contenant les limites des polygones de découpage.</summary>
    Protected gpLimiteDecoupage As IPolyline = New Polyline
    '''<summary>Nom de l'attribut à traiter.</summary>
    Protected gsNomAttribut As String = ""
    '''<summary>Expression régulière à traiter.</summary>
    Protected gsExpression As String = ""
    '''<summary>Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True</summary>
    Protected gbEnleverSelection As Boolean = True
    '''<summary>Contient un message pour indiquer si la contrainte est valide ou qu'est-ce qui est invalide.</summary>
    Protected gsMessage As String = ""
    '''<summary>Contient le numéro de l'index de l'attribut.</summary>
    Protected giAttribut As Integer = -1
    '''<summary>Précision des données.</summary>
    Protected gdPrecision As Double = 0.001
    '''<summary>Interface ESRI contenant les éléments en mémoire.</summary>
    Protected gpFeatureClassErreur As IFeatureClass = Nothing
    '''<summary>Interface ESRI utilisé pour ajouter des éléments dans la classe en mémoire.</summary>
    Protected gpFeatureCursorErreur As IFeatureCursor = Nothing
    ''' <summary>Indique si on doit créer la classe d'erreurs.</summary>
    Protected gbCreerClasseErreur As Boolean = True
    ''' <summary>Indique si on doit afficher la table d'erreurs.</summary>
    Protected gbAfficherTableErreur As Boolean = True
    ''' <summary>Contient la requête sous forme de commande.</summary>
    Protected gsCommande As String = ""

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = C_NOM_ATTRIBUT_DEFAUT
            Expression = C_EXPRESSION_REGULIERE_DEFAUT

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    ''' 
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sNomAttribut"> Nom de l'attribut à traiter.</param>
    '''<param name="sExpression"> Expression régulière à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap,
                   ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String,
                   ByVal sExpression As String)
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    ''' 
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pFeatureLayerSelection As IFeatureLayer,
                   Optional ByVal sParametres As String = C_NOM_ATTRIBUT_DEFAUT + " " + C_EXPRESSION_REGULIERE_DEFAUT)
        'Déclarer les variables de travail

        Try
            'Définir les valeurs par défaut
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    ''' 
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap,
                   ByRef pFeatureLayerSelection As IFeatureLayer,
                   Optional ByVal sParametres As String = C_NOM_ATTRIBUT_DEFAUT + " " + C_EXPRESSION_REGULIERE_DEFAUT)
        'Déclarer les variables de travail

        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de vider la mémoire des objets de la classe.
    '''</summary>
    '''
    Protected Overridable Overloads Sub finalize() Implements intRequete.finalize
        'Vider la mémoire
        gpApplication = Nothing
        gpMxDocument = Nothing
        gpMap = Nothing
        gpFeatureLayerSelection = Nothing
        gpFeatureLayersRelation = Nothing
        gpEnvelope = Nothing
        gpSpatialReference = Nothing
        gpRasterLayersRelation = Nothing
        gpFeatureLayerDecoupage = Nothing
        gpFeatureDecoupage = Nothing
        gpPolygoneDecoupage = Nothing
        gpLimiteDecoupage = Nothing
        gsNomAttribut = Nothing
        gsExpression = Nothing
        gbEnleverSelection = Nothing
        gsMessage = Nothing
        giAttribut = Nothing
        gdPrecision = Nothing
        gpFeatureClassErreur = Nothing
        gpFeatureCursorErreur = Nothing
        gbCreerClasseErreur = Nothing
        gbAfficherTableErreur = Nothing
        gsCommande = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
        'Finaliser
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de retourner le nom de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Overridable Overloads ReadOnly Property Nom() As String Implements intRequete.Nom
        Get
            Nom = "ValeurAttribut"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'application.
    '''</summary>
    ''' 
    Public Property Application() As IApplication Implements intRequete.Application
        Get
            Application = gpApplication
        End Get
        Set(ByVal value As IApplication)
            gpApplication = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le document ArcMap à traiter.
    '''</summary>
    ''' 
    Public Property MxDocument() As IMxDocument Implements intRequete.MxDocument
        Get
            MxDocument = gpMxDocument
        End Get
        Set(ByVal value As IMxDocument)
            gpMxDocument = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la Map à traiter.
    '''</summary>
    ''' 
    Public Property Map() As IMap Implements intRequete.Map
        Get
            Map = gpMap
        End Get
        Set(ByVal value As IMap)
            gpMap = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la Map à traiter.
    '''</summary>
    ''' 
    Public Property Envelope() As IEnvelope Implements intRequete.Envelope
        Get
            Envelope = gpEnvelope
        End Get
        Set(ByVal value As IEnvelope)
            gpEnvelope = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le FeatureLayer contenant les éléments sélectionnés à traiter.
    '''</summary>
    ''' 
    Public Property FeatureLayerSelection() As IFeatureLayer Implements intRequete.FeatureLayerSelection
        Get
            'Retourner 
            FeatureLayerSelection = gpFeatureLayerSelection
        End Get
        Set(ByVal value As IFeatureLayer)
            'Définir le FeatureLayer de sélection
            gpFeatureLayerSelection = value
            'Vérifier si le FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Déclarer les variables de travail
                Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.
                'Définir la référence spatiale
                gpSpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                'Interface contenant la tolrance de précision de la Référence spatiale
                pSRTolerance = CType(gpSpatialReference, ISpatialReferenceTolerance)
                'Définir la précision selon la référence spatiale
                gdPrecision = pSRTolerance.XYTolerance
                'Vider la mémoire
                pSRTolerance = Nothing
                'Conserver l'enveloppe de la classe de sélection
                gpEnvelope = gpFeatureLayerSelection.AreaOfInterest
            End If
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la référence spatiale utilisée pour le traitement.
    '''</summary>
    ''' 
    Public Property SpatialReference() As ISpatialReference Implements intRequete.SpatialReference
        Get
            'Retourner la référence spatiale
            SpatialReference = gpSpatialReference
        End Get
        Set(ByVal value As ISpatialReference)
            'Définir la référence spatiale
            gpSpatialReference = value
            'Vérifier si la référence spatiale est valide
            If gpSpatialReference IsNot Nothing Then
                'Déclarer les variables de travail
                Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.
                'Interface contenant la tolrance de précision de la Référence spatiale
                pSRTolerance = CType(gpSpatialReference, ISpatialReferenceTolerance)
                'Définir la précision selon la référence spatiale
                gdPrecision = pSRTolerance.XYTolerance
                'Vider la mémoire
                pSRTolerance = Nothing
            End If
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la collection des FeatureLayers en relation.
    '''</summary>
    ''' 
    Public Property FeatureLayersRelation() As Collection Implements intRequete.FeatureLayersRelation
        Get
            FeatureLayersRelation = gpFeatureLayersRelation
        End Get
        Set(ByVal value As Collection)
            gpFeatureLayersRelation = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la collection des RasterLayers en relation.
    '''</summary>
    ''' 
    Public Property RasterLayersRelation() As Collection Implements intRequete.RasterLayersRelation
        Get
            RasterLayersRelation = gpRasterLayersRelation
        End Get
        Set(ByVal value As Collection)
            gpRasterLayersRelation = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la FeatureClass des erreurs trouvées.
    '''</summary>
    ''' 
    Public ReadOnly Property FeatureClassErreur() As IFeatureClass Implements intRequete.FeatureClassErreur
        Get
            FeatureClassErreur = gpFeatureClassErreur
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le FeatureLayer contenant les géométries d'erreurs.
    '''</summary>
    ''' 
    Public ReadOnly Property FeatureLayerErreur() As IFeatureLayer Implements intRequete.FeatureLayerErreur
        Get
            'Par défaut le FeatureLayer est invalide
            FeatureLayerErreur = Nothing

            'Vérifier si la FeatureClass d'erreur est valide
            If gpFeatureClassErreur IsNot Nothing Then
                'Créer un nouveau FeatureLayer
                FeatureLayerErreur = New FeatureLayer
                'Définir le nom du FeatureLayer selon le nom et la date
                FeatureLayerErreur.Name = gpFeatureClassErreur.AliasName
                'Rendre visible le FeatureLayer
                FeatureLayerErreur.Visible = True
                'Définir la Featureclass dans le FeatureLayer
                FeatureLayerErreur.FeatureClass = gpFeatureClassErreur
            End If
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le FeatureLayer de découpage.
    '''</summary>
    ''' 
    Public Property FeatureLayerDecoupage() As IFeatureLayer Implements intRequete.FeatureLayerDecoupage
        Get
            FeatureLayerDecoupage = gpFeatureLayerDecoupage
        End Get
        Set(ByVal value As IFeatureLayer)
            'Initialiser les paramètres de l'élément de découpage
            FeatureDecoupage = Nothing

            'Vérifier si le Layer est valide
            If value IsNot Nothing Then
                'Vérifier si la FeatureClass est valide
                If value.FeatureClass IsNot Nothing Then
                    'Vérifier si l'élément de découpage est de type polygone
                    If value.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Conserver le Layer de découpage
                        gpFeatureLayerDecoupage = value
                    End If
                End If
            End If
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'élément de découpage.
    '''</summary>
    ''' 
    Public Property FeatureDecoupage() As IFeature Implements intRequete.FeatureDecoupage
        Get
            FeatureDecoupage = gpFeatureDecoupage
        End Get
        Set(ByVal value As IFeature)
            'Définir les valeurs par défaut
            gpFeatureDecoupage = Nothing
            PolygoneDecoupage = Nothing

            'Vérifier si l'élément est valide
            If value IsNot Nothing Then
                'Vérifier si l'élément de découpage est de type polygone
                If value.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Conserver l'élément de découpage
                    gpFeatureDecoupage = value

                    'Définir le polygone de découpage
                    PolygoneDecoupage = CType(gpFeatureDecoupage.ShapeCopy, IPolygon)
                End If
            End If
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le polygone de découpage.
    '''</summary>
    ''' 
    Public Property PolygoneDecoupage() As IPolygon Implements intRequete.PolygoneDecoupage
        Get
            PolygoneDecoupage = gpPolygoneDecoupage
        End Get
        Set(ByVal value As IPolygon)
            'Définir les valeurs par défaut
            gpPolygoneDecoupage = Nothing
            gpLimiteDecoupage = Nothing
            'Vérifier si la valeur est valide
            If value IsNot Nothing Then
                'Conserver le polygone de découpage
                gpPolygoneDecoupage = value

                'Interface pour extraire la limite du polygone de découpage
                Dim pTopoOp As ITopologicalOperator2 = CType(gpPolygoneDecoupage, ITopologicalOperator2)
                'Définir la limite du polygone de découpage
                gpLimiteDecoupage = CType(pTopoOp.Boundary, IPolyline)
                'Vider la mémoire
                pTopoOp = Nothing
            End If
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la limite du polygone de découpage.
    '''</summary>
    ''' 
    Public Property LimiteDecoupage() As IPolyline Implements intRequete.LimiteDecoupage
        Get
            'Vérifier si la limite est invalide
            LimiteDecoupage = gpLimiteDecoupage
        End Get
        Set(ByVal value As IPolyline)
            gpLimiteDecoupage = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom de l'attribut à traiter.
    '''</summary>
    ''' 
    Public Property NomAttribut() As String Implements intRequete.NomAttribut
        Get
            NomAttribut = gsNomAttribut
        End Get
        Set(ByVal value As String)
            gsNomAttribut = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'expression régulière à traiter.
    '''</summary>
    ''' 
    Public Overridable Overloads Property Expression() As String Implements intRequete.Expression
        Get
            Expression = gsExpression
        End Get
        Set(ByVal value As String)
            gsExpression = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la ligne de paramètre à traiter.
    '''</summary>
    ''' 
    Public Overridable Overloads Property Parametres() As String Implements intRequete.Parametres
        Get
            Parametres = gsNomAttribut + " " + gsExpression
        End Get
        Set(ByVal value As String)
            gsNomAttribut = value.Split(CChar(" "))(0)
            gsExpression = value.Split(CChar(" "))(1)
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la tolérance de précision.
    '''</summary>
    ''' 
    Public Overridable Overloads Property Precision() As Double Implements intRequete.Precision
        Get
            Precision = gdPrecision
        End Get
        Set(ByVal value As Double)
            gdPrecision = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner si on doit créer la classe d'erreurs.
    '''</summary>
    ''' 
    Public Property CreerClasseErreur() As Boolean Implements intRequete.CreerClasseErreur
        Get
            CreerClasseErreur = gbCreerClasseErreur
        End Get
        Set(ByVal value As Boolean)
            gbCreerClasseErreur = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner si on doit afficher la table d'erreurs.
    '''</summary>
    ''' 
    Public Property AfficherTableErreur() As Boolean Implements intRequete.AfficherTableErreur
        Get
            AfficherTableErreur = gbAfficherTableErreur
        End Get
        Set(ByVal value As Boolean)
            gbAfficherTableErreur = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner si on doit enlever les éléments de la sélection.
    '''</summary>
    ''' 
    Public Property EnleverSelection() As Boolean Implements intRequete.EnleverSelection
        Get
            EnleverSelection = gbEnleverSelection
        End Get
        Set(ByVal value As Boolean)
            gbEnleverSelection = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner le message de contrainte valide ou invalide.
    '''</summary>
    ''' 
    Public ReadOnly Property Message() As String Implements intRequete.Message
        Get
            Message = gsMessage
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de retourner la requête à effectuer sous forme de commande texte.
    '''</summary>
    ''' 
    Public Property Commande() As String Implements intRequete.Commande
        Get
            'Déclarer la variable de travail
            Dim pFeatuLayerDef As IFeatureLayerDefinition = Nothing     'Interface pour extraire la requête attributive.
            Dim pFeatureLayer As IFeatureLayer = Nothing                'Interface contenant La classe en relation avec sa requête attrbutive.
            Dim pDataset As IDataset = Nothing                          'Interface pour vérifier si la Géodatabase de la classe est en mémoire.
            Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
            Dim sNomClasse As String = ""                               'Contient le nom de la classe.

            'La commande est vide par défaut
            Commande = gsCommande

            Try
                'Vérifier si la commande initiale a été spécifiée
                If gsCommande.Length = 0 Then
                    'Définir le nom de la requête
                    Commande = Nom

                    'Interface pour vérifier si le type de Géodatabase est en mémoire
                    pDataset = CType(gpFeatureLayerSelection, IDataset)
                    'Vérifier si la classe est contenue dans une Géodatabase en mémoire
                    If pDataset.Workspace.PathName.StartsWith("{") Then
                        'Définir le nom de la classe
                        sNomClasse = gpFeatureLayerSelection.Name
                    Else
                        'Définir le nom de la classe
                        sNomClasse = gpFeatureLayerSelection.FeatureClass.AliasName.ToUpper
                        'Vérifier si un nom d'usager est présent dans le nom de la classe
                        If sNomClasse.Contains(".") Then
                            'Enlever le nom de la classe
                            sNomClasse = sNomClasse.Split(CChar("."))(1)
                        End If
                    End If

                    'Interface pour extraire la requête attributive
                    pFeatuLayerDef = CType(gpFeatureLayerSelection, IFeatureLayerDefinition)
                    'Vérifier si aucune requête attributive n'est présente
                    If pFeatuLayerDef.DefinitionExpression.Length = 0 Then
                        'Définir la classe de sélection
                        Commande = Commande & " " & sNomClasse
                        'Si une requête attributive est présente
                    Else
                        'Définir la classe de sélection avec sa requête attributive
                        Commande = Commande & " """ & sNomClasse & ":" & pFeatuLayerDef.DefinitionExpression & """"
                    End If

                    'Définir les paramêtres de la requête
                    Commande = Commande & " """ & Parametres & """"

                    'Vérifier si des Classes en relation sont présentes
                    If FeatureLayersRelation IsNot Nothing Then
                        'Vérifier si au moins une classe en relation est présente
                        If FeatureLayersRelation.Count > 0 Then
                            'Définir les classes en relation
                            For i = 1 To FeatureLayersRelation.Count
                                'Définir le FeatureLayer en relation
                                pFeatureLayer = CType(FeatureLayersRelation.Item(i), IFeatureLayer)

                                'Interface pour vérifier si le type de Géodatabase est en mémoire
                                pDataset = CType(pFeatureLayer, IDataset)
                                'Vérifier si la classe est contenue dans une Géodatabase en mémoire
                                If pDataset.Workspace.PathName.StartsWith("{") Then
                                    'Définir le nom de la classe
                                    sNomClasse = pFeatureLayer.Name.ToUpper
                                Else
                                    'Définir le nom de la classe
                                    sNomClasse = pFeatureLayer.FeatureClass.AliasName.ToUpper
                                    'Vérifier si un nom d'usager est présent dans le nom de la classe
                                    If sNomClasse.Contains(".") Then
                                        'Enlever le nom de la classe
                                        sNomClasse = sNomClasse.Split(CChar("."))(1)
                                    End If
                                End If

                                'Interface pour extraire la requête attributive
                                pFeatuLayerDef = CType(pFeatureLayer, IFeatureLayerDefinition)
                                'Vérifier si c'est le premier FeatureLayer en relation
                                If i = 1 Then
                                    'Vérifier si une requête attributive est présente
                                    If pFeatuLayerDef.DefinitionExpression.Length = 0 Then
                                        'Définir la classe en relation avec sa requête attributive au besoin
                                        Commande = Commande & " """ & sNomClasse
                                    Else
                                        'Définir la classe en relation avec sa requête attributive au besoin
                                        Commande = Commande & " """ & sNomClasse & ":" & pFeatuLayerDef.DefinitionExpression
                                    End If
                                    'Lorsque plusieurs classes en relation
                                Else
                                    'Vérifier si une requête attributive est présente
                                    If pFeatuLayerDef.DefinitionExpression.Length = 0 Then
                                        'Définir la classe en relation avec sa requête attributive au besoin
                                        Commande = Commande & ";" & sNomClasse
                                    Else
                                        'Définir la classe en relation avec sa requête attributive au besoin
                                        Commande = Commande & ";" & sNomClasse & ":" & pFeatuLayerDef.DefinitionExpression
                                    End If
                                End If
                            Next

                            'Fermer la liste des classes en relation
                            Commande = Commande & """"
                        End If
                    End If

                    'Indiquer le type de sélection
                    If gbEnleverSelection Then
                        'Enlever ce qui est trouvé dans la sélection
                        Commande = Commande & " Enlever"
                    Else
                        'Conserver ce qui est trouvé dans la sélection
                        Commande = Commande & " Conserver"
                    End If

                    'Interface pour extraire la précision
                    pSpatialRefTol = CType(gpSpatialReference, ISpatialReferenceTolerance)
                    'Indiquer la référence spatiale et la précision
                    Commande = Commande & " """ & gpSpatialReference.FactoryCode.ToString & ":" & gpSpatialReference.Name & "," & pSpatialRefTol.XYTolerance.ToString & """"
                End If

            Catch ex As Exception
                'Message d'erreur
                Throw
            Finally
                'Vider la mémoire
                pFeatuLayerDef = Nothing
                pFeatureLayer = Nothing
                pDataset = Nothing
                pSpatialRefTol = Nothing
                sNomClasse = Nothing
            End Try
        End Get
        Set(ByVal value As String)
            gsCommande = value
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    ''' Fonction qui permet de retourner la liste des paramètres possibles.
    '''</summary>
    ''' 
    Public Overridable Overloads Function ListeParametres() As Collection Implements intRequete.ListeParametres
        Try
            'Définir la liste des paramètres par défaut
            ListeParametres = New Collection

            'Définir le premier paramètre
            ListeParametres.Add(gsNomAttribut + " " + gsExpression)

            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Traiter tous les attribut
                For i As Integer = 0 To gpFeatureLayerSelection.FeatureClass.Fields.FieldCount - 1
                    'Vérifier si l'attribut est éditable et n'est pas celui de la géométrie
                    If gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Editable _
                    And gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Name <> gpFeatureLayerSelection.FeatureClass.ShapeFieldName Then
                        'Vérifier si l'attribut est différent de celui par défaut
                        If gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Name <> gsNomAttribut Then
                            'Définir le nom de l'attribut avec l'expression par défaut
                            ListeParametres.Add(gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Name + " *")
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si les traitement à effectuer est valide.
    ''' 
    '''<return>Boolean qui indique si les traitement à effectuer est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function EstValide() As Boolean Implements intRequete.EstValide
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            EstValide = False
            gsMessage = "La contrainte est invalide."

            'Vérifier si le FeatureLayer est valide
            If Me.FeatureLayerValide() Then
                'Vérifier si la FeatureClass est valide
                If Me.FeatureClassValide() Then
                    'Vérifier si l'attribut est valide
                    If Me.AttributValide() Then
                        'Vérifier si l'expression est valide
                        If Me.ExpressionValide() Then
                            'Vérifier si les FeatureLayers en relation sont valides
                            If Me.FeatureLayersRelationValide() Then
                                'Vérifier si les RasterLayers en relation sont valides
                                If Me.RasterLayersRelationValide() Then
                                    'La contrainte est valide
                                    EstValide = True
                                    gsMessage = "La contrainte est valide."
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si le FeatureLayer est valide.
    ''' 
    '''<return>Boolean qui indique si le FeatureLayer est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function FeatureLayerValide() As Boolean Implements intRequete.FeatureLayerValide
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            FeatureLayerValide = False
            gsMessage = "ERREUR : Le FeatureLayer est invalide."

            'Vérifier si le FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'La contrainte est valide
                FeatureLayerValide = True
                gsMessage = "La contrainte est valide."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si la FeatureClass est valide.
    ''' 
    '''<return>Boolean qui indique si la FeatureClass est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function FeatureClassValide() As Boolean Implements intRequete.FeatureClassValide
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            FeatureClassValide = False
            gsMessage = "ERREUR : La FeatureClass est invalide."

            'Vérifier si la FeatureClass est valide
            If gpFeatureLayerSelection.FeatureClass IsNot Nothing Then
                'La contrainte est valide
                FeatureClassValide = True
                gsMessage = "La contrainte est valide."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si les FeatureLayers en relation sont valides.
    ''' 
    '''<return>Boolean qui indique si les FeatureLayers en relation sont valides.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function FeatureLayersRelationValide() As Boolean Implements intRequete.FeatureLayersRelationValide
        'Déclarer les variable de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant un FeatureLayer.

        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            FeatureLayersRelationValide = False
            gsMessage = "ERREUR : FeatureLayer en relation est invalide."

            'Vérifier si les FeatureLayers en relation sont absent
            If gpFeatureLayersRelation Is Nothing Then
                'La contrainte est valide
                FeatureLayersRelationValide = True
                gsMessage = "La contrainte est valide."

                'Vérifier la présence des FeatureLayers en relation
            ElseIf gpFeatureLayersRelation.Count > 0 Then
                'Traiter tous les FeatureLayers en relation
                For Each pFeatureLayer In gpFeatureLayersRelation
                    'Vérifier si le FeatureLayer est valide
                    If pFeatureLayer IsNot Nothing Then
                        'Vérifier si la FeatureClass est invalide
                        If pFeatureLayer.FeatureClass Is Nothing Then
                            'sortir de la fonction
                            Exit Function
                        End If
                    Else
                        'Sortir de la fonction
                        Exit Function
                    End If
                Next

                'La contrainte est valide
                FeatureLayersRelationValide = True
                gsMessage = "La contrainte est valide."

            Else
                'Message d'erreur
                gsMessage = "ERREUR : Aucun FeatureLayer en relation n'est sélectionné."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si les RasterLayers en relation sont valides.
    ''' 
    '''<return>Boolean qui indique si les RasterLayers en relation sont valides.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function RasterLayersRelationValide() As Boolean Implements intRequete.RasterLayersRelationValide
        'Déclarer les variable de travail
        Dim pRasterLayer As IRasterLayer = Nothing    'Interface contenant un RasterLayer.

        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            RasterLayersRelationValide = False
            gsMessage = "ERREUR : RasterLayer en relation est invalide."

            'Vérifier si les RasterLayers en relation sont absent
            If gpRasterLayersRelation Is Nothing Then
                'La contrainte est valide
                RasterLayersRelationValide = True
                gsMessage = "La contrainte est valide."

                'Vérifier la présence des RasterLayers en relation
            ElseIf gpRasterLayersRelation.Count > 0 Then
                'Traiter tous les RasterLayers en relation
                For Each pRasterLayer In gpRasterLayersRelation
                    'Vérifier si le RasterLayer est valide
                    If pRasterLayer IsNot Nothing Then
                        'Vérifier si la RasterClass est invalide
                        If pRasterLayer.Raster Is Nothing Then
                            'sortir de la fonction
                            Exit Function
                        End If
                    Else
                        'Sortir de la fonction
                        Exit Function
                    End If
                Next

                'La contrainte est valide
                RasterLayersRelationValide = True
                gsMessage = "La contrainte est valide."

            Else
                'Message d'erreur
                gsMessage = "ERREUR : Aucun RasterLayer en relation n'est sélectionné."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pRasterLayer = Nothing
        End Try
    End Function


    '''<summary>
    ''' Routine qui permet d'indiquer si l'attribut est valide.
    ''' 
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function AttributValide() As Boolean Implements intRequete.AttributValide
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            AttributValide = False
            gsMessage = "ERREUR : L'attribut est invalide : " & gsNomAttribut

            'Définir la position de l'attribut dansla classe
            giAttribut = gpFeatureLayerSelection.FeatureClass.Fields.FindField(gsNomAttribut)

            'Vérifier si l'attribut est valide
            If giAttribut >= 0 Then
                'La contrainte est valide
                AttributValide = True
                gsMessage = "La contrainte est valide."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si l'expression régulière est valide.
    ''' 
    '''<return>Boolean qui indique si l'expression régulière est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function ExpressionValide() As Boolean Implements intRequete.ExpressionValide
        'Déclarer les variables de travail
        Dim oRegEx As Regex       'Objet utilisé pour valider l'expression régulière.

        Try
            'La contrainte est valide
            ExpressionValide = False
            gsMessage = "ERREUR : L'expression régulière est invalide."

            'Vérifier si l'expression régulière est présente
            If Len(gsExpression) > 0 Then
                Try
                    'Valider l'expression régulière
                    oRegEx = New Regex(gsExpression)
                    'La contrainte est valide
                    ExpressionValide = True
                    gsMessage = "La contrainte est valide."

                Catch ex As Exception
                    gsMessage = ex.Message
                End Try
            Else
                gsMessage = "ERREUR : L'expression régulière est absente."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner la MapTopology contenant les éléments en relations dont la topologie peut être créée et extraite.
    ''' 
    ''' La topologie doit être créée ou extraite en mode édition selon une tolérance de précision (Cluster).
    ''' Les Nodes et les Edges sont accessible via l'interface ITopologyGraph4 du IMapTopology.cache 
    ''' 
    '''<param name="pEnvelope">Interface contenant l'enveloppe pour construire le MapTopology.</param>
    '''<param name="dTolerance">Tolerance de proximité.</param>
    ''' 
    '''<return> IMapTopology contenant la topologie des éléments en relations.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function RelationsMapTopology(ByVal pEnvelope As IEnvelope, Optional ByVal dTolerance As Double = 0) As IMapTopology Implements intRequete.RelationsMapTopology
        'Déclarer les variables de travail
        Dim pTopologyUID As UID = New UIDClass()            'Interface contenant l'identifiant de l'extension de la topologie
        Dim pTopologyExt As ITopologyExtension = Nothing    'Interface contenant l'extension de la topologie
        Dim pMapTopology As IMapTopology = Nothing          'Interface contenant la topologie des éléments en relation.

        'Définir la valeur de retour par défaut
        RelationsMapTopology = Nothing

        Try
            'Définir l'extension de topologie
            pTopologyUID.Value = "esriEditorExt.TopologyExtension"
            pTopologyExt = CType(Application.FindExtensionByCLSID(pTopologyUID), ITopologyExtension)

            'Vérifier si la topology est valide
            If pTopologyExt.CurrentTopology IsNot Nothing Then
                'Vérifier si la topologie courante est un MapTopology
                If TypeOf pTopologyExt.CurrentTopology Is IMapTopology Then
                    'Définir l'interface pour créer la topologie
                    pMapTopology = pTopologyExt.MapTopology
                    pMapTopology = TryCast(pTopologyExt.CurrentTopology, IMapTopology)

                    'Définir la tolérance de proximité
                    If dTolerance > 0 Then pMapTopology.ClusterTolerance = dTolerance

                    Try
                        'Construire la topologie selen la fenêtre active et conserver la sélection
                        pMapTopology.Cache.Build(pEnvelope, True)
                    Catch ex As Exception
                        'Sortir de la fonction
                        Exit Function
                    End Try

                    'Vérifier si la topologie est présente
                    If pMapTopology.Cache.Edges.Count > 0 Or pMapTopology.Cache.Nodes.Count > 0 Then
                        'Retourner la topologie créée
                        RelationsMapTopology = pMapTopology
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pTopologyUID = Nothing
            pTopologyExt = Nothing
            pMapTopology = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner une copie de la MAP contenant les FeatureLayers en relation.
    ''' 
    ''' 
    '''<return> Copie de la Map contenant les FeatureLayers en relation.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function RelationsMap() As IMap Implements intRequete.RelationsMap
        'Déclarer les varibles de travail
        Dim pMap As IMap = Nothing                      'Interface ESRI contenant des Layers.
        Dim pMapFrame As IMapFrame = Nothing            'Interface ESRI contenant un MapFrame (Fenêtre de données).
        Dim pElement As IElement = Nothing              'Interface ESRI contenant un objet dans le PageLayout.
        Dim pEnvelope As IEnvelope = Nothing            'Interface ESRI utilisée pour positionner le MapFrame.
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface ESRI contenant une classe de données.
        Dim pObjectCopy As IObjectCopy = Nothing        'Interface utilisé pour copier un FeatureLayer
        Dim pObj As Object = Nothing                    'Résultat de l'objet copié

        'Définir la valeur de retour par défaut
        RelationsMap = Nothing

        Try
            ''Créer et nommé une nouvelle Fenêtre
            'pMap = New Map
            'pMap.Name = "Relations"

            ''Créer un Frame associé à la Fenêtre
            'pMapFrame = New MapFrame
            'pMapFrame.Map = pMap

            ''Définir la référence spatiale
            'pMap.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            ''Positionner le Frame
            'pElement = CType(pMapFrame, IElement)
            'pEnvelope = New Envelope
            'pEnvelope.PutCoords(0, 0, 5, 5)
            'pElement.Geometry = pEnvelope

            ''Ajouter tous les FeatureLayer à la fenêtre de données
            ''Traiter tous les Layers
            'For i = 0 To gpMap.LayerCount - 1
            '    'Vérifier si le Layer est un FeatureLayer
            '    If TypeOf (gpMap.Layer(i)) Is IFeatureLayer Then
            '        'Créer un nouvel objet vide
            '        pObjectCopy = New ObjectCopyClass()
            '        'Faire une copie du FeatureLayer dans le nouvel objet
            '        pObj = pObjectCopy.Copy(CObj(gpMap.Layer(i)))
            '        'Interface pour le FeatureLayer
            '        pFeatureLayer = CType(pObj, IFeatureLayer)
            '        If pFeatureLayer.Visible Then
            '            'Ajouter un FeatureLayer à la fenêtre de données
            '            pMap.AddLayer(pFeatureLayer)
            '        End If
            '    End If
            'Next

            ''Retourner la fenêtre de données vide
            'RelationsMap = pMap

            'Faire une copie de la Map
            pObjectCopy = New ObjectCopyClass()
            pObj = pObjectCopy.Copy(CObj(gpMap))
            RelationsMap = CType(pObj, IMap)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pMap = Nothing
            pMapFrame = Nothing
            pElement = Nothing
            pEnvelope = Nothing
            pFeatureLayer = Nothing
            pObjectCopy = Nothing
            pObj = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de définir les limites de découpage à partir des éléments sélectionner d'un Layer de découpage. 
    ''' 
    '''<param name="pFeatureLayerDecoupage">Interface contenant les élements de découpage sélectionnés.</param>
    ''' 
    '''<return>IPolyline contenant la les limites de découpage des éléments sélectionner dans le Layer de découpage, Vide si aucun sélectionné.</return>
    ''' 
    '''</summary>
    '''
    Public Function DefinirLimiteLayerDecoupage(ByRef pFeatureLayerDecoupage As IFeatureLayer) As IPolyline Implements intRequete.DefinirLimiteLayerDecoupage
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour vérifier la présence des éléments sélectionnés.
        Dim pSelectionSet As ISelectionSet = Nothing    'Interface contenant les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface utilisé pour extraire les éléments sélectionnés.
        Dim pCursor As ICursor = Nothing                'Interface utilisé pour extraire les éléments sélectionnés.
        Dim pFeature As IFeature = Nothing              'Interface contenant un élément sélectionné.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface utilisé pour pour calculer l'intersection des limites commune.
        Dim pGeoDataset As IGeoDataset = Nothing        'Interface contenant la référence spatiale.

        'Définir la limite du découpage vide par défaut 
        DefinirLimiteLayerDecoupage = New Polyline

        Try
            'Initialiser le FeatureLayer de découpage, l'élément, le polygon et la limite
            FeatureLayerDecoupage = Nothing

            'Vérifier si le FeatureLayer est spécifié
            If pFeatureLayerDecoupage IsNot Nothing Then
                'Vérifier si le Layer de découpage est de type polygone
                If pFeatureLayerDecoupage.FeatureClass IsNot Nothing Then
                    'Vérifier si le Layer de découpage est de type polygone
                    If pFeatureLayerDecoupage.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Conserver le FeatureLayer de découpage
                        gpFeatureLayerDecoupage = pFeatureLayerDecoupage

                        'Interface pour définir la référence spatiale
                        pGeoDataset = CType(pFeatureLayerDecoupage, IGeoDataset)

                        'Définir la référence spatiale
                        DefinirLimiteLayerDecoupage.SpatialReference = pGeoDataset.SpatialReference

                        'Interface pour extraire les éléments sélectionnés
                        pFeatureSel = CType(pFeatureLayerDecoupage, IFeatureSelection)

                        'Vérifier si des découpages sont sélectionnés
                        If pFeatureSel.SelectionSet.Count > 0 Then
                            'Interface contenant les éléments sélectionnés
                            pSelectionSet = pFeatureSel.SelectionSet

                            'Extraire les éléments sélectionnés
                            pSelectionSet.Search(Nothing, True, pCursor)

                            'Interface pour extraire chaque élément
                            pFeatureCursor = CType(pCursor, IFeatureCursor)

                            'Extraire le premier élément
                            pFeature = pFeatureCursor.NextFeature

                            'Vérifier si plusieurs découpages sont sélectionnés
                            If pSelectionSet.Count > 1 Then
                                'Interface pour ajouter des limites dans le Polyline de retour
                                pGeomColl = CType(DefinirLimiteLayerDecoupage, IGeometryCollection)

                                'Traiter tous les éléments
                                Do Until pFeature Is Nothing
                                    'Interface pour extraire la limite du polygone de découpage
                                    pTopoOp = CType(pFeature.ShapeCopy, ITopologicalOperator2)

                                    'Ajouter la limite du polygone de découpage dans la polyline de retour
                                    pGeomColl.AddGeometryCollection(CType(pTopoOp.Boundary, IGeometryCollection))

                                    'Extraire le premier élément
                                    pFeature = pFeatureCursor.NextFeature
                                Loop

                                'Si un seul découpage est sélectionné
                            Else
                                'Conserver l'élément de découpage, son polygone et sa limite
                                gpFeatureDecoupage = pFeature

                                'Conserver le polygone de découpage
                                gpPolygoneDecoupage = CType(pFeature.ShapeCopy, IPolygon)

                                'Interface pour extraire la limite du polygone de découpage
                                pTopoOp = CType(gpPolygoneDecoupage, ITopologicalOperator2)

                                'Définir la polyline de retour
                                DefinirLimiteLayerDecoupage = CType(pTopoOp.Boundary, IPolyline)
                            End If

                            'Interface pour simplifier la limite des polygones de découpage
                            pTopoOp = CType(DefinirLimiteLayerDecoupage, ITopologicalOperator2)

                            'Simplifier la limite des polygones de découpage
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()

                            'Conserver la limite de découpage
                            gpLimiteDecoupage = DefinirLimiteLayerDecoupage
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureCursor = Nothing
            pCursor = Nothing
            pFeature = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
            pGeoDataset = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de définir une table contenue dans une géodatabase.
    '''</summary>
    ''' 
    '''<param name="sNomTable"> Nom complet de la table incluant le nom de la géodatabase.</param>
    '''<param name="sNomGeodatabaseDefaut"> Contient le nom de la Géodatabase Enterprise par défaut.</param>
    '''<param name="sNomProprietaireDefaut"> Contient le nom du propriétaire des tables par défaut pour les Géodatabase Enterprise.</param>
    ''' 
    '''<returns>ITable contenant la table, sinon Nothing.</returns>
    ''' 
    Public Overridable Overloads Function DefinirTable(ByVal sNomTable As String,
                                                       Optional ByVal sNomGeodatabaseDefaut As String = "database connections\BDRS_PRO_BDG_DBA.sde",
                                                       Optional ByVal sNomProprietaireDefaut As String = "BDG_DBA") As ITable Implements intRequete.DefinirTable
        'Déclaration des variables de travail
        Dim sNomTbl As String = ""                      'Contient le nom de la table.
        Dim sNomGeodatabase As String = ""              'Contient le nom de la Géodatabase.
        Dim sRepArcCatalog As String = ""               'Nom du répertoire contenant les connexions des Géodatabase .sde.
        Dim pGeodatabase As IFeatureWorkspace = Nothing 'Interface pour ouvrir une table.
        Dim pWorkspace2 As IWorkspace2 = Nothing        'Interface pour vérifier si la table existe.
        Dim pWorkspace As IWorkspace = Nothing          'Interface pour vérifier le type de Géodatabase.

        'Par défaut, aucune table n'est présente
        DefinirTable = Nothing

        Try
            'Vérifier si le nom de la table est présent
            If sNomTable.Length > 0 Then
                'Extraire le nom du répertoire contenant les connexions des Géodatabase .sde.
                sRepArcCatalog = IO.Directory.GetDirectories(Environment.GetEnvironmentVariable("APPDATA"), "ArcCatalog", IO.SearchOption.AllDirectories)(0)

                'Vérifier si le nom de la Géodatabase de la table est absent
                If Not sNomTable.Contains("\") Then
                    'Définir le nom de la Géodatabase par défaut
                    sNomTable = sNomGeodatabaseDefaut & "\" & sNomTable
                End If

                'Redéfinir le nom complet de la Géodatabase .sde
                sNomTable = sNomTable.ToLower.Replace("database connections", sRepArcCatalog)

                'Définir le nom de la table sans le nom de la Géodatabase
                sNomTbl = System.IO.Path.GetFileName(sNomTable)

                'Définir le nom de la géodatabase sans celui du nom de la table
                sNomGeodatabase = sNomTable.Replace("\" & sNomTbl, "")
                sNomGeodatabase = sNomGeodatabase.Replace(sNomTbl, "")

                'Ouvrir la géodatabase de la table
                pGeodatabase = CType(DefinirGeodatabase(sNomGeodatabase), IFeatureWorkspace)

                'Vérifier si la Géodatabase est valide
                If pGeodatabase IsNot Nothing Then
                    'Interface pour vérifier le type de Géodatabase.
                    pWorkspace = CType(pGeodatabase, IWorkspace)
                    'Vérifier si la Géodatabase est de type "Enterprise" 
                    If pWorkspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                        'Vérifier si le nom de la table contient le nom du propriétaire
                        If Not sNomTbl.Contains(".") Then
                            'Définir le nom de la table avec le nom du propriétaire
                            sNomTbl = sNomProprietaireDefaut & "." & sNomTbl
                        End If
                    End If

                    'Interface pour vérifier si la table existe
                    pWorkspace2 = CType(pGeodatabase, IWorkspace2)
                    'Vérifier si la table existe
                    If pWorkspace2.NameExists(esriDatasetType.esriDTTable, sNomTbl) Then
                        'Ouvrir la table
                        DefinirTable = pGeodatabase.OpenTable(sNomTbl)
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            sNomTable = Nothing
            sNomGeodatabase = Nothing
            sRepArcCatalog = Nothing
            pGeodatabase = Nothing
            pWorkspace2 = Nothing
            pWorkspace = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet d'ouvrir et retourner la Geodatabase à partir d'un nom de Géodatabase.
    '''</summary>
    '''
    '''<param name="sNomGeodatabase"> Nom de la géodatabase à ouvrir.</param>
    ''' 
    '''<returns>IWorkspace contenant la géodatabase.</returns>
    ''' 
    Public Overridable Overloads Function DefinirGeodatabase(ByVal sNomGeodatabase As String) As IWorkspace Implements intRequete.DefinirGeodatabase
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la valeur d'attribut respecte ou non l'expression régulière spécifiée.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les enveloppes des géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function Selectionner(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry Implements intRequete.Selectionner
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface pour projeter.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.

        Try
            'Sortir si la contrainte est invalide
            If Me.EstValide() = False Then Err.Raise(1, , Me.Message)

            'Définir la géométrie par défaut
            Selectionner = New GeometryBag
            Selectionner.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour ajouter l'enveloppe des géométries sélectionnées
            pGeomColl = CType(Selectionner, IGeometryCollection)

            'Conserver le type de sélection à effectuer
            gbEnleverSelection = bEnleverSelection

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer
            pSelectionSet = pFeatureSel.SelectionSet

            'Vérifier si des éléments sont sélectionnés
            If pSelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                pSelectionSet = pFeatureSel.SelectionSet
            End If

            'Enlever la sélection
            pFeatureSel.Clear()

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Interface pour ajouter l'enveloppe des géométries des éléments sélectionnés
            pGeomColl = CType(Selectionner, IGeometryCollection)

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Interface pour projeter
                pGeometry = pFeature.Shape
                pGeometry.Project(Selectionner.SpatialReference)

                'Valider la valeur d'attribut selon l'expression régulière
                'oMatch = oRegEx.Match(Me.ValeurAttribut(pFeature))
                oMatch = oRegEx.Match(CStr(pFeature.Value(giAttribut)))

                'Vérifier si on doit sélectionner l'élément
                If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                    'Ajouter l'élément dans la sélection
                    pFeatureSel.Add(pFeature)
                    'Ajouter l'enveloppe de l'élément sélectionné
                    pGeomColl.AddGeometry(pGeometry)
                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Valeur=" & Me.ValeurAttribut(pFeature) & "/" & gsExpression, pGeometry)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Vérifier si la classe d'erreur est présente
            If gpFeatureCursorErreur IsNot Nothing Then
                'Conserver toutes les modifications
                gpFeatureCursorErreur.Flush()
                'Release the update cursor to remove the lock on the input data.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(gpFeatureCursorErreur)
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            oMatch = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeature = Nothing
            pFeatureCursor = Nothing
            pGeomColl = Nothing
            pGeometry = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner la valeur de l'attribut sous forme de texte.
    ''' 
    '''<param name="pFeature">Interface contenant l'élément traité.</param>
    ''' 
    '''<return>Texte contenant la valeur de l'attribut.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function ValeurAttribut(ByRef pFeature As IFeature) As String Implements intRequete.ValeurAttribut
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            ValeurAttribut = ""

            'Vérifier si la valeur est Nulle
            If IsDBNull(pFeature.Value(giAttribut)) Then Exit Function

            'Retourner la valeur de l'attribut
            ValeurAttribut = CStr(pFeature.Value(giAttribut))

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Fonction qui permet d'indiquer si la classe, la définition et la sélection sont les mêmes.
    '''</summary>
    '''
    '''<param name="pFeatureLayerA">Interface contenant contenant un FeatureLayer à traiter.</param>
    '''<param name="pFeatureLayerB">Interface contenant contenant un FeatureLayer en relation.</param>
    '''<param name="bMemeClasse">Indique si la classe est la même.</param>
    '''<param name="bMemeDefinition">Indique si la définition est la même.</param>
    '''<param name="bMemeSelection">Indique si la définition est la même.</param>
    ''' 
    '''<remarks>Si la classe est différente, alors la définition et la sélection sont automatiquement différentes.</remarks>
    ''' 
    Protected Sub MemeClasseMemeDefinition(ByVal pFeatureLayerA As IFeatureLayer, ByVal pFeatureLayerB As IFeatureLayer, _
                                           ByRef bMemeClasse As Boolean, ByRef bMemeDefinition As Boolean, ByRef bMemeSelection As Boolean)
        'Déclarer les variables de travail
        Dim pDatasetA As IDataset = Nothing                     'Interface utilisé pour extraire la nom complet de la classe A.
        Dim pDatasetB As IDataset = Nothing                     'Interface utilisé pour extraire la nom complet de la classe B.
        Dim pFeatureDefA As IFeatureLayerDefinition = Nothing   'Interface contenant la définition des éléments à extraire du Layer A.
        Dim pFeatureDefB As IFeatureLayerDefinition = Nothing   'Interface contenant la définition des éléments à extraire du Layer B.
        Dim pFeatureSelA As IFeatureSelection = Nothing         'Interface utilisé pour extraire les éléments sélectionnés du Layer A.
        Dim pFeatureSelB As IFeatureSelection = Nothing         'Interface utilisé pour extraire les éléments sélectionnés du Layer B.
        Dim pSelectionSet As ISelectionSet = Nothing            'Interface contenant les éléments sélectionnés.

        'Par défaut, la classe n'est pas la même
        bMemeClasse = False

        'Par défaut, la définition n'est pas la même
        bMemeDefinition = False

        'Par défaut, la sélection n'est pas la même
        bMemeSelection = False

        Try
            'Sortir si les classes sont invalides
            If pFeatureLayerA.FeatureClass Is Nothing Or pFeatureLayerA.FeatureClass Is Nothing Then Exit Sub

            'Vérifier si les FeatureLayer sont les mêmes
            If pFeatureLayerA.Equals(pFeatureLayerB) Then
                'La classe est la même
                bMemeClasse = True
                'La définition est la même
                bMemeDefinition = True
                'La sélection est la même
                bMemeSelection = True

                'Si les FeatureLayer ne sont pas les mêmes
            Else
                'Interface pour extraire le nom complet du Layer A
                pDatasetA = CType(pFeatureLayerA.FeatureClass, IDataset)

                'Interface pour extraire le nom complet du Layer B
                pDatasetB = CType(pFeatureLayerB.FeatureClass, IDataset)

                'Vérifier si les noms complets des classes sont les mêmes
                If pDatasetA.Workspace.PathName = pDatasetB.Workspace.PathName And pDatasetA.Name = pDatasetB.Name Then
                    'Indiquer que les classes sont les mêmes
                    bMemeClasse = True

                    'Interface pour extraire la définition du Layer de A
                    pFeatureDefA = CType(pFeatureLayerA, IFeatureLayerDefinition)

                    'Interface pour extraire la définition du Layer de B
                    pFeatureDefB = CType(pFeatureLayerB, IFeatureLayerDefinition)

                    'Vérifier si la définition est la même
                    If pFeatureDefA.DefinitionExpression = pFeatureDefB.DefinitionExpression Then
                        'Indiquer que la définition est la même
                        bMemeDefinition = True

                        'Interface pour extraire les élément sélectionnés du Layer A
                        pFeatureSelA = CType(pFeatureLayerA, IFeatureSelection)

                        'Interface pour extraire les élément sélectionnés du Layer B
                        pFeatureSelB = CType(pFeatureLayerB, IFeatureSelection)

                        'Interface contenant les élément sélectionnés du Layer A
                        pSelectionSet = pFeatureSelA.SelectionSet

                        'Définir les différences
                        pSelectionSet.Combine(pFeatureSelB.SelectionSet, esriSetOperation.esriSetSymDifference, pSelectionSet)

                        'Vérifier si la sélection est la même
                        If pSelectionSet.Count = 0 Then
                            'Indiquer que la définition est la même
                            bMemeSelection = True
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pDatasetA = Nothing
            pDatasetB = Nothing
            pFeatureDefA = Nothing
            pFeatureDefB = Nothing
            pFeatureSelA = Nothing
            pFeatureSelB = Nothing
            pSelectionSet = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire les éléments en relation selon un filtre spatial.
    '''</summary>
    '''
    '''<param name="pSpatialFilter">Interface contenant la relation spatiale de base.</param>
    '''<param name="pFeatureLayerRel">Interface contenant contenant un FeatureLayer en relation.</param>
    ''' 
    '''<returns>"Collection" contenant les éléments en relation.</returns>
    '''
    Protected Function ExtraireElementsRelation(ByRef pSpatialFilter As ISpatialFilter,
                                                ByRef pFeatureLayerRel As IFeatureLayer) As Collection
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface contenant les éléments sélectionnés.
        Dim pFeatureRel As IFeature = Nothing           'Interface contenant un élément en relation.
        Dim pEnumId As IEnumIDs = Nothing               'Interface qui permet d'extraire les Ids des éléments sélectionnés.
        Dim iOid As Integer = Nothing                   'Contient un Id d'un élément sélectionné.

        Try
            'Définir la valeur de retour par défaut
            ExtraireElementsRelation = New Collection

            'Sélectionner les éléments en relation
            pFeatureSel = CType(pFeatureLayerRel, IFeatureSelection)

            'Sélectionner les éléments en relation avec l'élément traité
            pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, False)

            'Interface pour extrire la liste des Ids
            pEnumId = pFeatureSel.SelectionSet.IDs

            'Trouver le premier élément en relation
            pEnumId.Reset()
            iOid = pEnumId.Next

            'Traiter tant qu'il y a des éléments en relation
            Do Until iOid = -1
                'Extraire l'élément en relation
                pFeatureRel = pFeatureLayerRel.FeatureClass.GetFeature(iOid)

                'Ajouter l'élément en relation
                ExtraireElementsRelation.Add(pFeatureRel)

                'Définir le prochain élément en relation
                iOid = pEnumId.Next
            Loop

            'Vider la sélection
            pFeatureSel.Clear()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pFeatureRel = Nothing
            pEnumId = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'initialiser la barre de progression.
    ''' 
    '''<param name="iMin">Valeur minimum.</param>
    '''<param name="iMax">Valeur maximum.</param>
    '''<param name="pTrackCancel">Interface contenant la barre de progression.</param>
    ''' 
    '''</summary>
    '''
    Protected Sub InitBarreProgression(ByVal iMin As Integer, ByVal iMax As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pStepPro As IStepProgressor = Nothing   'Interface qui permet de modifier les paramètres de la barre de progression.

        Try
            'sortir si le progressor est absent
            If pTrackCancel.Progressor Is Nothing Then Exit Sub

            'Interface pour modifier les paramètres de la barre de progression.
            pTrackCancel.Progressor = gpApplication.StatusBar.ProgressBar
            pStepPro = CType(pTrackCancel.Progressor, IStepProgressor)

            'Changer les paramètres
            pStepPro.MinRange = iMin
            pStepPro.MaxRange = iMax
            pStepPro.Position = 0
            pStepPro.Show()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pStepPro = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de lire les géométries et les OIDs des éléments d'un FeatureLayer.
    ''' 
    ''' Seuls les éléments sélectionnés sont lus.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    '''<param name="pFeatureLayer"> Interface contenant les éléments à lire.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomColl"> Interface contenant les géométries des éléments lus.</param>
    '''<param name="iOid"> Vecteur des OIDs d'éléments lus.</param>
    '''<param name="bLimite"> Indique si on veut utiliser la limite de la géométrie.</param>
    ''' 
    '''</summary>
    '''
    Protected Sub LireGeometrie(ByRef pFeatureLayer As IFeatureLayer, ByRef pTrackCancel As ITrackCancel, ByRef pGeomColl As IGeometryCollection,
                                Optional ByRef iOid() As Integer = Nothing, Optional ByVal bLimite As Boolean = False)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour lire les éléments.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour lire les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément lu.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément lu.
        Dim pTopoOp As ITopologicalOperator = Nothing       'Interface pour extraire la limite de la géométrie.

        Try
            'Créer un nouveau Bag vide
            pGeometry = New GeometryBag

            'Définir la référence spatiale
            pGeometry.SpatialReference = pFeatureLayer.AreaOfInterest.SpatialReference
            pGeometry.SnapToSpatialReference()

            'Interface pour ajouter les géométries dans le Bag
            pGeomColl = CType(pGeometry, IGeometryCollection)

            'Interface pour sélectionner les éléments
            pFeatureSel = CType(pFeatureLayer, IFeatureSelection)

            ''Vérifier si des éléments sont sélectionnés
            'If pFeatureSel.SelectionSet.Count = 0 Then
            '    'Sélectionnées tous les éléments du FeatureLayer
            '    pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            'End If

            'Interface pour extraire les éléments sélectionnés
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Augmenter le vecteur des Oid selon le nombre d'éléments
            ReDim Preserve iOid(pSelectionSet.Count)

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments du FeatureLayer
            For i = 0 To pSelectionSet.Count - 1
                'Vérifier si l'élément est présent
                If pFeature IsNot Nothing Then
                    'Définir la géométrie à traiter
                    pGeometry = pFeature.ShapeCopy

                    'Projeter la géométrie à traiter
                    pGeometry.Project(pFeatureLayer.AreaOfInterest.SpatialReference)
                    pGeometry.SnapToSpatialReference()

                    'Vérifier si on traite la limite de la géométrie
                    If bLimite Then
                        'Interface qui permet d'extraire l'intersection entre deux géométries.
                        pTopoOp = CType(pGeometry, ITopologicalOperator)
                        'Définir la limite de la géométrie
                        pGeometry = pTopoOp.Boundary()
                    End If

                    'Ajouter la géométrie dans le Bag
                    pGeomColl.AddGeometry(pGeometry)

                    'Ajouter le OID de l'élément avec sa séquence 
                    iOid(i) = pFeature.OID

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                End If

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

            'Enlever la sélection des éléments
            pFeatureSel.Clear()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de lire les géométries des éléments sélectionnés d'un FeatureLayer.
    ''' 
    '''<param name="pFeatureLayer"> Interface contenant les éléments à lire.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    ''' 
    '''</summary>
    '''
    Protected Function LireGeometrie(ByRef pFeatureLayer As IFeatureLayer, ByRef pTrackCancel As ITrackCancel) As IGeometryBag
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface utilisé pour ajouter des géométries dans le Bag.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour lire les éléments.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour lire les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément lu.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément lu.

        Try
            'Créer un nouveau Bag vide
            LireGeometrie = New GeometryBag
            pGeomColl = CType(LireGeometrie, IGeometryCollection)

            'Interface pour sélectionner les éléments
            pFeatureSel = CType(pFeatureLayer, IFeatureSelection)

            'Interface pour extraire les éléments sélectionnés
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments du FeatureLayer
            For i = 0 To pSelectionSet.Count - 1
                'Définir la géométrie à traiter
                pGeometry = pFeature.ShapeCopy

                'Projeter la géométrie à traiter
                pGeometry.Project(pFeatureLayer.AreaOfInterest.SpatialReference)
                pGeometry.SnapToSpatialReference()

                'Ajouter la géométrie dans le Bag
                pGeomColl.AddGeometry(pGeometry)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeometry = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer la Classe d'erreur et d'initialiser le curseur d'écriture des erreurs mais seulement si spécifié
    ''' via la variable globale 'm_CreerClasseErreur'.
    '''</summary>
    '''
    '''<param name="sNom">Nom de la classe à créer.</param>
    '''<param name="pFeatureLayer">Interface contenant la FeatureClass des éléments traités.</param>
    '''<param name="pEsriGeometryType">Indique le type de géométrie de la FeatureClass.</param>
    '''
    Protected Sub CreerFeatureClassErreur(ByVal sNom As String, ByVal pFeatureLayer As IFeatureLayer, Optional ByVal pEsriGeometryType As esriGeometryType = esriGeometryType.esriGeometryNull)
        'Initialisation pour l'écriture des erreurs
        gpFeatureClassErreur = Nothing
        gpFeatureCursorErreur = Nothing

        Try
            'Vérifier si la classe d'erreur doit être créée
            If Not gbCreerClasseErreur Then Exit Sub

            'Créer une classe d'erreurs en mémoire
            gpFeatureClassErreur = CreateInMemoryFeatureClass(sNom, gpFeatureLayerSelection, pEsriGeometryType)

            'Interface pour créer les erreurs
            gpFeatureCursorErreur = gpFeatureClassErreur.Insert(True)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de créer une FeatureClass en mémoire.
    '''</summary>
    '''
    '''<param name="sNom">Nom de la classe à créer.</param>
    '''<param name="pFeatureLayer">Interface contenant la FeatureClass des éléments traités.</param>
    '''<param name="pEsriGeometryType">Indique le type de géométrie de la FeatureClass.</param>
    ''' 
    '''<returns>"IFeatureClass" contenant la description et la géométrie trouvées.</returns>
    '''
    Protected Function CreateInMemoryFeatureClass(ByVal sNom As String, ByVal pFeatureLayer As IFeatureLayer, Optional ByVal pEsriGeometryType As esriGeometryType = esriGeometryType.esriGeometryNull) As IFeatureClass
        'Déclarer les variables de travail
        Dim pWorkspaceFactory As IWorkspaceFactory = Nothing    'Interface pour créer un Workspace en mémoire.
        Dim pName As IName = Nothing                            'Interface pour ouvrir un workspace.
        Dim pFeatureWorkspace As IFeatureWorkspace = Nothing    'Interface contenant un FeatureWorkspace.
        Dim pFields As IFields = Nothing                        'Interface pour contenant les attributs de la Featureclass.
        Dim pFieldsEdit As IFieldsEdit = Nothing                'Interface pour créer les attributs.
        Dim pFieldEdit As IFieldEdit = Nothing                  'Interface pour créer un attribut.
        Dim pClone As IClone = Nothing                          'Interface utilisé pour cloner un attribut.
        Dim pGeometryDef As IGeometryDefEdit = Nothing          'Interface ESRI utilisé pour créer la structure d'un géométrie.
        Dim qFactoryType As Type = Nothing                      'Interface contenant le type d'objet à créer.
        Dim pUID As New UID                                     'Interface pour générer un UID.

        'Définir la valeur par défaut
        CreateInMemoryFeatureClass = Nothing

        Try
            'Définir le type de Factory
            qFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.InMemoryWorkspaceFactory")
            'Générer l'interface pour créer un Workspace en mémoire
            pWorkspaceFactory = CType(Activator.CreateInstance(qFactoryType), IWorkspaceFactory)

            'Créer un nouveau UID
            pUID.Generate()

            'Creer un workspace en mémoire
            pName = CType(pWorkspaceFactory.Create("", pUID.Value.ToString, Nothing, 0), IName)

            'Définir le FeatureWorkspace pour créer une Featureclass
            pFeatureWorkspace = CType(pName.Open, IFeatureWorkspace)

            'Définir le type d'élément
            pUID.Value = "esriGeodatabase.Feature"

            'Interface pour créer des attributs
            pFieldsEdit = New Fields

            'Définir le nombre d'attributs
            pFieldsEdit.FieldCount_2 = 4

            'Créer l'attribut du OBJECTID
            pFieldEdit = New Field
            With pFieldEdit
                .Name_2 = "OBJECTID"
                .AliasName_2 = "OBJECTID"
                .Type_2 = esriFieldType.esriFieldTypeOID
            End With
            'Ajouter l'attribut
            pFieldsEdit.Field_2(0) = pFieldEdit

            'Créer l'attribut de description
            pFieldEdit = New Field
            With pFieldEdit
                .Name_2 = "DESCRIPTION"
                .AliasName_2 = "DESCRIPTION"
                .Type_2 = esriFieldType.esriFieldTypeString
                .Length_2 = 256
                .IsNullable_2 = True
            End With
            'Ajouter l'attribut
            pFieldsEdit.Field_2(1) = pFieldEdit

            'Créer l'attribut de la valeur obtenue
            pFieldEdit = New Field
            With pFieldEdit
                .Name_2 = "VALEUR"
                .AliasName_2 = "VALEUR"
                .Type_2 = esriFieldType.esriFieldTypeSingle
                .IsNullable_2 = True
            End With
            'Ajouter l'attribut
            pFieldsEdit.Field_2(2) = pFieldEdit

            'Interface pour extraire l'attribut de géométrie de la Featureclass
            pFields = pFeatureLayer.FeatureClass.Fields
            'Définir l'attribut de Géométrie
            pFieldEdit = CType(pFields.Field(pFields.FindField(pFeatureLayer.FeatureClass.ShapeFieldName)), IFieldEdit)
            'Interface pour clone l'attribut
            pClone = CType(pFieldEdit, IClone)
            'Cloner l'attribut
            pFieldEdit = CType(pClone.Clone, IFieldEdit)
            'Vérifier si le Type de Géométrie n'est pas spécifié
            If pEsriGeometryType <> esriGeometryType.esriGeometryNull Then
                'Interface pour définir le type de géométrie
                pGeometryDef = CType(pFieldEdit.GeometryDef, IGeometryDefEdit)
                'Définir le type de géométrie
                pGeometryDef.GeometryType_2 = pEsriGeometryType
                'Enlever le Z
                pGeometryDef.HasZ_2 = False
                'Enlever le M
                pGeometryDef.HasM_2 = False
            End If
            'Ajouter l'attribut
            pFieldsEdit.Field_2(3) = pFieldEdit

            'Créer la Featureclass
            CreateInMemoryFeatureClass = pFeatureWorkspace.CreateFeatureClass(sNom, pFieldsEdit, pUID, Nothing, esriFeatureType.esriFTSimple, pFeatureLayer.FeatureClass.ShapeFieldName, "")

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pWorkspaceFactory = Nothing
            pName = Nothing
            pFeatureWorkspace = Nothing
            pFields = Nothing
            pFieldsEdit = Nothing
            pFieldEdit = Nothing
            pClone = Nothing
            pGeometryDef = Nothing
            qFactoryType = Nothing
            pUID = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'afficher la FeatureClass d'erreur dans la Map et dans le TableWindow d'attributs.
    '''</summary>
    ''' 
    '''<param name="sDescription"> Description de l'erreur à écrire.</param>
    '''<param name="pGeometry"> Géométrie de l'erreur à écrire.</param>
    '''<param name="dValeur"> Valeur obtenue de l'erreur à écrire.</param>
    '''
    Protected Sub EcrireFeatureErreur(ByVal sDescription As String, ByVal pGeometry As IGeometry, Optional ByVal dValeur As Single = Nothing)
        'Déclarer les variables de travail
        Dim pFeatureBuffer As IFeatureBuffer = Nothing      'Interface ESRI contenant l'élément de l'incohérence à créer.
        Dim pClone As IClone = Nothing                      'Interface pour cloner une géométrie.

        Try
            'Sortir si la classe d'erreurs est absente
            If gpFeatureCursorErreur Is Nothing Then Exit Sub

            'Créer un FeatureBuffer Point
            pFeatureBuffer = gpFeatureClassErreur.CreateFeatureBuffer

            'Interface pour cloner la géométrie
            pClone = CType(pGeometry, IClone)

            'Définir la géométrie
            pFeatureBuffer.Shape = CType(pClone.Clone, IGeometry)

            'Définir la description
            pFeatureBuffer.Value(1) = sDescription

            'Définir la valeur obtenue
            pFeatureBuffer.Value(2) = dValeur

            'Insérer un nouvel élément dans la FeatureClass d'erreur
            gpFeatureCursorErreur.InsertFeature(pFeatureBuffer)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureBuffer = Nothing
            pClone = Nothing
        End Try
    End Sub

    ' '''<summary>
    ' ''' Routine qui permet d'afficher la FeatureClass d'erreur dans la Map et dans le TableWindow d'attributs.
    ' ''' 
    ' '''<param name="pMap"> Interface contenant la Map dans lequel le FeatureLayer sera ajouté.</param>
    ' '''<param name="pFeatureClass"> FeatureClass à ajouter dans la Map.</param>
    ' '''<param name="sNom"> Nom du FeatureLayer à ajouter dans la Map.</param>
    ' ''' 
    ' '''</summary>
    ' '''
    'Protected Sub AfficherFeatureClassErreur(ByVal pMap As IMap, ByVal pFeatureClass As IFeatureClass, ByVal sNom As String)
    '    'Déclarer les variables de travail
    '    Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface ESRI contenant la Featureclass des éléments en mémoire.
    '    Dim pTableWindow2 As ITableWindow2 = Nothing        'Interface qui permet de verifier la présence du menu des tables d'attributs et de les manipuler.
    '    Dim pExistTableWindow As ITableWindow = Nothing     'Interface contenant le menu de la table d'attributs existente.
    '    'Dim pGeoFeatureLayer As IGeoFeatureLayer = Nothing  'Interface pour extraire le Renderer contenant la symbologie.
    '    'Dim pSimpleRenderer As ISimpleRenderer = Nothing    'Interface contenant la symbologie.

    '    Try
    '        'Sortie si la classe d'erreur est absente
    '        If pFeatureClass Is Nothing Then Exit Sub

    '        'Créer un nouveau FeatureLayer
    '        pFeatureLayer = New FeatureLayer
    '        'Définir le nom du FeatureLayer selon le nom et la date
    '        pFeatureLayer.Name = sNom
    '        'Rendre visible le FeatureLayer
    '        pFeatureLayer.Visible = True
    '        ''Interface pour changer la symbologie
    '        'pGeoFeatureLayer = CType(pFeatureLayer, IGeoFeatureLayer)
    '        'pSimpleRenderer = CType(pGeoFeatureLayer.Renderer, ISimpleRenderer)
    '        'Définir la Featureclass dans le FeatureLayer
    '        pFeatureLayer.FeatureClass = gpFeatureClassErreur
    '        'Ajouter le FeatureLayer dans la Map
    '        pMap.AddLayer(pFeatureLayer)

    '        'Vérifier si on doit afficher la table d'erreur
    '        If gbAfficherTableErreur = False Then Exit Sub

    '        'Interface pour vérifier la présence du menu des tables d'attributs
    '        pTableWindow2 = CType(New TableWindow, ITableWindow2)

    '        'Définir le menu de la table d'attribut de la table s'il est présent
    '        pExistTableWindow = pTableWindow2.FindViaLayer(pFeatureLayer)

    '        'Vérifier si le menu de la table d'attribut est absent
    '        If pExistTableWindow Is Nothing Then
    '            'Définir le FeatureLayer à afficher
    '            pTableWindow2.Layer = pFeatureLayer
    '        End If

    '        'Vérifier le menu de la table d'attribut est absent
    '        If pExistTableWindow Is Nothing Then
    '            'Définir les paramètre d'affichage du menu des tables d'attributs
    '            pTableWindow2.TableSelectionAction = esriTableSelectionActions.esriSelectFeatures
    '            pTableWindow2.ShowSelected = False
    '            pTableWindow2.ShowAliasNamesInColumnHeadings = True
    '            pTableWindow2.Application = gpApplication

    '            'Si le menu de la table d'attribut est présent
    '        Else
    '            'Redéfinir le menu des tables d'attributs pour celui existant
    '            pTableWindow2 = CType(pExistTableWindow, ITableWindow2)
    '        End If

    '        'Afficher le menu des tables d'attributs s'il n'est pas affiché
    '        If Not pTableWindow2.IsVisible Then pTableWindow2.Show(True)

    '    Catch ex As Exception
    '        'Retourner l'erreur
    '        Throw ex
    '    Finally
    '        'Vider la mémoire
    '        pFeatureLayer = Nothing
    '        pTableWindow2 = Nothing
    '        pExistTableWindow = Nothing
    '    End Try
    'End Sub

    '''<summary>
    ''' Routine qui permet de créer un Multipoint à partir d'une géométrie.
    ''' 
    '''<param name="pGeometry"> Interface contenant la géométrie utilisée pour créer le Multipoint.</param>
    '''
    '''<return>Le Multipoint contenant les points de la géométrie spécifiée.</return>
    ''' 
    '''</summary>
    '''
    Protected Function GeometrieToMultiPoint(ByVal pGeometry As IGeometry) As IMultipoint
        'Déclarer les variables de travail
        Dim pGeomColl As IPointCollection = Nothing     'Interface pour ajouter les sommets de la géométrie.
        Dim pPointColl As IPointCollection = Nothing    'Interface pour extraire les points de la géométrie.
        Dim pClone As IClone = Nothing                  'Interface pour cloner la géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface ESRI utilisée pour simplifier la géométrie.

        'Définir la valeur par défaut
        GeometrieToMultiPoint = New Multipoint
        'Définir la référence spatial
        GeometrieToMultiPoint.SpatialReference = pGeometry.SpatialReference

        Try
            'Interface pour ajouter les points de la géométrie dans le multipoint
            pGeomColl = CType(GeometrieToMultiPoint, IPointCollection)

            'Interface pour clone la géométrie
            pClone = CType(pGeometry, IClone)

            'Vérifier si la géométrie est un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Ajouter le point dans le multipoint
                pGeomColl.AddPoint(CType(pClone.Clone, IPoint))
                'Si la géométrie est un multipoint
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Retourner le Multipoint
                GeometrieToMultiPoint = CType(pClone.Clone, IMultipoint)
                'Si la géométrie est un Polyline
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                'Convertir le polyline en multipoint
                pGeomColl.AddPointCollection(CType(pClone.Clone, IPointCollection))
                'Simplifier la géométrie
                pTopoOp = CType(pGeomColl, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
                'Si la géométrie est un Polygon
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Convertir le polygon en multipoint
                pGeomColl.AddPointCollection(CType(pClone.Clone, IPointCollection))
                'Simplifier la géométrie
                pTopoOp = CType(pGeomColl, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPointColl = Nothing
            pClone = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer un Polyline à partir d'une géométrie.
    ''' 
    '''<param name="pGeometry"> Interface contenant la géométrie utilisée pour créer le Polyline.</param>
    '''
    '''<return>Le Polyline contenant les lignes de la géométrie spécifiée.</return>
    ''' 
    '''</summary>
    '''
    Protected Function GeometrieToPolyline(ByVal pGeometry As IGeometry) As IPolyline
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes de la géométrie.
        Dim pPolylineColl As IGeometryCollection = Nothing  'Interface pour extraire les lignes de la géométrie.
        Dim pClone As IClone = Nothing                      'Interface pour cloner la géométrie.

        'Définir la valeur par défaut
        GeometrieToPolyline = New Polyline
        'Définir la référence spatial
        GeometrieToPolyline.SpatialReference = pGeometry.SpatialReference

        Try
            'Interface pour ajouter les lignes de la géométrie dans le polyline
            pGeomColl = CType(GeometrieToPolyline, IGeometryCollection)

            'Interface pour clone la géométrie
            pClone = CType(pGeometry, IClone)

            'Vérifier si la géométrie est un polyline
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                'Retourner le polyline
                GeometrieToPolyline = CType(pClone.Clone, IPolyline)
                'Si la géométrie est un Polygon
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Convertir le polygon en polyline
                pGeomColl.AddGeometryCollection(CType(pClone.Clone, IGeometryCollection))
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPolylineColl = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer un Polyline à partir d'une Line.
    ''' 
    '''<param name="pLine"> Interface contenant une Line.</param>
    '''
    '''<return>Le Polyline contenant la Line spécifiée.</return>
    ''' 
    '''</summary>
    '''
    Protected Function LineToPolyline(ByVal pLine As ILine) As IPolyline
        'Déclarer les variables de travail
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter des sommets.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisée pour simplifier la géométrie.
        Dim pClone As IClone = Nothing                      'Interface pour cloner la géométrie.

        'Définir la valeur par défaut
        LineToPolyline = New Polyline
        'Définir la référence spatial
        LineToPolyline.SpatialReference = pLine.SpatialReference

        Try
            'Interface pour ajouter la Line dans la Polyline
            pPointColl = CType(LineToPolyline, IPointCollection)

            'Interface pour cloner la géométrie
            pClone = CType(pLine, IClone)
            'cloner la géométrie
            pLine = CType(pClone.Clone, ILine)

            'Ajouter le premier point de la Line
            pPointColl.AddPoint(pLine.FromPoint)

            'Ajouter le dernier point de la Line
            pPointColl.AddPoint(pLine.ToPoint)

            'Interface pour simplifier
            pTopoOp = CType(LineToPolyline, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPointColl = Nothing
            pTopoOp = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer un Polyline à partir d'un Path ou un Ring.
    ''' 
    '''<param name="pPath"> Interface contenant un Path.</param>
    '''
    '''<return>Le Polyline contenant le Path spécifié.</return>
    ''' 
    '''</summary>
    '''
    Protected Function PathToPolyline(ByVal pPath As IPath) As IPolyline
        'Déclarer les variables de travail
        Dim pPointColl As IPointCollection = Nothing      'Interface pour ajouter les Path.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisée pour simplifier la géométrie.
        Dim pClone As IClone = Nothing                      'Interface pour cloner la géométrie.

        'Définir la valeur par défaut
        PathToPolyline = New Polyline
        'Définir la référence spatial
        PathToPolyline.SpatialReference = pPath.SpatialReference

        Try
            'Interface pour ajouter le Path dans la Polyline
            pPointColl = CType(PathToPolyline, IPointCollection)

            'Interface pour cloner la géométrie
            pClone = CType(pPath, IClone)

            'Ajouter le Path
            pPointColl.AddPointCollection(CType(pClone.Clone, IPointCollection))

            'Interface pour simplifier
            pTopoOp = CType(PathToPolyline, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPointColl = Nothing
            pTopoOp = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer un Polyline à partir d'un Polygon.
    ''' 
    '''<param name="pPolygon"> Interface contenant un Polygon.</param>
    '''
    '''<return>Le Polyline contenant les anneaux du Polygon.</return>
    ''' 
    '''</summary>
    '''
    Protected Function PolygonToPolyline(ByVal pPolygon As IPolygon) As IPolyline
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes.
        Dim pRingColl As IGeometryCollection = Nothing      'Interface pour extraire les anneaux.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisée pour simplifier la géométrie.

        'Définir la valeur par défaut
        PolygonToPolyline = New Polyline
        'Définir la référence spatial
        PolygonToPolyline.SpatialReference = pPolygon.SpatialReference

        Try
            'Interface pour ajouter les anneaux dans le Polyline
            pGeomColl = CType(PolygonToPolyline, IGeometryCollection)

            'Interface pour extraire les anneaux du Polygon
            pRingColl = CType(pPolygon, IGeometryCollection)

            'Traiter tous les anneaux intérieurs
            For i = 0 To pRingColl.GeometryCount - 1
                'Ajouter les anneaux dans la Polyline
                pGeomColl.AddGeometryCollection(CType(PathToPolyline(CType(pRingColl.Geometry(i), IPath)), IGeometryCollection))
            Next

            'Interface pour simplifier
            pTopoOp = CType(PolygonToPolyline, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pRingColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer un Polygon à partir d'un anneau extérieur et un GeometryBag d'anneaux intérieurs.
    ''' 
    '''<param name="pRingExt"> Interface contenant un anneau extérieur.</param>
    '''<param name="pGeomCollInt"> Interface contenant les anneaux intérieurs.</param>
    '''
    '''<return>Le Polygon contenant les anneaux spécifiés.</return>
    ''' 
    '''</summary>
    '''
    Protected Function RingToPolygon(ByVal pRingExt As IRing, ByVal pGeomCollInt As IGeometryCollection) As IPolygon
        'Déclarer les variables de travail
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieurs.
        Dim pRingInt As IRing = Nothing                     'Interface contenant l'anneau intérieur.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisée pour simplifier la géométrie.
        Dim pClone As IClone = Nothing                      'Interface pour cloner la géométrie.

        'Définir la valeur par défaut
        RingToPolygon = New Polygon
        'Définir la référence spatial
        RingToPolygon.SpatialReference = pRingExt.SpatialReference

        Try
            'Interface pour ajouter les anneaux dans le Polygone
            pGeomCollExt = CType(RingToPolygon, IGeometryCollection)

            'Interface pour cloner la géométrie
            pClone = CType(pRingExt, IClone)

            'Ajouter l'anneau extérieur
            pGeomCollExt.AddGeometry(CType(pClone.Clone, IGeometry))

            'Traiter tous les anneaux intérieurs
            For i = 0 To pGeomCollInt.GeometryCount - 1
                'Interface contenant le Ring intérieur
                pRingInt = CType(pGeomCollInt.Geometry(i), IRing)
                'Projeter l'anneau
                'pRingInt.Project(pRingExt.SpatialReference)
                'Interface pour cloner la géométrie
                pClone = CType(pRingInt, IClone)
                'Ajouter les anneaux intérieurs
                pGeomCollExt.AddGeometry(CType(pClone.Clone, IGeometry))
            Next

            'Interface pour simplifier
            pTopoOp = CType(RingToPolygon, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomCollExt = Nothing
            pRingInt = Nothing
            pTopoOp = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de transformer un Enveloppe en Polygon. 
    ''' Le Polygon va contenir cinq Points.
    '''</summary>
    '''
    '''<param name="pEnvelope">Interface ESRI contenant l'enveloppe à traiter.</param>
    ''' 
    '''<returns>La fonction va retourner un "IPolygon" contenant cinq "Point". Sinon "Nothing".</returns>
    '''
    Protected Function EnvelopeToPolygon(ByVal pEnvelope As IEnvelope) As IPolygon
        'Déclarer les variables de travail
        Dim pPointColl As IPointCollection = Nothing    'Interface utilisé pour ajouter des sommets au Polygon

        'Créer une nouvelle surface vide par défaut
        EnvelopeToPolygon = New Polygon
        'Définir la référence spatiale
        EnvelopeToPolygon.SpatialReference = pEnvelope.SpatialReference

        Try
            'Créer une nouvelle surface vide
            pPointColl = CType(EnvelopeToPolygon, IPointCollection)

            'Définir le coin SE
            pPointColl.AddPoint(pEnvelope.LowerRight)

            'Définir le coin SW
            pPointColl.AddPoint(pEnvelope.LowerLeft)

            'Définir le coin NW
            pPointColl.AddPoint(pEnvelope.UpperLeft)

            'Définir le coin NE
            pPointColl.AddPoint(pEnvelope.UpperRight)

            'Fermer le polygone
            EnvelopeToPolygon.Close()

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pPointColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner le centre d'une géométrie.
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments qui respectent ou non la comparaison avec ses éléments en relation.</return>
    ''' 
    '''</summary>
    '''
    Protected Function CentreGeometrie(ByRef pGeometry As IGeometry) As IPoint
        'Déclarer les variables de travail
        Dim pArea As IArea = Nothing                        'Interface pour extraire un point à l'intérieur du polygon.
        Dim pPolyline As IPolyline = Nothing                'Interface pour extraire un point sur la ligne.
        Dim pPath As IPath = Nothing                        'Interface pour extraire un point sur le Path.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire un point sur le multipoint.
        Dim pPoint As IPoint = New Point                    'Interface contenant le point du centre de la ligne.
        Dim pClone As IClone = Nothing                      'Interface pour cloner une géométrie.

        'Définir la géométrie par défaut
        CentreGeometrie = Nothing

        Try
            'Vérifier si la géométrie est vide
            If pGeometry.IsEmpty Then
                'Interface pour extraire le centre de la classe de sélection
                pArea = CType(gpEnvelope, IArea)
                'Extraire le centre de la classe de sélection
                pPoint = pArea.LabelPoint

            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour extraire un point à l'intérieur du polygon
                pArea = CType(pGeometry, IArea)
                'Extraire le centroide intérieur
                pPoint = pArea.LabelPoint

                'si la géométrie est un Polyline
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                'Interface pour extraire un point sur la ligne
                pPolyline = CType(pGeometry, IPolyline)
                'Extraire le point au centre de la ligne
                pPolyline.QueryPoint(esriSegmentExtension.esriNoExtension, pPolyline.Length / 2, False, pPoint)

                'si la géométrie est un Polyline
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPath Or pGeometry.GeometryType = esriGeometryType.esriGeometryRing Then
                'Interface pour extraire un point sur la ligne
                pPath = CType(pGeometry, IPath)
                'Extraire le point au centre de la ligne
                pPath.QueryPoint(esriSegmentExtension.esriNoExtension, pPath.Length / 2, False, pPoint)

                'si la géométrie est un Multipoint
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Interface pour extraire un point sur le multipoint
                pPointColl = CType(pGeometry, IPointCollection)
                'Extraire le premier point sur le multipoint
                pClone = CType(pPointColl.Point(0), IClone)
                'Cloner le premier point sur le multipoint
                pPoint = CType(pClone.Clone, IPoint)

                'si la géométrie est un Point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Extraire le premier point sur le multipoint
                pClone = CType(pGeometry, IClone)
                'Cloner le premier point sur le multipoint
                pPoint = CType(pClone.Clone, IPoint)
            End If

            'Définir le centre de la géométrie
            CentreGeometrie = pPoint

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pArea = Nothing
            pPolyline = Nothing
            pPath = Nothing
            pPointColl = Nothing
            pPoint = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner le premier sommet de chaque composante d'une géométrie.
    ''' 
    '''<param name="pGeometry"> Interface contenant la géométrie utilisée pour créer le Multipoint.</param>
    '''
    '''<return>Le Multipoint contenant le premier sommet de chaque composante d'une géométrie.</return>
    ''' 
    '''</summary>
    '''
    Protected Function PremierSommetComposanteGeometrie(ByVal pGeometry As IGeometry) As IMultipoint
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter les sommets de la géométrie.
        Dim pPointColl As IPointCollection = Nothing    'Interface pour extraire les points de la géométrie.
        Dim pClone As IClone = Nothing                  'Interface pour cloner la géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface ESRI utilisée pour simplifier la géométrie.
        Dim pPath As IPath = Nothing                    'Interface contenant la composante de la géométrie

        'Définir la valeur par défaut
        PremierSommetComposanteGeometrie = New Multipoint
        'Définir la référence spatial
        PremierSommetComposanteGeometrie.SpatialReference = pGeometry.SpatialReference

        Try
            'Interface pour ajouter les points de la géométrie dans le multipoint
            pPointColl = CType(PremierSommetComposanteGeometrie, IPointCollection)

            'Interface pour clone la géométrie
            pClone = CType(pGeometry, IClone)

            'Vérifier si la géométrie est un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Ajouter le point dans le multipoint
                pPointColl.AddPoint(CType(pClone.Clone, IPoint))

                'Si la géométrie est un multipoint
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Retourner le Multipoint
                PremierSommetComposanteGeometrie = CType(pClone.Clone, IMultipoint)

                'Si la géométrie est un Polyline ou un polygon
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Or _
                   pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then

                'Interface pour extraire les composantes de la géométrie
                pGeomColl = CType(pGeometry, IGeometryCollection)

                'Traiter toutes les composantes
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Interface pour extraire le premier sommet de la composante
                    pPath = CType(pGeomColl.Geometry(i), IPath)
                    'Ajouter le point dans le multipoint
                    pPointColl.AddPoint(pPath.FromPoint)
                Next

                'Simplifier la géométrie résultante
                pTopoOp = CType(pPointColl, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPointColl = Nothing
            pClone = Nothing
            pTopoOp = Nothing
            pPath = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner l'enveloppe des éléments sélectionnés dans une Map.
    ''' 
    '''<param name="pMap"> Interface contenant les FeatureLayers des éléments.</param>
    '''
    '''<return>L'enveloppe contenant les éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Protected Function EnveloppeMapSelection(ByRef pMap As IMap) As IEnvelope
        'Déclarer les variables de travail
        Dim pEnumFeature As IEnumFeature = Nothing  'Interface pour extraire les éléments sélectionnés.
        Dim pFeature As IFeature = Nothing          'Interface contenant un élément.
        Dim pEnvelope As IEnvelope = Nothing        'Interface contenant l'envelope d'un élément.

        'Définir l'enveloppe vide par défaut
        EnveloppeMapSelection = Nothing

        Try
            'Interface pour extraire les éléments sélectionnés
            pEnumFeature = CType(pMap.FeatureSelection, IEnumFeature)

            'Initialiser l'énumération
            pEnumFeature.Reset()

            'Extraire le premier élément
            pFeature = pEnumFeature.Next

            Do While Not pFeature Is Nothing
                'Extraire l'enveloppe de l'élément
                pEnvelope = pFeature.Extent

                'Projeter selon la référence spatiale
                pEnvelope.Project(pMap.SpatialReference)

                'Si l'enveloppe est vide
                If EnveloppeMapSelection Is Nothing Then
                    'Définir l'enveloppe de départ
                    EnveloppeMapSelection = pEnvelope
                Else
                    'Union entre les enveloppes
                    EnveloppeMapSelection.Union(pFeature.Extent)
                End If

                'Extraire le prochain élément
                pFeature = pEnumFeature.Next
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de retourner l'enveloppe des éléments sélectionnés dans un SelectionSet.
    ''' 
    '''<param name="pSelectionSet"> Interface contenant les éléments sélectionnés d'un FeatureLayer.</param>
    '''<param name="pSpatialReference"> Interface contenant la référence spatiale de l'enveloppe.</param>
    '''
    '''<return>L'enveloppe contenant les éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Protected Function EnveloppeSelectionSet(ByRef pSelectionSet As ISelectionSet, ByVal pSpatialReference As ISpatialReference) As IEnvelope
        'Déclarer les variables de travail
        Dim pCursor As ICursor = Nothing                'Interface pour extraire les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface pour extraire les éléments sélectionnés.
        Dim pFeature As IFeature = Nothing              'Interface contenant un élément.
        Dim pEnvelope As IEnvelope = Nothing            'Interface contenant l'envelope d'un élément.

        'Définir l'enveloppe vide par défaut
        EnveloppeSelectionSet = Nothing

        Try
            'Interface pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature

            Do While Not pFeature Is Nothing
                'Extraire l'enveloppe de l'élément
                pEnvelope = pFeature.Extent

                'Projeter selon la référence spatiale
                pEnvelope.Project(pSpatialReference)

                'Si l'enveloppe est vide
                If EnveloppeSelectionSet Is Nothing Then
                    'Définir l'enveloppe de départ
                    EnveloppeSelectionSet = pEnvelope
                Else
                    'Union entre les enveloppes
                    EnveloppeSelectionSet.Union(pEnvelope)
                End If

                'Extraire le prochain élément
                pFeature = pFeatureCursor.NextFeature
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de créer et retourner la topologie en mémoire des éléments entre une collection de FeatureLayer.
    '''</summary>
    '''
    '''<param name="pEnvelope">Interface ESRI contenant l'enveloppe de création de la topologie traitée.</param>
    '''<param name="qFeatureLayerColl">Interface ESRI contenant les FeatureLayers pour traiter la topologie.</param> 
    '''<param name="dTolerance">Tolerance de proximité.</param> 
    '''
    '''<returns>"ITopologyGraph" contenant la topologie des classes de données, "Nothing" sinon.</returns>
    '''
    Protected Function CreerTopologyGraph(ByVal pEnvelope As IEnvelope, ByVal qFeatureLayerColl As Collection, ByVal dTolerance As Double) As ITopologyGraph4
        'Déclarer les variables de travail
        Dim qType As Type = Nothing                         'Interface contenant le type d'objet à créer.
        Dim oObjet As System.Object = Nothing               'Interface contenant l'objet correspondant à l'application.
        Dim pTopologyExt As ITopologyExtension = Nothing    'Interface contenant l'extension de la topologie.
        Dim pMapTopology As IMapTopology2 = Nothing         'Interface utilisé pour créer la topologie.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la classe de données.

        'Définir la valeur de retour par défaut
        CreerTopologyGraph = Nothing

        Try
            'Définir l'extension de topologie
            qType = Type.GetTypeFromProgID("esriEditorExt.TopologyExtension")
            oObjet = Activator.CreateInstance(qType)
            pTopologyExt = CType(oObjet, ITopologyExtension)

            'Définir l'interface pour créer la topologie
            pMapTopology = CType(pTopologyExt.MapTopology, IMapTopology2)

            'S'assurer que laliste des Layers est vide
            pMapTopology.ClearLayers()

            'Traiter tous les FeatureLayer présents
            For Each pFeatureLayer In qFeatureLayerColl
                'Ajouter le FeatureLayer à la topologie
                pMapTopology.AddLayer(pFeatureLayer)
            Next

            'Changer la référence spatiale selon l'enveloppe
            pMapTopology.SpatialReference = pEnvelope.SpatialReference

            'Définir la tolérance de connexion et de partage
            pMapTopology.ClusterTolerance = dTolerance

            'Interface pour construre la topologie
            pTopologyGraph = CType(pMapTopology.Cache, ITopologyGraph4)
            pTopologyGraph.SetEmpty()

            Try
                'Construire la topologie
                pTopologyGraph.Build(pEnvelope, False)
            Catch ex As OutOfMemoryException
                'Retourner une erreur de création de la topologie
                Throw New Exception("Incapable de créer la topologie : OutOfMemoryException")
            Catch ex As Exception
                'Retourner une erreur de création de la topologie
                Throw New Exception("Incapable de créer la topologie : " & ex.Message)
            End Try

            'Retourner la topologie
            CreerTopologyGraph = pTopologyGraph

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            qType = Nothing
            oObjet = Nothing
            pTopologyExt = Nothing
            pTopologyGraph = Nothing
            pMapTopology = Nothing
            pFeatureLayer = Nothing
            'Récupération de la mémoire disponible
            GC.Collect()
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de convertir et indiquer si la valeur numérique d'une chaine de caractère représentant un double.
    ''' 
    ''' Cette fonction tient compte des paramètres du "Regional settings" selon le langage spécifié.
    '''</summary>
    '''
    '''<param name="sValeur">Valeur numérique d'une chaine de caractère représentant un double.</param>
    '''
    '''<returns>"Boolean" qui indique si la valeur peut être convertit en double.</returns>
    '''
    Protected Function TestDBL(ByRef sValeur As String) As Boolean
        'Déclarer les variables de travail
        Dim sRetVal As String = Nothing     'Valeur de retour convertit.

        Try
            'Indique que la valeur n'est pas numérique par défaut
            TestDBL = False

            'Vérifier si la valeur est numérique
            If IsNumeric(sValeur) Then
                'Indique que la valeur est numérique
                TestDBL = True

                'Si la valeur n'est pas numérique
            Else
                'Tester si le séparateur de décimal est une virgule.
                If IsNumeric("3,24") Then
                    'Remplacer le point par une virgule
                    sRetVal = sValeur.Replace(".", ",")

                    'Vérifier si la valeur est numérique
                    If IsNumeric(sRetVal) Then
                        'Indique que la valeur est numérique
                        TestDBL = True
                        'Redéfinir la valeur
                        sValeur = sRetVal
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            'Throw
        Finally
            'Vider la mémoire
            sRetVal = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de convertir et retourner la valeur numérique d'une chaine de caractère représentant un double.
    ''' 
    ''' Cette fonction tient compte des paramètres du "Regional settings" selon le langage spécifié.
    '''</summary>
    '''
    '''<param name="sValeur">Valeur numérique d'une chaine de caractère représentant un double.</param>
    '''
    '''<returns>"Double" contenant la valeur convertit.</returns>
    '''
    Protected Function ConvertDBL(ByVal sValeur As String) As Double
        'Déclarer les variables de travail
        Dim sRetVal As String = Nothing     'Valeur de retour convertit.
        Dim sSepDecimal As String = Nothing 'Séparateur de décimal.
        Dim sSepMiles As String = Nothing   'Séparateur de miles.

        Try
            'Initialiser la valeur de retour
            If sValeur = "" Then sValeur = "0"
            sRetVal = Replace(Trim(sValeur), " ", "")

            'Vérifier si le séparateur de miles est présent
            If CDbl("3,24") = 324 Then
                'Définir le séparateur de décimal
                sSepDecimal = "."
                'Définir le séparateur de miles
                sSepMiles = ","

                'Si le séparateur de miles est absent
            Else
                'Définir le séparateur de décimal
                sSepDecimal = ","
                'Définir le séparateur de miles
                sSepMiles = "."
            End If

            'Vérifier si le texte contient un séparateur de décimal
            If InStr(sRetVal, sSepDecimal) > 0 Then
                'Vérifier si le texte contient un séparateur de miles
                If InStr(sRetVal, sSepMiles) > 0 Then
                    'Vérifier si la position du séparateur de décimal est placé après celui du miles
                    If InStr(sRetVal, sSepDecimal) > InStr(sRetVal, sSepMiles) Then
                        'Définir la valeur de retour
                        sRetVal = Replace(sRetVal, sSepMiles, "")
                    Else
                        'Définir la valeur de retour
                        sRetVal = Replace(sRetVal, sSepDecimal, "")
                        sRetVal = Replace(sRetVal, sSepMiles, sSepDecimal)
                    End If
                End If
            Else
                'Définir la valeur de retour
                sRetVal = Replace(sRetVal, sSepMiles, sSepDecimal)
            End If

            'Retourner la valeur
            ConvertDBL = CDbl(sRetVal)

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            sRetVal = Nothing
            sSepDecimal = Nothing
            sSepMiles = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de corriger la topologie des éléments de toutes les classes de données spécifiées.
    '''</summary>
    '''
    '''<param name="qFeatureLayerColl">Collection des FeatureLayers à traiter.</param>
    '''<param name="dPrecision">Contient la précision des coordonnées XY.</param>
    '''<param name="bCorriger">Indiquer si on doit corriger.</param>
    '''
    Public Sub CorrigerTopologieFeatureLayerColl(ByVal qFeatureLayerColl As Collection, ByVal dPrecision As Double, Optional ByVal bCorriger As Boolean = False)
        'Définir les variables de travail
        Dim qFeatureColl As Collection = Nothing        'Objet contenant le liste des éléments en relation
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant les paramètres d'affichage d'une classe de données
        Dim pFeatureClass As IFeatureClass = Nothing    'Interface contenant la classe de données
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface qui permet de sélectionner tous les éléments du FeatureLayer 
        Dim pSelectionSet As ISelectionSet2 = Nothing   'Interface utilisé pour vérifier le nombre d'éléments sélectionnés
        Dim pCursor As ICursor = Nothing                'Interface utilisé pour extraire les éléments à traiter
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface utilisé pour extraire les éléments à traiter
        Dim pFeature As IFeature = Nothing              'Interface contenant l'élément à traiter
        Dim pTopologyGraph As ITopologyGraph = Nothing  'Interface contenant la topologie
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie de l'élément
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Interface contenant l'information sur la classe de données
            pFeatureLayer = CType(qFeatureLayerColl.Item(1), IFeatureLayer)

            'Créer la topologie
            pTopologyGraph = CreerTopologyGraph(pFeatureLayer.AreaOfInterest, qFeatureLayerColl, dPrecision)

            'Traiter toutes les classes
            For i = 1 To qFeatureLayerColl.Count
                'Interface contenant l'information sur la classe de données
                pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                'Interface contenant l'information sur la classe de données
                pFeatureClass = pFeatureLayer.FeatureClass
                'Interfaces pour traiter les éléments sélectionnés
                pFeatureSel = CType(pFeatureLayer, IFeatureSelection)
                'Interfaces pour vérifier les éléments sélectionnés
                pSelectionSet = CType(pFeatureSel.SelectionSet, ISelectionSet2)
                'Vérifier si des éléments sont sélectionnés
                If pSelectionSet.Count = 0 Then
                    'Sélectionnées tous les éléments du FeatureLayer
                    pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                    pSelectionSet = CType(pFeatureSel.SelectionSet, ISelectionSet2)
                End If

                'Vérifier si on doit corriger
                If bCorriger Then
                    'Interfaces pour extraire les éléments sélectionnés
                    pSelectionSet.Update(Nothing, False, pCursor)
                Else
                    'Interfaces pour extraire les éléments sélectionnés
                    pSelectionSet.Search(Nothing, False, pCursor)
                End If
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Définir la nouvelle géométrie de l'élément
                    pGeometry = pTopologyGraph.GetParentGeometry(pFeatureClass, pFeature.OID)

                    'Vérifier si la géométrie est invalide
                    If pGeometry Is Nothing Then
                        'Vider la géométrie
                        pFeature.Shape.SetEmpty()
                        'Détruire l'élément au besoin
                        If bCorriger Then pFeatureCursor.DeleteFeature()
                    Else
                        'Changer la géométrie de l'élément
                        pFeature.Shape = pGeometry
                        'Conserver la correction
                        If bCorriger Then pFeatureCursor.UpdateFeature(pFeature)
                    End If

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop

                'Conserver toutes les corrections
                If bCorriger Then pFeatureCursor.Flush()
                'Permet de libérer la mémoire sur les données des classes traitées
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatureCursor)
            Next

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            qFeatureColl = Nothing
            pFeatureLayer = Nothing
            pFeatureClass = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pTopologyGraph = Nothing
            pGeometry = Nothing
        End Try
    End Sub
#End Region
End Class
