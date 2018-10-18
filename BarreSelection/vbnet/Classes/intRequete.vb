Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework

'**
'Nom de la composante : intRequete.vb
'
'''<summary>
''' Interface qui permet l'implantation d'une requête de contrainte d'intégrité.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 15 avril 2015
'''</remarks>
''' 
Public Interface intRequete

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de retourner le nom de la requête à traiter.
    '''</summary>
    ''' 
    ReadOnly Property Nom() As String

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'application.
    '''</summary>
    ''' 
    Property Application() As IApplication

    '''<summary>
    ''' Propriété qui permet de définir et retourner le document ArcMap à traiter.
    '''</summary>
    ''' 
    Property MxDocument() As IMxDocument

    '''<summary>
    ''' Propriété qui permet de définir et retourner la Map à traiter.
    '''</summary>
    ''' 
    Property Map() As IMap

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'enveloppe des éléments sélectionnés à traiter.
    '''</summary>
    ''' 
    Property Envelope() As IEnvelope

    '''<summary>
    ''' Propriété qui permet de définir et retourner le FeatureLayer à traiter.
    '''</summary>
    ''' 
    Property FeatureLayerSelection() As IFeatureLayer

    '''<summary>
    ''' Propriété qui permet de définir et retourner la référence spatiale utilisée pour le traitement.
    '''</summary>
    ''' 
    Property SpatialReference() As ISpatialReference

    '''<summary>
    ''' Propriété qui permet de définir et retourner la collection des FeatureLayers en relation.
    '''</summary>
    ''' 
    Property FeatureLayersRelation() As Collection

    '''<summary>
    ''' Propriété qui permet de définir et retourner la collection des RasterLayers en relation.
    '''</summary>
    ''' 
    Property RasterLayersRelation() As Collection

    '''<summary>
    ''' Propriété qui permet de définir et retourner la FeatureClass des erreurs trouvées.
    '''</summary>
    ''' 
    ReadOnly Property FeatureClassErreur() As IFeatureClass

    '''<summary>
    ''' Propriété qui permet de définir et retourner le FeatureLayer des erreurs trouvées.
    '''</summary>
    ''' 
    ReadOnly Property FeatureLayerErreur() As IFeatureLayer

    '''<summary>
    ''' Propriété qui permet de définir et retourner le FeatureLayer de découpage.
    '''</summary>
    ''' 
    Property FeatureLayerDecoupage() As IFeatureLayer

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'élément de découpage.
    '''</summary>
    ''' 
    Property FeatureDecoupage() As IFeature

    '''<summary>
    ''' Propriété qui permet de définir et retourner le polygone de découpage.
    '''</summary>
    ''' 
    Property PolygoneDecoupage() As IPolygon

    '''<summary>
    ''' Propriété qui permet de définir et retourner la limite du polygone de découpage.
    '''</summary>
    ''' 
    Property LimiteDecoupage() As IPolyline

    '''<summary>
    ''' Propriété qui permet de définir et retourner le nom de l'attribut à traiter.
    '''</summary>
    ''' 
    Property NomAttribut() As String

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'expression régulière à traiter.
    '''</summary>
    ''' 
    Property Expression() As String

    '''<summary>
    ''' Propriété qui permet de définir et retourner la ligne de paramètre à traiter.
    '''</summary>
    ''' 
    Property Parametres() As String

    '''<summary>
    ''' Propriété qui permet de définir et retourner la tolérance de précision.
    '''</summary>
    ''' 
    Property Precision() As Double

    '''<summary>
    ''' Propriété qui permet de définir et retourner si on doit créer la classe d'erreurs.
    '''</summary>
    ''' 
    Property CreerClasseErreur() As Boolean

    '''<summary>
    ''' Propriété qui permet de définir et retourner si on doit afficher la table d'erreurs.
    '''</summary>
    ''' 
    Property AfficherTableErreur() As Boolean

    '''<summary>
    ''' Propriété qui permet de définir et retourner si on doit enlever les éléments de la sélection.
    '''</summary>
    ''' 
    Property EnleverSelection() As Boolean

    '''<summary>
    ''' Propriété qui permet de retourner le message de requête valide ou invalide.
    '''</summary>
    ''' 
    ReadOnly Property Message() As String

    '''<summary>
    ''' Propriété qui permet de définir et retourner la requête sous forme de commande texte.
    '''</summary>
    ''' 
    Property Commande() As String
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    ''' Routine qui permet de vider la mémoire des objets de la classe.
    '''</summary>
    '''
    Sub finalize()

    '''<summary>
    ''' Fonction qui permet de retourner la liste des paramètres possibles.
    '''</summary>
    ''' 
    Function ListeParametres() As Collection

    '''<summary>
    ''' Fonction qui permet d'indiquer si le traitement à effectuer est valide.
    '''</summary>
    ''' 
    '''<return>Boolean qui indique si le traitement à effectuer est valide.</return>
    '''
    Function EstValide() As Boolean

    '''<summary>
    ''' Fonction qui permet d'indiquer si le FeatureLayer est valide.
    '''</summary>
    '''
    '''<return>Boolean qui indique si le FeatureLayer est valide.</return>
    ''' 
    Function FeatureLayerValide() As Boolean

    '''<summary>
    ''' Fonction qui permet d'indiquer si les FeatureLayers en relation sont valides.
    '''</summary>
    '''
    '''<return>Boolean qui indique si les FeatureLayers en relation sont valides.</return>
    ''' 
    Function FeatureLayersRelationValide() As Boolean

    '''<summary>
    ''' Fonction qui permet d'indiquer si les RasterLayers en relation sont valides.
    '''</summary>
    '''
    '''<return>Boolean qui indique si les RasterLayers en relation sont valides.</return>
    ''' 
    Function RasterLayersRelationValide() As Boolean

    '''<summary>
    ''' Fonction qui permet d'indiquer si la FeatureClass est valide.
    '''</summary>
    '''
    '''<return>Boolean qui indique si la FeatureClass est valide.</return>
    ''' 
    Function FeatureClassValide() As Boolean

    '''<summary>
    ''' Fonction qui permet d'indiquer si l'attribut est valide.
    '''</summary>
    '''
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    Function AttributValide() As Boolean

    '''<summary>
    ''' Fonction qui permet d'indiquer si l'expression régulière est valide.
    '''</summary>
    '''
    '''<return>Boolean qui indique si l'expression régulière est valide.</return>
    ''' 
    Function ExpressionValide() As Boolean

    '''<summary>
    ''' Fonction qui permet de retourner la valeur de l'attribut sous forme de texte.
    '''</summary>
    ''' 
    '''<param name="pFeature">Interface contenant l'élément traité.</param>
    ''' 
    '''<return>Texte contenant la valeur de l'attribut.</return>
    ''' 
    Function ValeurAttribut(ByRef pFeature As IFeature) As String

    '''<summary>
    ''' Fonction qui permet de retourner la MapTopology contenant les éléments en relations dont la topologie peut être créée et extraite.
    ''' 
    ''' La topologie doit être créée ou extraite en mode édition selon une tolérance de précision (Cluster).
    ''' Les Nodes et les Edges sont accessible via l'interface ITopologyGraph4 du IMapTopology.cache 
    '''</summary>
    ''' 
    '''<param name="pEnvelope">Interface contenant l'enveloppe pour construire le MapTopology.</param>
    '''<param name="dTolerance">Tolerance de proximité.</param>
    ''' 
    '''<return> IMapTopology contenant la topologie des éléments en relations.</return>
    ''' 
    Function RelationsMapTopology(ByVal pEnvelope As IEnvelope, Optional ByVal dTolerance As Double = 0) As IMapTopology

    '''<summary>
    ''' Fonction qui permet de retourner une copie de la MAP contenant les FeatureLayers en relation.
    '''</summary>
    ''' 
    '''<return> Copie de la Map contenant les FeatureLayers en relation.</return>
    ''' 
    Function RelationsMap() As IMap

    '''<summary>
    ''' Fonction qui permet de définir les limites de découpage à partir des éléments sélectionner d'un Layer de découpage. 
    '''</summary>
    ''' 
    '''<param name="pFeatureLayerDecoupage">Interface contenant les élements de découpage sélectionnés.</param>
    ''' 
    '''<return>IPolyline contenant la les limites de découpage des éléments sélectionner dans le Layer de découpage, Vide si aucun sélectionné.</return>
    ''' 
    Function DefinirLimiteLayerDecoupage(ByRef pFeatureLayerDecoupage As IFeatureLayer) As IPolyline

    '''<summary>
    '''Fonction qui permet de définir une table contenue dans une géodatabase.
    '''</summary>
    ''' 
    '''<param name="sNomTable"> Nom complet de la table incluant le nom de la géodatabase.</param>
    ''' 
    '''<returns>ITable contenant la table, sinon Nothing.</returns>
    ''' 
    Function DefinirTable(ByVal sNomTable As String,
                          Optional ByVal sNomGeodatabaseDefaut As String = "database connections\BDRS_PRO_BDG_DBA.sde",
                          Optional ByVal sNomProprietaireDefaut As String = "BDG_DBA") As ITable

    '''<summary>
    '''Fonction qui permet d'ouvrir et retourner la Geodatabase à partir d'un nom de Géodatabase.
    '''</summary>
    '''
    '''<param name="sNomGeodatabase"> Nom de la géodatabase à ouvrir.</param>
    ''' 
    '''<returns>IWorkspace contenant la géodatabase.</returns>
    ''' 
    Function DefinirGeodatabase(ByVal sNomGeodatabase As String) As IWorkspace

    '''<summary>
    ''' Fonction qui permet de sélectionner les éléments du FeatureLayer dont la valeur d'attribut respecte ou non l'expression régulière spécifiée.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la requête sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la requête sont conservés dans la sélection.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Function Selectionner(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
#End Region
End Interface
