Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display

'**
'Nom de la composante : clsEquidistanceCourbe.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont l’élévation des courbes de niveau respecte ou non
''' les codes et les équidistances spécifiés.
''' 
''' La classe permet de traiter les trois attributs d'élévation des courbes de niveau CODE, EQUIDISTANCE et TABLE.
''' 
''' CODE : Traite les trois codes d'équidistance soit ceux des courbes maitresse, intermédiaire et intercalaire.
''' 
''' EQUIDISTANCE : Traite les différences entre les valeurs d'élévation des courbes qui intersecte les lignes verticales d'une grille de validation 
'''                selon une équidistance obligatoire et optionnelle.
''' 
''' TABLE : Traite les différences entre les valeurs d'élévation des courbes qui intersecte les lignes verticales d'une grille de validation 
'''         selon des équidistances contenue dans une table d'équidistances.
''' 
''' Note : 
''' Le code est présent dans un attribut de la classe des courbes de niveau.
''' La valeur d'élévation est présente dans un attribut de la classe des courbes de niveau.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 22 juillet 2015
'''</remarks>
''' 
Public Class clsEquidistanceCourbe
    Inherits clsValeurAttribut

    '''<summary>Contient le nom de l'attribut du code.</summary>
    Protected gsAttCode As String = ""
    '''<summary>Contient la position de l'attribut du code.</summary>
    Protected gnPosAttCode As Integer = -1
    '''<summary>Contient le code pour les courbes de niveau Maitresse.</summary>
    Protected gnCodeCourbeMaitresse As Integer = 0
    '''<summary>Contient le code pour les courbes de niveau Intermediaire.</summary>
    Protected gnCodeCourbeIntermediaire As Integer = 0
    '''<summary>Contient le code pour les courbes de niveau Intercalaire.</summary>
    Protected gnCodeCourbeIntercalaire As Integer = 0
    '''<summary>Contient le nom de l'attribut d'élévation.</summary>
    Protected gsAttElevation As String = ""
    '''<summary>Contient la position de l'attribut d'élévation</summary>
    Protected gnPosAttElevation As Integer = -1
    '''<summary>Contient l'équidistance des courbes de niveau Maitresse.</summary>
    Protected gnValeurCourbeMaitresse As Integer = 0
    '''<summary>Contient l'équidistance des courbes de niveau Intermediaire.</summary>
    Protected gnValeurCourbeIntermediaire As Integer = 0
    '''<summary>Contient l'équidistance des courbes de niveau Intercalaire.</summary>
    Protected gnValeurCourbeIntercalaire As Integer = 0
    '''<summary>Contient le nom de l'attribut de découpage.</summary>
    Protected gsAttDecoupage As String = ""
    '''<summary>Contient la position de l'attribut de découpage.</summary>
    Protected gnPosAttDecoupage As Integer = -1
    '''<summary>Contient le nom de la table des équidistances.</summary>
    Protected gsTblEquidistance As String = ""
    '''<summary>Contient la table des équidistances.</summary>
    Protected gpTblEquidistance As ITable = Nothing
    '''<summary>Contient le nom de l'attribut d'identifiant.</summary>
    Protected gsAttIdentifiant As String = ""
    '''<summary>Contient la position de l'attribut d'identifiant.</summary>
    Protected gnPosAttIdentifiant As Integer = -1

    ''' <summary>Structure contenant une équidistance d'un identifiant à valider.</summary>
    Public Structure Structure_Equidistance
        Dim Identifiant As String
        Dim Min As Integer
        Dim Max As Integer
        Dim Equidistance As Integer
        Dim Intercalaire As Integer
    End Structure

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "EQUIDISTANCE"
            Expression = "ELEVATION=10"
            gpFeatureLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    '''</summary>
    ''' 
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sNomAttribut"> Nom de l'attribut à traiter.</param>
    '''<param name="sExpression"> Expression régulière à traiter.</param>
    ''' 
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String, ByVal sExpression As String)
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression
            gpFeatureLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    '''</summary>
    ''' 
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    Public Sub New(ByRef pFeatureLayerSelection As IFeatureLayer, ByVal sParametres As String)
        Try
            'Définir les valeurs par défaut
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres
            gpFeatureLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    '''</summary>
    ''' 
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer, ByVal sParametres As String)
        Try
            'Définir les valeurs par défaut
            gpMap = pMap
            gpFeatureLayerSelection = pFeatureLayerSelection
            gpFeatureLayersRelation = New Collection
            Parametres = sParametres

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        gsAttCode = Nothing
        gnPosAttCode = Nothing
        gnCodeCourbeMaitresse = Nothing
        gnCodeCourbeIntermediaire = Nothing
        gnCodeCourbeIntercalaire = Nothing
        gsAttElevation = Nothing
        gnPosAttElevation = Nothing
        gnValeurCourbeMaitresse = Nothing
        gnValeurCourbeIntermediaire = Nothing
        gnValeurCourbeIntercalaire = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
        'Finaliser
        MyBase.finalize()
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de retourner le nom de la contrainte d'intégrité à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides ReadOnly Property Nom() As String
        Get
            Nom = "EquidistanceCourbe"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la ligne de paramètre à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides Property Parametres() As String
        Get
            'Retourner la valeur des paramètres
            Parametres = NomAttribut & " " & Expression
        End Get
        Set(ByVal value As String)
            'Déclarer les variables de travail
            Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression régulière

            Try
                'Extraire les paramètres
                params = value.Split(CChar(" "))
                'Vérifier si les deux paramètres sont présents
                If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION")

                'Définir les valeurs par défaut
                NomAttribut = params(0)
                Expression = value.Replace(NomAttribut & " ", "")

            Catch ex As Exception
                'Retourner l'erreur
                Throw ex
            Finally
                'Vider la mémoire
                params = Nothing
            End Try
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"

    '''<summary>
    ''' Fonction qui permet de retourner la liste des paramètres possibles.
    '''</summary>
    ''' 
    Public Overloads Overrides Function ListeParametres() As Collection
        'Définir la liste des paramètres par défaut
        ListeParametres = New Collection

        Try
            'Définir un paramètre
            ListeParametres.Add("TABLE ELEVATION DATASET_NAME TBL_EQUIDISTANCE_COURBE IDENTIFIANT")

            'Définir un paramètre
            ListeParametres.Add("EQUIDISTANCE ELEVATION=10")

            'Définir un paramètre
            ListeParametres.Add("EQUIDISTANCE ELEVATION=10,5")

            'Définir un paramètre
            ListeParametres.Add("CODE ELEVATION=10")

            'Définir un paramètre
            ListeParametres.Add("CODE ELEVATION=10,5")

            'Définir un paramètre
            'ListeParametres.Add("CODE ELEVATION=10 CODE_SPEC=1030101")

            'Définir un paramètre
            'ListeParametres.Add("CODE ELEVATION=100,10,5 CODE_SPEC=1030101,1030102,1030103")

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
    Public Overloads Overrides Function EstValide() As Boolean
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
    ''' Routine qui permet d'indiquer si la FeatureClass est valide.
    '''</summary>
    ''' 
    '''<return>Boolean qui indique si la FeatureClass est valide.</return>
    ''' 
    Public Overloads Overrides Function FeatureClassValide() As Boolean
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            FeatureClassValide = False
            gsMessage = "ERREUR : La FeatureClass est invalide."

            'Vérifier si la FeatureClass est valide
            If gpFeatureLayerSelection.FeatureClass IsNot Nothing Then
                'Vérifier si la FeatureClass est de type Polyline
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                    'La contrainte est valide
                    FeatureClassValide = True
                    gsMessage = "La contrainte est valide."
                Else
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polyline."
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si l'attribut est valide.
    '''</summary>
    ''' 
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    Public Overloads Overrides Function AttributValide() As Boolean
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            AttributValide = False
            gsMessage = "ERREUR : L'attribut est invalide : " & gsNomAttribut

            'Vérifier si l'attribut est valide
            If gsNomAttribut = "CODE" Or gsNomAttribut = "EQUIDISTANCE" Or gsNomAttribut = "TABLE" Then
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
    ''' Routine qui permet d'indiquer si l'expression utilisée pour extraire l'élément de découpage est valide.
    '''</summary>
    ''' 
    '''<return>Boolean qui indique si l'expression est valide.</return>
    ''' 
    Public Overloads Overrides Function ExpressionValide() As Boolean
        'Déclarer les variables de travail
        Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression régulière

        Try
            'Définir la valeur par défaut
            ExpressionValide = True
            gsMessage = "La contrainte est valide"
            gnValeurCourbeMaitresse = 0
            gnValeurCourbeIntermediaire = 0
            gnValeurCourbeIntercalaire = 0

            'Vérifier si c'est le traitement de la table des équidistances
            If NomAttribut = "TABLE" Then
                'Extraire le nom d'attribut d'élévation et les valeurs des équidistances
                params = Expression.Split(CChar(" "))
                'Vérifier si l'attribut d'élévation est présent
                If params.Length > 0 Then
                    'Traiter l'attribut d'élévation
                    gsAttElevation = params(0)
                    gnPosAttElevation = gpFeatureLayerSelection.FeatureClass.FindField(gsAttElevation)
                    'Vérifier si l'attribut d'élévation est absent
                    If gnPosAttElevation = -1 Then
                        'Définir l'erreur
                        ExpressionValide = False
                        gsMessage = "ERREUR : L'attribut d'élévation n'existe pas : " & gsAttElevation
                    End If

                    'Vérifier si l'attribut de découpage est présent
                    If params.Length > 1 Then
                        'Vérifier si l'élément de découpage est présent
                        If gpFeatureDecoupage IsNot Nothing Then
                            'Traiter le nom d'attribut d'élévation
                            gsAttDecoupage = params(1)
                            gnPosAttDecoupage = gpFeatureDecoupage.Class.FindField(gsAttDecoupage)
                            'Vérifier si l'attribut de découpage est absent
                            If gnPosAttDecoupage = -1 Then
                                'Définir l'erreur
                                ExpressionValide = False
                                gsMessage = "ERREUR : L'attribut de découpage n'existe pas : " & gsAttDecoupage
                            End If

                            'Vérifier si la table des équidistances est présente
                            If params.Length > 2 Then
                                'Traiter la table d'équidistance
                                gsTblEquidistance = params(2)
                                gpTblEquidistance = DefinirTable(gsTblEquidistance)
                                'Vérifier si l'attribut de découpage est absent
                                If gpTblEquidistance Is Nothing Then
                                    'Définir l'erreur
                                    ExpressionValide = False
                                    gsMessage = "ERREUR : La table des équidistances n'existe pas : " & gsTblEquidistance
                                    'Sortir de la fonction
                                    Exit Function
                                End If

                                'Vérifier si l'attribut d'identifiant est présent
                                If params.Length > 3 Then
                                    'Traiter la table d'équidistance
                                    gsAttIdentifiant = params(3)
                                    gnPosAttIdentifiant = gpTblEquidistance.FindField(gsAttIdentifiant)
                                    'Vérifier si l'attribut de découpage est absent
                                    If gnPosAttIdentifiant = -1 Then
                                        'Définir l'erreur
                                        ExpressionValide = False
                                        gsMessage = "ERREUR : L'attribut d'identifiant n'existe pas dans la table des équidistances : " & gsAttIdentifiant
                                    End If
                                Else
                                    'Définir l'erreur
                                    ExpressionValide = False
                                    gsMessage = "ERREUR : Le nom de l'attribut d'identifiant est absent de la table des équidistances"
                                End If
                            Else
                                'Définir l'erreur
                                ExpressionValide = False
                                gsMessage = "ERREUR : Le nom de la table des équidistances est absent"
                            End If
                        Else
                            'Définir l'erreur
                            ExpressionValide = False
                            gsMessage = "ERREUR : L'élément de découpage est absent"
                        End If
                    Else
                        'Définir l'erreur
                        ExpressionValide = False
                        gsMessage = "ERREUR : Le nom de l'attribut de découpage est absent"
                    End If
                Else
                    'Définir l'erreur
                    ExpressionValide = False
                    gsMessage = "ERREUR : Les paramètres pour utiliser la table des équidistances sont absents"
                End If

                'Si ce n'est pas le traitement de la table des équidistances
            Else
                'Extraire le nom d'attribut d'élévation et les valeurs des équidistances
                params = Expression.Split(CChar("="))
                'Vérifier si le nom de l'attribut d'élévation est présent
                If params.Length > 0 Then
                    'Traiter le nom d'attribut d'élévation
                    gsAttElevation = params(0)
                    gnPosAttElevation = gpFeatureLayerSelection.FeatureClass.FindField(gsAttElevation)
                    'Vérifier si l'attribut d'élévation est présent
                    If gnPosAttElevation >= 0 Then
                        'Vérifier si les valeurs des équidistances sont absentes
                        If params.Length > 1 Then
                            'Extraire le nom d'attribut d'élévation et les valeurs des équidistances
                            params = params(1).Split(CChar(","))
                            'Vérifier si une seule valeur est présente
                            If params.Length = 1 Then
                                'Définir l'équidistance des courbes intermédiaires
                                gnValeurCourbeIntermediaire = CInt(params(0))
                                'Si deux valeurs sont présentes
                            ElseIf params.Length = 2 Then
                                'Définir l'équidistance des courbes intermédiaires
                                gnValeurCourbeIntermediaire = CInt(params(0))
                                'Définir l'équidistance des courbes intermédiaires
                                gnValeurCourbeIntercalaire = CInt(params(1))
                                'Si les trois valeurs sont présentes
                            Else
                                'Définir l'équidistance des courbes intermédiaires
                                gnValeurCourbeMaitresse = CInt(params(0))
                                'Définir l'équidistance des courbes intermédiaires
                                gnValeurCourbeIntermediaire = CInt(params(1))
                                'Définir l'équidistance des courbes intermédiaires
                                gnValeurCourbeIntercalaire = CInt(params(2))
                            End If
                            'Vérifier si c'est le traitement du code
                            If NomAttribut = "CODE" Then
                                'Aucune classe en relation n'est requis et on enlève la validation
                                gpFeatureLayersRelation = Nothing
                            End If
                        Else
                            'Définir l'erreur
                            ExpressionValide = False
                            gsMessage = "ERREUR : Les valeurs des équidistances sont absentes"
                        End If
                    Else
                        'Définir l'erreur
                        ExpressionValide = False
                        gsMessage = "ERREUR : L'attribut d'élévation n'existe pas : " & gsAttElevation
                    End If
                Else
                    'Définir l'erreur
                    ExpressionValide = False
                    gsMessage = "ERREUR : Le nom de l'attribut d'élévation est absent"
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            params = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si le FeatureLayer est valide.
    '''</summary>
    ''' 
    '''<return>Boolean qui indique si le FeatureLayer est valide.</return>
    '''
    Public Overloads Overrides Function FeatureLayersRelationValide() As Boolean
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface pour sélectionner et extraire les éléments d'un FeatureLayer.
        Dim pEnumId As IEnumIDs = Nothing               'Interface qui permet d'extraire les Ids des éléments sélectionnés.
        Dim iOid As Integer = Nothing                   'Contient un Id d'un élément sélectionné.

        'Définir la valeur par défaut, la contrainte est invalide.
        FeatureLayersRelationValide = False
        gsMessage = "ERREUR : Le FeatureLayer en relation est invalide."

        Try
            'Initialiser le FeatureLayer de découpage et l'élément de découpage
            'gpFeatureDecoupage = Nothing
            'gpFeatureLayerDecoupage = Nothing

            'Vérifier si les FeatureLayers en relation sont absent
            If gpFeatureLayersRelation Is Nothing Then
                'La contrainte est valide
                FeatureLayersRelationValide = True
                gsMessage = "La contrainte est valide"

                'Si des FeatureLayers en relation sont présents
            Else
                'Retirer le FeatureLayer en relation correspondant au FeatureLayer de sélection
                If gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then gpFeatureLayersRelation.Remove(gpFeatureLayerSelection.Name)

                'Vérifier si aucun élément en relation 
                If gpFeatureLayersRelation.Count = 0 Then
                    'La contrainte est valide
                    FeatureLayersRelationValide = True
                    gsMessage = "La contrainte est valide"

                    'Vérifier la présence des FeatureLayers en relation
                ElseIf gpFeatureLayersRelation.Count = 1 Then
                    'Traiter tous les FeatureLayers en relation
                    gpFeatureLayerDecoupage = CType(gpFeatureLayersRelation.Item(1), IFeatureLayer)

                    'Vérifier si le FeatureLayer est valide
                    If gpFeatureLayerDecoupage IsNot Nothing Then
                        'Vérifier si la FeatureClass est invalide
                        If gpFeatureLayerDecoupage.FeatureClass IsNot Nothing Then
                            'Vérifier si la FeatureClass est invalide
                            If gpFeatureLayerDecoupage.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                                'Interface pour extraire les éléments sélectionnés
                                pFeatureSel = CType(gpFeatureLayerDecoupage, IFeatureSelection)

                                'Vérifier si aucun élément n'est sélectionné
                                If pFeatureSel.SelectionSet.Count = 0 Then
                                    'Sélectionné l'élément tous les éléments
                                    pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, True)
                                End If

                                'Vérifier si un élément est sélectionné
                                If pFeatureSel.SelectionSet.Count = 1 Then
                                    'Interface pour extrire la liste des Ids
                                    pEnumId = pFeatureSel.SelectionSet.IDs

                                    'Trouver le premier élément en relation
                                    pEnumId.Reset()
                                    iOid = pEnumId.Next

                                    'Vérifier si l'élément est trouvé
                                    If iOid > -1 Then
                                        'Définir l'élément de découpage
                                        gpFeatureDecoupage = gpFeatureLayerDecoupage.FeatureClass.GetFeature(iOid)
                                        'Vider la sélection de découpage
                                        'pFeatureSel.Clear()
                                    End If
                                End If

                                'La contrainte est valide
                                FeatureLayersRelationValide = True
                                gsMessage = "La contrainte est valide"

                            Else
                                'Message d'erreur
                                gsMessage = "ERREUR : La FeatureClass de découpage doit être de type <Polygon>."
                            End If

                        Else
                            'Message d'erreur
                            gsMessage = "ERREUR : La FeatureClass de découpage est invalide."
                        End If

                    Else
                        'Message d'erreur
                        gsMessage = "ERREUR : Le FeatureLayer de découpage est invalide."
                    End If

                Else
                    'Message d'erreur
                    gsMessage = "ERREUR : Un seul FeatureLayer de découpage doit être sélectionné."
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pEnumId = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'identifiant ou la géométrie de découpage sont invalides.
    '''
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Public Overloads Overrides Function Selectionner(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.

        Try
            'Sortir si la contrainte est invalide
            If Me.EstValide() = False Then Err.Raise(1, , Me.Message)

            'Définir la géométrie par défaut
            Selectionner = New GeometryBag
            Selectionner.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Conserver le type de sélection à effectuer
            gbEnleverSelection = bEnleverSelection

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Vérifier si des éléments sont sélectionnés
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

            'Vérifier si on traite les codes des courbes
            If NomAttribut = "CODE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterCodeElevationCourbe(pTrackCancel, bEnleverSelection)

                'Si on traite les équidistances des courbes
            ElseIf NomAttribut = "EQUIDISTANCE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterTopologieCourbe(pTrackCancel, bEnleverSelection)
                'Traiter le FeatureLayer
                'Selectionner = TraiterElevationCourbe(pTrackCancel, bEnleverSelection)
                'Traiter le FeatureLayer
                'Selectionner = TraiterEquidistanceCourbe(pTrackCancel, bEnleverSelection)

                'Si on traite les équidistances des courbes via la tables des équidistances
            ElseIf NomAttribut = "TABLE" Then
                'Traiter le FeatureLayer
                Selectionner = TraiterEquidistanceTable(pTrackCancel, bEnleverSelection)
            End If

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
            pFeatureSel = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non les équidistances d'élévation spécifiées 
    ''' contenues dans une table par identifiant 50K.
    ''' L'équidistance correspond à une distance verticale séparant deux courbes de niveau.
    ''' Le calcul de la distance verticale est obtenue à l'aide d'une grille de lignes verticales.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les intersection entre les éléments sélectionnés et les lignes de la grille de validation qui respecte ou non les équidistances d'élévation spécifiées.</return>
    ''' 
    Private Function TraiterEquidistanceTable(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pFeatureLayerColl As Collection = New Collection 'Contient les FeatureLayer `ajouter dans la topologie.
        Dim pFeatureLayerGrille As IFeatureLayer = Nothing  'Interface contenant le FeatureLayer de la Grille de validation.
        Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les lignes de la grille de validation.
        Dim pGeometryDecoupage As IPolygon = Nothing        'Interface contenant la géométrie de l'élément de découpage.
        Dim pGrilleValidation As IGeometryBag = Nothing     'Interface contenant la grille de validation.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim oIncoherences As Collection = New Collection    'Objet contenant la liste des incohérences trouvées afin de ne pas faire de duplication.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément traité.
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant les géométries en erreur.
        Dim qEquidistances As Collection = Nothing          'Contient une collection de structures d'équidistances à valider.
        Dim sIdentifiant As String = ""                     'contient l'identifiant à traiter.

        'Définir la géométrie par défaut
        pGeometryBag = New GeometryBag
        pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

        Try
            'Interface pour extraire les éléments sélectionnés
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Vérifier si l'élément de découpage est valide
            If gpFeatureDecoupage Is Nothing Then
                'Retourner une erreur de traitement
                Throw New Exception("ERREUR : Aucun élément de découpage n'est spécifié")
            End If

            'Définir la géométrie de découpage
            pGeometryDecoupage = CType(gpFeatureDecoupage.ShapeCopy, IPolygon)
            pGeometryDecoupage.Project(pGeometryBag.SpatialReference)

            'Afficher le message
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la grille de validation ..."
            'Créer une grille de validation vide
            pGrilleValidation = New GeometryBag
            pGrilleValidation.SpatialReference = pGeometryDecoupage.SpatialReference
            'Vérifier comment on doit créer les lignes de la grile pour être plus rapide
            If pGeometryDecoupage.Envelope.Height > pGeometryDecoupage.Envelope.Width Then
                'Créer les lignes verticales dans la grille de validation pour chaque élément traité
                Call CreerLigneVerticaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            Else
                'Ajouter les lignes horizontales dans la grille de validation pour chaque élément traité
                Call CreerLigneHorizontaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            End If
            'Créer les lignes de la grille de validation dans un GeometryBag
            pGeomRelColl = CType(pGrilleValidation, IGeometryCollection)

            'Afficher le message
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer le FeatureLayer de la grille ..."
            'Créer le FeatureLayer pour la grille de validation
            pFeatureLayerGrille = CreerFeatureLayerGrilleValidation(pGrilleValidation)

            'Afficher le message
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & pGeomRelColl.GeometryCount.ToString & " Lignes) ..."
            'Interface pour extraire la tolérance de précision de la référence spatiale
            pSRTolerance = CType(pGeometryDecoupage.SpatialReference, ISpatialReferenceTolerance)
            'Ajouter les FeatureLayer de la topologie
            pFeatureLayerColl.Add(gpFeatureLayerSelection)
            pFeatureLayerColl.Add(pFeatureLayerGrille)
            'Création de la topologie selon la zone à traiter
            pTopologyGraph = CreerTopologyGraph(pGeometryDecoupage.Envelope, pFeatureLayerColl, pSRTolerance.XYTolerance)

            'Extraire l'identifiant à traiter
            sIdentifiant = gpFeatureDecoupage.Value(gnPosAttDecoupage).ToString
            'Extraire les équidistances de la table selon l'identifiant traité
            qEquidistances = DefinirEquidistancesTable(sIdentifiant)

            'Conserver la sélection de départ
            pFeatureSel.Clear()
            pSelectionSet = pFeatureSel.SelectionSet

            'Sélectionner tous les éléments de la grille
            pFeatureSel = CType(pFeatureLayerGrille, IFeatureSelection)
            pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)

            'Afficher le message de validation des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Valider les élévations des courbes selon la grille : " + " (" + pTopologyGraph.Edges.Count.ToString + " Edges / " & sIdentifiant & " : " & qEquidistances.Count.ToString & " équidistances) ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomRelColl.GeometryCount, pTrackCancel)
            'Interfaces pour extraire les lignes de la grille de validation
            pFeatureSel.SelectionSet.Search(Nothing, True, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire la première ligne de la grille de validation
            pFeature = pFeatureCursor.NextFeature()

            'Traiter toutes les lignes de la grille de validation
            Do While Not pFeature Is Nothing
                'Traiter les équidistances pour une ligne de la grille en utilisant les équidistances d'un identifiant contenues dans la table
                Call TraiterEquidistanceTableElement(qEquidistances, pTopologyGraph, pFeatureLayerGrille, pFeature, bEnleverSelection,
                                                     pSelectionSet, oIncoherences, pGeometryBag)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Ajouter les éléments sélectionneés dans le FeatureLayer
            pFeatureSel.SelectionSet = pSelectionSet

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

            'Retrourner les géométries en erreur
            TraiterEquidistanceTable = pGeometryBag

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerColl = Nothing
            pFeatureLayerGrille = Nothing
            pSRTolerance = Nothing
            pGeomRelColl = Nothing
            pGeometryDecoupage = Nothing
            pSRTolerance = Nothing
            pTopologyGraph = Nothing
            pGrilleValidation = Nothing
            oIncoherences = Nothing
            pGeometryBag = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            qEquidistances = Nothing
            GC.Collect()
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de retourner une collection de structures d'équidistances à valider contenues dans une table d'équidistances.
    ''' 
    ''' La table contient entre autre les attributs suivants :
    '''   IDENTIFIANT  : Identifiant de découpage.
    '''   ELEVATION_MIN: Élévation minimum de l'équidistance.
    '''   ELEVATION_MAX: Élévation maximum de l'équidistance.
    '''   EQUIDISTANCES: Équidistance à valider.
    '''</summary>
    ''' 
    '''<param name="sIdentifiant"> Contient l'identifiant de découpage à valider.</param>
    ''' 
    Private Function DefinirEquidistancesTable(ByVal sIdentifiant As String) As Collection
        'Déclarer les variables de travail
        Dim qEquidistance As Structure_Equidistance = Nothing    'Contient une équidistance à valider.
        Dim pQueryFilter As IQueryFilter = Nothing
        Dim pCursor As ICursor = Nothing
        Dim pRow As IRow = Nothing

        'Par défaut, la collection des structures d'équidistances à valider est vide
        DefinirEquidistancesTable = New Collection

        Try
            'Interface pour définir la requête pour extraire les éléments du découpage
            pQueryFilter = New QueryFilter
            'Définir la requête pour extraire les éléments du découpage
            pQueryFilter.WhereClause = gsAttIdentifiant & "='" & sIdentifiant & "'"

            'Extraire les équidistances à traiter
            pCursor = gpTblEquidistance.Search(pQueryFilter, False)

            'Extraire la première équidistance
            pRow = pCursor.NextRow

            'Traiter toutes les équidistances
            Do Until pRow Is Nothing
                'Créer une nouvelle equidistance
                qEquidistance = New Structure_Equidistance

                'Définir l'identifiant à valider
                qEquidistance.Identifiant = sIdentifiant
                'Définir l'élévation minimum à valider
                qEquidistance.Min = CInt(pRow.Value(gnPosAttIdentifiant + 1))
                'Définir l'élévation maximum à valider
                qEquidistance.Max = CInt(pRow.Value(gnPosAttIdentifiant + 2))
                'Définir l'équidistance à valider
                qEquidistance.Equidistance = CInt(pRow.Value(gnPosAttIdentifiant + 3))
                'Vérifier la présence d'une équidistance intercalaire
                If pRow.Value(gnPosAttIdentifiant + 4) IsNot DBNull.Value Then
                    'Définir l'équidistance intercalaire à valider
                    qEquidistance.Intercalaire = CInt(pRow.Value(gnPosAttIdentifiant + 4))
                End If

                'Conserver l'équidistance dans la collection
                DefinirEquidistancesTable.Add(qEquidistance)

                'Extraire la prochaine équidistance
                pRow = pCursor.NextRow
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            qEquidistance = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de valider les courbes de niveau à partir d'une ligne de la grille de validation en utilisant les équidistances contenues dans une table.
    '''</summary>
    ''' 
    '''<param name="qEquidistances"> Contient la collection des structures d'équidistances de la table à valider.</param>
    '''<param name="pTopologyGraph"> Contient la topologie entre les courbes et le FeatureLayer de la grille de validation.</param>
    '''<param name="pFeatureLayerGrille"> Contient les éléments des lignes de la grille de validation.</param>
    '''<param name="pFeature"> Contient l'élément de la ligne de la grille de validation.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''<param name="pSelectionSet"> Contient les éléments sélectionnés en erreur.</param>
    '''<param name="oIncoherences"> Contient les incohérences d'erreurs.</param>
    '''<param name="pGeometryBag"> Contient les géométries en erreur.</param>
    '''
    Private Sub TraiterEquidistanceTableElement(ByVal qEquidistances As Collection, ByVal pTopologyGraph As ITopologyGraph4, ByVal pFeatureLayerGrille As IFeatureLayer,
                                                ByVal pFeature As IFeature, ByVal bEnleverSelection As Boolean,
                                                ByRef pSelectionSet As ISelectionSet, ByRef oIncoherences As Collection, ByRef pGeometryBag As IGeometryBag)
        'Déclarer les variables de travail
        Dim qEquidistance As Structure_Equidistance = Nothing    'Contient une équidistance à valider.
        Dim pGeomResColl As IGeometryCollection = Nothing       'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeometry As IGeometry = Nothing                    'Interface contenant la géométrie de l'élément traité.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing        'Interface contenant tous les Edges à traiter.
        Dim pTopoEdge As ITopologyEdge = Nothing                'Interface contenant un Edge de la topologie.
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing    'Interface utilisé pour extraire les éléments du Edge.
        Dim pEsriTopoParent As esriTopologyParent = Nothing     'Interface contenant la structure d'information du parent sélectionné.
        Dim pFeatureFrom As IFeature = Nothing                  'Interface contenant l'élément du premier node du edge.
        Dim pFeatureTo As IFeature = Nothing                    'Interface contenant l'élément du dernier node du edge.
        Dim nDifference As Integer = Nothing                    'Contient la différence de valeur d'élévation
        Dim sOid As String = Nothing                            'Contient la valeur du OID de l'élément
        Dim sMessage As String = ""                             'Contient le message à écrire.
        Dim bSucces As Boolean = False                          'Indique si l'identifiant et la géométrie respecte l'élément de découpage.
        Dim bValide As Boolean = False                          'Indique si la validation de l'équidistance a été effectué

        Try
            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(pGeometryBag, IGeometryCollection)

            'Interface pour extraire tous les edges d'une ligne de la grille de validation
            pEnumTopoEdge = pTopologyGraph.GetParentEdges(pFeatureLayerGrille.FeatureClass, pFeature.OID)

            'Initialisation de l'extraction
            pEnumTopoEdge.Reset()

            'Extraire le premier edge de la ligne
            pTopoEdge = pEnumTopoEdge.Next()

            'Traiter tous les edges de la ligne
            Do Until pTopoEdge Is Nothing
                'Intialisation des éléments de courbes
                pFeatureFrom = Nothing
                pFeatureTo = Nothing

                'Contient les parents sélectionnés
                pEnumTopoParent = pTopoEdge.FromNode.Parents
                'Initialiser l'extraction
                pEnumTopoParent.Reset()
                'Extraire le premier élément
                pEsriTopoParent = pEnumTopoParent.Next()
                'Traiter tous les élément
                Do Until pEsriTopoParent.m_pFC Is Nothing
                    'Vérifier si la FeatureClass est celle des courbes
                    If pEsriTopoParent.m_pFC.AliasName = gpFeatureLayerSelection.FeatureClass.AliasName Then
                        'Extraire l'élément
                        pFeatureFrom = pEsriTopoParent.m_pFC.GetFeature(pEsriTopoParent.m_FID)
                        'sortir de la boucle
                        Exit Do
                    End If
                    'Extraire le prochain élément
                    pEsriTopoParent = pEnumTopoParent.Next()
                Loop

                'Contient les parents sélectionnés
                pEnumTopoParent = pTopoEdge.ToNode.Parents
                'Initialiser l'extraction
                pEnumTopoParent.Reset()
                'Extraire le premier élément
                pEsriTopoParent = pEnumTopoParent.Next()
                'Traiter tous les élément
                Do Until pEsriTopoParent.m_pFC Is Nothing
                    'Vérifier si la FeatureClass est celle des courbes
                    If pEsriTopoParent.m_pFC.AliasName = gpFeatureLayerSelection.FeatureClass.AliasName Then
                        'Extraire l'élément
                        pFeatureTo = pEsriTopoParent.m_pFC.GetFeature(pEsriTopoParent.m_FID)
                        'sortir de la boucle
                        Exit Do
                    End If
                    'Extraire le prochain élément
                    pEsriTopoParent = pEnumTopoParent.Next()
                Loop

                'Vérifier si les deux éléments des courbes sont présents
                If pFeatureFrom IsNot Nothing And pFeatureTo IsNot Nothing Then
                    'Vérifier si les OIDs sont différents
                    If pFeatureFrom.OID <> pFeatureTo.OID Then
                        'Calcul de la différence entre deux élévations consécutive
                        nDifference = CInt(Math.Abs(CInt(pFeatureFrom.Value(gnPosAttElevation)) - CInt(pFeatureTo.Value(gnPosAttElevation))))

                        'Initialiser la validation
                        bValide = False

                        'Valider toutes les équidistances
                        For Each qEquidistance In qEquidistances
                            'Vérifier si l'équidistance à valider est la bonne
                            If CInt(pFeatureFrom.Value(gnPosAttElevation)) >= qEquidistance.Min And CInt(pFeatureFrom.Value(gnPosAttElevation)) <= qEquidistance.Max _
                            And CInt(pFeatureTo.Value(gnPosAttElevation)) >= qEquidistance.Min And CInt(pFeatureTo.Value(gnPosAttElevation)) <= qEquidistance.Max Then
                                'Indiquer que la validation de l'équidistance a été effectué
                                bValide = True

                                'Vérifier la différence d'élévation entre deux sommets consécutifs
                                If nDifference <> 0 And nDifference <> qEquidistance.Equidistance And nDifference <> qEquidistance.Intercalaire Then
                                    'Définir le résultat
                                    sMessage = "#Équidistance invalide"
                                    bSucces = False
                                Else
                                    'Définir le résultat
                                    sMessage = "#Équidistance valide"
                                    bSucces = True
                                End If

                                'Vérifier si on doit sélectionner l'élément
                                If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                    'Définir la clé de l'incohérence contenant les deux OIDs d'éléments en erreur
                                    If pFeatureFrom.OID < pFeatureTo.OID Then
                                        sOid = pFeatureFrom.OID & "," & pFeatureTo.OID
                                    Else
                                        sOid = pFeatureTo.OID & "," & pFeatureFrom.OID
                                    End If

                                    'Vérifier si l'incohérence existe déja
                                    If Not oIncoherences.Contains(sOid) Then
                                        'Ajouter la note de l'incohérence
                                        sMessage = "OID=" & sOid & sMessage _
                                                & " /Élévation=" & CInt(pFeatureFrom.Value(gnPosAttElevation)).ToString & "," & CInt(pFeatureTo.Value(gnPosAttElevation)).ToString _
                                                & " /Min,Max,Equidistance,Intercalaire=" & qEquidistance.Min.ToString & "," & qEquidistance.Max.ToString & "," & qEquidistance.Equidistance.ToString & "," & qEquidistance.Intercalaire.ToString _
                                                & " /Identifiant=" & qEquidistance.Identifiant & "/Précision=" & gdPrecision.ToString("F3")

                                        'Ajouter l'incohérence
                                        oIncoherences.Add(sMessage, sOid)

                                        'Ajouter les éléments dans la sélection
                                        pSelectionSet.Add(pFeatureFrom.OID)
                                        pSelectionSet.Add(pFeatureTo.OID)

                                        'Créer une nouvelle géométrie de type ligne
                                        pGeometry = pTopoEdge.Geometry
                                        'pGeometry.Project(pGeometryDecoupage.SpatialReference)

                                        'Ajouter l'enveloppe de l'élément sélectionné
                                        pGeomResColl.AddGeometry(pGeometry)

                                        'Écrire une erreur
                                        EcrireFeatureErreur(sMessage, pGeometry, CSng(pFeatureTo.Value(gnPosAttElevation)))
                                    End If
                                End If

                                'Sortir de la boucle des équidistances
                                Exit For
                            End If
                        Next

                        'Vérifier si l'équidistance a été validée
                        If Not bValide Then
                            'Définir la clé de l'incohérence contenant les deux OIDs d'éléments en erreur
                            If pFeatureFrom.OID < pFeatureTo.OID Then
                                sOid = pFeatureFrom.OID & "," & pFeatureTo.OID
                            Else
                                sOid = pFeatureTo.OID & "," & pFeatureFrom.OID
                            End If

                            'Vérifier si l'incohérence existe déja
                            If Not oIncoherences.Contains(sOid) Then
                                'Définir l'erreur
                                sMessage = "#Élévation exclut des équidistances"
                                'Ajouter la note de l'incohérence
                                sMessage = "OID=" & sOid & sMessage _
                                        & " /Élévation=" & CInt(pFeatureFrom.Value(gnPosAttElevation)).ToString & "," & CInt(pFeatureTo.Value(gnPosAttElevation)).ToString _
                                        & " /Identifiant=" & qEquidistance.Identifiant & "/Précision=" & gdPrecision.ToString("F3")

                                'Ajouter l'incohérence
                                oIncoherences.Add(sMessage, sOid)

                                'Ajouter les éléments dans la sélection
                                pSelectionSet.Add(pFeatureFrom.OID)
                                pSelectionSet.Add(pFeatureTo.OID)

                                'Créer une nouvelle géométrie de type ligne
                                pGeometry = pTopoEdge.Geometry
                                'pGeometry.Project(pGeometryDecoupage.SpatialReference)

                                'Ajouter l'enveloppe de l'élément sélectionné
                                pGeomResColl.AddGeometry(pGeometry)

                                'Écrire une erreur
                                EcrireFeatureErreur(sMessage, pGeometry, CSng(pFeatureTo.Value(gnPosAttElevation)))
                            End If
                        End If
                    End If
                End If

                'Extraire le premier edge de la ligne
                pTopoEdge = pEnumTopoEdge.Next()
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            qEquidistance = Nothing
            pGeomResColl = Nothing
            pGeometry = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pEnumTopoParent = Nothing
            pEsriTopoParent = Nothing
            pFeatureFrom = Nothing
            pFeatureTo = Nothing
            GC.Collect()
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non les équidistances d'élévation spécifiées.
    ''' L'équidistance correspond à une distance verticale séparant deux courbes de niveau.
    ''' Le calcul de la distance verticale est obtenue à l'aide d'une grille de lignes verticales.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les intersection entre les éléments sélectionnés et les lignes de la grille de validation qui respecte ou non les équidistances d'élévation spécifiées.</return>
    ''' 
    Private Function TraiterTopologieCourbe(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pFeatureLayerColl As Collection = New Collection 'Contient les FeatureLayer `ajouter dans la topologie.
        Dim pFeatureLayerGrille As IFeatureLayer = Nothing  'Interface contenant le FeatureLayer de la Grille de validation.
        Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les lignes de la grille de validation.
        Dim pGeometryDecoupage As IPolygon = Nothing        'Interface contenant la géométrie de l'élément de découpage.
        Dim pGrilleValidation As IGeometryBag = Nothing     'Interface contenant la grille de validation.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim oIncoherences As Collection = New Collection    'Objet contenant la liste des incohérences trouvées afin de ne pas faire de duplication.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément traité.
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant les géométries en erreur.

        'Définir la géométrie par défaut
        pGeometryBag = New GeometryBag
        pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

        Try
            'Interface pour extraire les éléments sélectionnés
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Vérifier si l'élément de découpage est valide
            If gpFeatureDecoupage Is Nothing Then
                'Créer la géométrie de découpage à partir de l'enveloppe du FeatureLayer de sélection
                'pGeometryDecoupage = EnvelopeToPolygon(gpFeatureLayerSelection.AreaOfInterest)
                pGeometryDecoupage = EnvelopeToPolygon(EnveloppeSelectionSet(pFeatureSel.SelectionSet, pGeometryBag.SpatialReference))
            Else
                'Définir la géométrie de découpage
                pGeometryDecoupage = CType(gpFeatureDecoupage.ShapeCopy, IPolygon)
                pGeometryDecoupage.Project(pGeometryBag.SpatialReference)
            End If

            'Afficher le message
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la grille de validation ..."
            'Créer une grille de validation vide
            pGrilleValidation = New GeometryBag
            pGrilleValidation.SpatialReference = pGeometryDecoupage.SpatialReference
            'Vérifier comment on doit créer les lignes de la grile pour être plus rapide
            If pGeometryDecoupage.Envelope.Height > pGeometryDecoupage.Envelope.Width Then
                'Créer les lignes verticales dans la grille de validation pour chaque élément traité
                Call CreerLigneVerticaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            Else
                'Ajouter les lignes horizontales dans la grille de validation pour chaque élément traité
                Call CreerLigneHorizontaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            End If
            'Créer les lignes de la grille de validation dans un GeometryBag
            pGeomRelColl = CType(pGrilleValidation, IGeometryCollection)

            'Afficher le message
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer le FeatureLayer de la grille ..."
            'Créer le FeatureLayer pour la grille de validation
            pFeatureLayerGrille = CreerFeatureLayerGrilleValidation(pGrilleValidation)

            'Afficher le message
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & pGeomRelColl.GeometryCount.ToString & " Lignes) ..."
            'Interface pour extraire la tolérance de précision de la référence spatiale
            pSRTolerance = CType(pGeometryDecoupage.SpatialReference, ISpatialReferenceTolerance)
            'Ajouter les FeatureLayer de la topologie
            pFeatureLayerColl.Add(gpFeatureLayerSelection)
            pFeatureLayerColl.Add(pFeatureLayerGrille)
            'Création de la topologie selon la zone à traiter
            pTopologyGraph = CreerTopologyGraph(pGeometryDecoupage.Envelope, pFeatureLayerColl, pSRTolerance.XYTolerance)

            'Conserver la sélection de départ
            pFeatureSel.Clear()
            pSelectionSet = pFeatureSel.SelectionSet

            'Sélectionner tous les éléments de la grille
            pFeatureSel = CType(pFeatureLayerGrille, IFeatureSelection)
            pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)

            'Afficher le message de validation des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Valider les élévations des courbes selon la grille : " + " (" + pTopologyGraph.Edges.Count.ToString + " Edges) ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomRelColl.GeometryCount, pTrackCancel)
            'Interfaces pour extraire les lignes de la grille de validation
            pFeatureSel.SelectionSet.Search(Nothing, True, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire la première ligne de la grille de validation
            pFeature = pFeatureCursor.NextFeature()

            'Traiter toutes les lignes de la grille de validation
            Do While Not pFeature Is Nothing
                'Traiter les équidistances pour une ligne de la grille
                Call TraiterTopologieCourbeElement(pTopologyGraph, pFeatureLayerGrille, pFeature, bEnleverSelection, pSelectionSet, oIncoherences, pGeometryBag)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Ajouter les éléments sélectionneés dans le FeatureLayer
            pFeatureSel.SelectionSet = pSelectionSet

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

            'Retrourner les géométries en erreur
            TraiterTopologieCourbe = pGeometryBag

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerColl = Nothing
            pFeatureLayerGrille = Nothing
            pSRTolerance = Nothing
            pGeomRelColl = Nothing
            pGeometryDecoupage = Nothing
            pSRTolerance = Nothing
            pTopologyGraph = Nothing
            pGrilleValidation = Nothing
            oIncoherences = Nothing
            pGeometryBag = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            GC.Collect()
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de valider les courbes de niveau à partir d'une ligne de la grille de validation.
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie entre les courbes et le FeatureLayer de la grille de validation.</param>
    '''<param name="pFeatureLayerGrille"> Contient les éléments des lignes de la grille de validation.</param>
    '''<param name="pFeature"> Contient l'élément de la ligne de la grille de validation.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''<param name="pSelectionSet"> Contient les éléments sélectionnés en erreur.</param>
    '''<param name="oIncoherences"> Contient les incohérences d'erreurs.</param>
    '''<param name="pGeometryBag"> Contient les géométries en erreur.</param>
    '''
    Private Sub TraiterTopologieCourbeElement(ByVal pTopologyGraph As ITopologyGraph4, ByVal pFeatureLayerGrille As IFeatureLayer,
                                              ByVal pFeature As IFeature, ByVal bEnleverSelection As Boolean,
                                              ByRef pSelectionSet As ISelectionSet, ByRef oIncoherences As Collection, ByRef pGeometryBag As IGeometryBag)
        'Déclarer les variables de travail
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément traité.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface contenant tous les Edges à traiter.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie.
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface utilisé pour extraire les éléments du Edge.
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.
        Dim pFeatureFrom As IFeature = Nothing              'Interface contenant l'élément du premier node du edge.
        Dim pFeatureTo As IFeature = Nothing                'Interface contenant l'élément du dernier node du edge.
        Dim nDifference As Integer = Nothing                'Contient la différence de valeur d'élévation
        Dim sOid As String = Nothing                        'Contient la valeur du OID de l'élément
        Dim sMessage As String = ""                         'Contient le message à écrire.
        Dim bSucces As Boolean = False                      'Indique si l'identifiant et la géométrie respecte l'élément de découpage.

        Try
            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(pGeometryBag, IGeometryCollection)

            'Interface pour extraire tous les edges d'une ligne de la grille de validation
            pEnumTopoEdge = pTopologyGraph.GetParentEdges(pFeatureLayerGrille.FeatureClass, pFeature.OID)

            'Initialisation de l'extraction
            pEnumTopoEdge.Reset()

            'Extraire le premier edge de la ligne
            pTopoEdge = pEnumTopoEdge.Next()

            'Traiter tous les edges de la ligne
            Do Until pTopoEdge Is Nothing
                'Intialisation des éléments de courbes
                pFeatureFrom = Nothing
                pFeatureTo = Nothing

                'Contient les parents sélectionnés
                pEnumTopoParent = pTopoEdge.FromNode.Parents
                'Initialiser l'extraction
                pEnumTopoParent.Reset()
                'Extraire le premier élément
                pEsriTopoParent = pEnumTopoParent.Next()
                'Traiter tous les élément
                Do Until pEsriTopoParent.m_pFC Is Nothing
                    'Vérifier si la FeatureClass est celle des courbes
                    If pEsriTopoParent.m_pFC.AliasName = gpFeatureLayerSelection.FeatureClass.AliasName Then
                        'Extraire l'élément
                        pFeatureFrom = pEsriTopoParent.m_pFC.GetFeature(pEsriTopoParent.m_FID)
                        'sortir de la boucle
                        Exit Do
                    End If
                    'Extraire le prochain élément
                    pEsriTopoParent = pEnumTopoParent.Next()
                Loop

                'Contient les parents sélectionnés
                pEnumTopoParent = pTopoEdge.ToNode.Parents
                'Initialiser l'extraction
                pEnumTopoParent.Reset()
                'Extraire le premier élément
                pEsriTopoParent = pEnumTopoParent.Next()
                'Traiter tous les élément
                Do Until pEsriTopoParent.m_pFC Is Nothing
                    'Vérifier si la FeatureClass est celle des courbes
                    If pEsriTopoParent.m_pFC.AliasName = gpFeatureLayerSelection.FeatureClass.AliasName Then
                        'Extraire l'élément
                        pFeatureTo = pEsriTopoParent.m_pFC.GetFeature(pEsriTopoParent.m_FID)
                        'sortir de la boucle
                        Exit Do
                    End If
                    'Extraire le prochain élément
                    pEsriTopoParent = pEnumTopoParent.Next()
                Loop

                'Vérifier si les deux éléments des courbes sont présents
                If pFeatureFrom IsNot Nothing And pFeatureTo IsNot Nothing Then
                    'Vérifier si les OIDs sont différents
                    If pFeatureFrom.OID <> pFeatureTo.OID Then
                        'Calcul de la différence entre deux élévations consécutive
                        nDifference = CInt(Math.Abs(CInt(pFeatureFrom.Value(gnPosAttElevation)) - CInt(pFeatureTo.Value(gnPosAttElevation))))

                        'Vérifier la différence d'élévation entre deux sommets consécutifs
                        If nDifference <> 0 And nDifference <> gnValeurCourbeIntermediaire And nDifference <> gnValeurCourbeIntercalaire Then
                            'Définir le résultat
                            sMessage = "#Équidistance invalide"
                            bSucces = False
                        Else
                            'Définir le résultat
                            sMessage = "#Équidistance valide"
                            bSucces = True
                        End If

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Définir la clé de l'incohérence contenant les deux OIDs d'éléments en erreur
                            If pFeatureFrom.OID < pFeatureTo.OID Then
                                sOid = pFeatureFrom.OID & "," & pFeatureTo.OID
                            Else
                                sOid = pFeatureTo.OID & "," & pFeatureFrom.OID
                            End If

                            'Vérifier si l'incohérence existe déja
                            If Not oIncoherences.Contains(sOid) Then
                                'Ajouter la note de l'incohérence
                                sMessage = "OID=" & sOid & sMessage _
                                        & " /Élévation:" & CInt(pFeatureFrom.Value(gnPosAttElevation)).ToString & "," & CInt(pFeatureTo.Value(gnPosAttElevation)).ToString _
                                        & " /" & Expression & " /Précision=" & gdPrecision.ToString("F3")

                                'Ajouter l'incohérence
                                oIncoherences.Add(sMessage, sOid)

                                'Ajouter les éléments dans la sélection
                                pSelectionSet.Add(pFeatureFrom.OID)
                                pSelectionSet.Add(pFeatureTo.OID)

                                'Créer une nouvelle géométrie de type ligne
                                pGeometry = pTopoEdge.Geometry
                                'pGeometry.Project(pGeometryDecoupage.SpatialReference)

                                'Ajouter l'enveloppe de l'élément sélectionné
                                pGeomResColl.AddGeometry(pGeometry)

                                'Écrire une erreur
                                EcrireFeatureErreur(sMessage, pGeometry, CSng(pFeatureTo.Value(gnPosAttElevation)))
                            End If
                        End If
                    End If
                End If

                'Extraire le premier edge de la ligne
                pTopoEdge = pEnumTopoEdge.Next()
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeometry = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pEnumTopoParent = Nothing
            pEsriTopoParent = Nothing
            pFeatureFrom = Nothing
            pFeatureTo = Nothing
            GC.Collect()
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non les codes et les valeurs d'élévation spécifiés.
    ''' Il y a trois types de courbes: 
    ''' -les courbes maitresses sont optionnelles et la valeur d'élévation est supérieure et est un multiple de celle des courbes intermédiaire.
    ''' -les courbes intermédiaires sont obligatoires et la valeur d'élévation est supérieure et est un multiple de celle des courbes intercalaires.
    ''' -les courbes intercalaires sont optionnelles.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    '''
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés qui respecte ou non les codes et les valeurs d'élévation spécifiés.</return>
    ''' 
    Private Function TraiterCodeElevationCourbe(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                        'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing          'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeature As IFeature = Nothing                      'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing       'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                    'Interface contenant la géométrie de l'élément traité.
        Dim sMessage As String = ""                             'Contient le message à écrire.
        Dim bSucces As Boolean = False                          'Indique si l'identifiant et la géométrie respecte l'élément de découpage.

        Try
            'Définir la géométrie par défaut
            TraiterCodeElevationCourbe = New GeometryBag
            TraiterCodeElevationCourbe.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer
            pSelectionSet = pFeatureSel.SelectionSet

            'Enlever la sélection
            pFeatureSel.Clear()

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterCodeElevationCourbe, IGeometryCollection)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Initialisation du traitement
                bSucces = False

                'Interface pour projeter
                pGeometry = pFeature.ShapeCopy
                pGeometry.Project(TraiterCodeElevationCourbe.SpatialReference)

                'Valider les valeurs d'élévation et les codes des courbes de niveau
                sMessage = ValiderValeurEquidistanceFeature(pFeature, _
                gnPosAttCode, gnCodeCourbeMaitresse, gnCodeCourbeIntermediaire, gnCodeCourbeIntercalaire, _
                gnPosAttElevation, gnValeurCourbeMaitresse, gnValeurCourbeIntermediaire, gnValeurCourbeIntercalaire)

                'Vérifier si aucune erreur trouvée
                If sMessage = "" Then
                    'Définir le résultat
                    sMessage = "Courbe de niveau valide"
                    bSucces = True
                End If

                'Vérifier si on doit sélectionner l'élément
                If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                    'Ajouter l'élément dans la sélection
                    pFeatureSel.Add(pFeature)
                    'Ajouter l'enveloppe de l'élément sélectionné
                    pGeomSelColl.AddGeometry(pGeometry)
                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & sMessage & "/" & Expression, pGeometry, CSng(pFeature.Value(gnPosAttElevation)))
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

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
            pGeomSelColl = Nothing
            pGeometry = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de valider les valeurs d'élévation des courbes de niveau qui respectent les multiples des différents
    ''' intervalles des courbes de niveau.
    '''</summary>
    '''
    '''<param name="pFeature">Interface contenant l'élément à valider.</param>
    '''<param name="nPosAttCode">Position du nom de l'attribut du code des courbes dans la structure d'attributs.</param>
    '''<param name="nCodeCourbeMaitresse">Code de l'équidistance des courbes maîtresses.</param>
    '''<param name="nCodeCourbeIntermediaire">Code de l'équidistance des courbes Intermédiaire.</param>
    '''<param name="nCodeCourbeIntercalaire">Code de l'équidistance des courbes Intercalaire.</param>
    '''<param name="nPosAttElevation">Position du nom de l'attribut d'élévation des courbes dans la structure d'attributs.</param>
    '''<param name="nValeurCourbeMaitresse">Valeur de l'équidistance des courbes maîtresses.</param>
    '''<param name="nValeurCourbeIntermediaire">Valeur de l'équidistance des courbes Intermédiaire.</param>
    '''<param name="nValeurCourbeIntercalaire">Valeur de l'équidistance des courbes Intercalaire.</param>
    '''
    Private Function ValiderValeurEquidistanceFeature(ByVal pFeature As IFeature, _
    ByVal nPosAttCode As Integer, ByVal nCodeCourbeMaitresse As Integer, ByVal nCodeCourbeIntermediaire As Integer, _
    ByVal nCodeCourbeIntercalaire As Integer, ByVal nPosAttElevation As Integer, ByVal nValeurCourbeMaitresse As Integer, _
    ByVal nValeurCourbeIntermediaire As Integer, ByVal nValeurCourbeIntercalaire As Integer) As String
        'Déclarer les variables de travail
        Dim nEquidistance As Integer = Nothing              'Contient l'intervale d'équidistance à valider
        Dim nNouvelleValeur As Integer = Nothing            'Contient la nouvelle valeur d'élévation corrigée
        Dim dDiffMaitresse As Double = 1000000000           'Contient la différence du multiple des courbes maîtresse
        Dim dDiffIntermediaire As Double = 1000000000       'Contient la différence du multiple des courbes intermédiaire
        Dim dDiffIntercalaire As Double = 1000000000        'Contient la différence du multiple des courbes intercalaire
        Dim dDifference As Double = Nothing                 'Contient la différence du multiple des courbes
        Dim sNote As String = ""                            'Contient la note de l'incohérence

        'Définir le message d'erreur par défaut
        ValiderValeurEquidistanceFeature = ""

        Try
            'Vérifier si on doit valider l'équidistance maîtresse
            If nValeurCourbeMaitresse > 0 Then
                'Calcul du modulo entre la valeur de l'équidistance de l'élément et l'équidistance maîtresse
                dDiffMaitresse = CInt(pFeature.Value(nPosAttElevation)) Mod nValeurCourbeMaitresse

                'Vérifier si l'élévation correspond à une courbe maîtresse
                If dDiffMaitresse = 0 Then
                    'Vérifier si le code de la courbe maîtresse est spécifié
                    If nCodeCourbeMaitresse > 0 Then
                        'Vérifier si le code de l'élément est invalide
                        If CInt(pFeature.Value(nPosAttCode)) <> nCodeCourbeMaitresse Then
                            'Ajouter la note de l'incohérence pour le code de courbe maîtresse
                            sNote = "Code de la courbe maîtresse:" & nValeurCourbeMaitresse _
                            & ", Attribut:" & pFeature.Fields.Field(nPosAttCode).Name _
                            & ", Ancienne valeur:" & CInt(pFeature.Value(nPosAttCode)) _
                            & ", Nouvelle valeur:" & nCodeCourbeMaitresse
                        End If
                    End If

                    'Sortir de la validation
                    Return sNote
                End If

                'L'équidistance maîtresse est en erreur par défaut
                sNote = "Équidistance maîtresse:"
                nEquidistance = nValeurCourbeMaitresse
                dDifference = dDiffMaitresse
            End If

            'Vérifier si on doit valider l'équidistance intermédiaire
            If nValeurCourbeIntermediaire > 0 Then
                'Calcul du modulo entre la valeur de l'équidistance de l'élément et l'équidistance Intermediaire
                dDiffIntermediaire = CInt(pFeature.Value(nPosAttElevation)) Mod nValeurCourbeIntermediaire

                'Vérifier si l'élévation correspond à une courbe Intermediaire
                If dDiffIntermediaire = 0 Then
                    'Vérifier si le code de la courbe Intermediaire est spécifié
                    If nCodeCourbeIntermediaire > 0 Then
                        'Vérifier si le code de l'élément est invalide
                        If CInt(pFeature.Value(nPosAttCode)) <> nCodeCourbeIntermediaire Then
                            'Ajouter la note de l'incohérence pour le code de courbe Intermediaire
                            sNote = "Code de la courbe intermédiaire:" & nValeurCourbeIntermediaire _
                            & ", Attribut:" & pFeature.Fields.Field(nPosAttCode).Name _
                            & ", Ancienne valeur:" & CInt(pFeature.Value(nPosAttCode)) _
                            & ", Nouvelle valeur:" & nCodeCourbeIntermediaire
                        End If
                    End If

                    'Sortir de la validation
                    Return sNote
                End If

                'Vérifier si la différence Intermediaire possède la plus petite différence et s'il n'y a pas d'équidistance intercalaire
                If dDiffIntermediaire < dDiffMaitresse And nValeurCourbeIntercalaire = 0 Then
                    'L'équidistance Intermediaire est en erreur
                    sNote = "Équidistance intermédiaire:"
                    nEquidistance = nValeurCourbeIntermediaire
                    dDifference = dDiffIntermediaire
                End If
            End If

            'Vérifier si on doit valider l'équidistance Intercalaire
            If nValeurCourbeIntercalaire > 0 Then
                'Calcul du modulo entre la valeur de l'équidistance de l'élément et l'équidistance Intercalaire
                dDiffIntercalaire = CInt(pFeature.Value(nPosAttElevation)) Mod nValeurCourbeIntercalaire

                'Vérifier si l'élévation correspond à une courbe Intermediaire
                If dDiffIntercalaire = 0 Then
                    'Vérifier si le code de la courbe Intercalaire est spécifié
                    If nCodeCourbeIntercalaire > 0 Then
                        'Vérifier si le code de l'élément est invalide
                        If CInt(pFeature.Value(nPosAttCode)) <> nCodeCourbeIntercalaire Then
                            'Ajouter la note de l'incohérence pour le code de courbe Intermediaire
                            sNote = "Code de la courbe intercalaire:" & nValeurCourbeIntercalaire _
                            & ", Attribut:" & pFeature.Fields.Field(nPosAttCode).Name _
                            & ", Ancienne valeur:" & CInt(pFeature.Value(nPosAttCode)) _
                            & ", Nouvelle valeur:" & nCodeCourbeIntercalaire
                        End If
                    End If

                    'Sortir de la validation
                    Return sNote
                End If

                'Vérifier si la différence intercalaire possède la plus petite différence
                If dDiffIntercalaire < dDifference Then
                    'L'équidistance intercalaire est en erreur
                    sNote = "Équidistance intercalaire:"
                    nEquidistance = nValeurCourbeIntercalaire
                End If
            End If

            'Calcul de la nouvelle valeur d'élévation
            nNouvelleValeur = CInt((CInt(pFeature.Value(nPosAttElevation)) / nEquidistance))
            'Calcul de la nouvelle valeur d'élévation
            nNouvelleValeur = nNouvelleValeur * nEquidistance
            'Ajouter la note de l'incohérence
            sNote = sNote & nEquidistance _
                    & ", Attribut:" & pFeature.Fields.Field(nPosAttElevation).Name _
                    & ", Ancienne valeur:" & CInt(pFeature.Value(nPosAttElevation)) _
                    & ", Nouvelle valeur:" & nNouvelleValeur

            'Définir le message d'erreur trouvé
            ValiderValeurEquidistanceFeature = sNote

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            nEquidistance = Nothing
            nNouvelleValeur = Nothing
            dDiffMaitresse = Nothing
            dDiffIntermediaire = Nothing
            dDiffIntercalaire = Nothing
            dDifference = Nothing
            sNote = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de créer un FeatureLayer contenant la FeatureClass de la grille de validation en mémoire.
    '''</summary>
    ''' 
    '''<param name="pGrilleValidation"> GeometryBag  contenant les lignes de la grille de validation.</param>
    '''
    '''<return>Le FeatureLayer de la grille de validation.</return>
    ''' 
    Private Function CreerFeatureLayerGrilleValidation(ByVal pGrilleValidation As IGeometryBag) As IFeatureLayer
        'Déclarer les variables de travail
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface contenant la FeatureClass de la grille de validation.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisée pour ajouter les lignes de grille de validation.
        Dim pFeatureBuffer As IFeatureBuffer = Nothing      'Interface ESRI contenant l'élément de l'incohérence à créer.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes de la grille de validation.
        Dim pClone As IClone = Nothing                      'Interface pour cloner une géométrie.
        Dim pGeoDataset As IGeoDataset = Nothing            'Interface pour extraire la référence spatiale

        'Définir la valeur par défaut
        CreerFeatureLayerGrilleValidation = Nothing

        Try
            'Créer une classe contenant la grille de validation en mémoire
            pFeatureClass = CreateInMemoryFeatureClass("GrilleValidation", gpFeatureLayerSelection)

            'Interface pour extraire la référence spatiale
            pGeoDataset = CType(pFeatureClass, IGeoDataset)

            'Projeter la grille de validation
            'pGrilleValidation.Project(pGeoDataset.SpatialReference)

            'Interface pour créer les lignes de la grille de validation
            pFeatureCursor = pFeatureClass.Insert(True)

            'Interface pour extraire les lignes de la grille de validation
            pGeomColl = CType(pGrilleValidation, IGeometryCollection)

            'Traiter toutes les lignes de la grille
            For i = 0 To pGeomColl.GeometryCount - 1
                'Créer un FeatureBuffer Point
                pFeatureBuffer = pFeatureClass.CreateFeatureBuffer

                'Interface pour cloner la géométrie
                pClone = CType(pGeomColl.Geometry(i), IClone)

                'Définir la géométrie
                pFeatureBuffer.Shape = CType(pClone.Clone, IGeometry)

                'Définir la description
                pFeatureBuffer.Value(1) = "Grille " & i.ToString

                'Insérer un nouvel élément dans la FeatureClass d'erreur
                pFeatureCursor.InsertFeature(pFeatureBuffer)
            Next

            'Conserver toutes les modifications
            pFeatureCursor.Flush()

            'Release the update cursor to remove the lock on the input data.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatureCursor)

            'Créer un nouveau FeatureLayer
            CreerFeatureLayerGrilleValidation = New FeatureLayer

            'Définir le nom du FeatureLayer selon le nom et la date
            CreerFeatureLayerGrilleValidation.Name = pFeatureClass.AliasName

            'Rendre visible le FeatureLayer
            CreerFeatureLayerGrilleValidation.Visible = True

            'Définir la Featureclass dans le FeatureLayer
            CreerFeatureLayerGrilleValidation.FeatureClass = pFeatureClass

            'Définir la référence spatiale
            CreerFeatureLayerGrilleValidation.SpatialReference = pGeoDataset.SpatialReference

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureClass = Nothing
            pFeatureCursor = Nothing
            pFeatureBuffer = Nothing
            pGeomColl = Nothing
            pClone = Nothing
            pGeoDataset = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non les équidistances d'élévation spécifiées.
    ''' L'équidistance correspond à une distance verticale séparant deux courbes de niveau.
    ''' Le calcul de la distance verticale est obtenue à l'aide d'une grille de lignes verticales.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les intersection entre les éléments sélectionnés et les lignes de la grille de validation qui respecte ou non les équidistances d'élévation spécifiées.</return>
    ''' 
    Private Function TraiterElevationCourbe(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les lignes de la grille de validation.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément traité.
        Dim pGeometryDecoupage As IPolygon = Nothing        'Interface contenant la géométrie de l'élément de découpage.
        Dim pGrilleValidation As IGeometryBag = Nothing     'Interface contenant la grille de validation.
        Dim oIncoherences As Collection = New Collection    'Objet contenant la liste des incohérences trouvées afin de ne pas faire de duplication.
        Dim pPointColl As IPointCollection = Nothing        'Interface qui permet d'extraire les sommets d'une ligne de la grille.
        Dim pPointCollErr As IPointCollection = Nothing     'Interface utilisé pour construire la géométrie de l'incohérence.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément traité.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim iOidSel(1) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iSel As Integer = -1                            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                            'Numéro de séquence de la géométrie en relation.
        Dim nDifference As Integer = Nothing                'Contient la différence de valeur d'élévation
        Dim sOid As String = Nothing                        'Contient la valeur du OID de l'élément
        Dim sMessage As String = ""                         'Contient le message à écrire.
        Dim bSucces As Boolean = False                      'Indique si l'identifiant et la géométrie respecte l'élément de découpage.

        'Définir la géométrie par défaut
        TraiterElevationCourbe = New GeometryBag
        TraiterElevationCourbe.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

        Try
            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterElevationCourbe, IGeometryCollection)

            'Vérifier si l'élément de découpage est valide
            If gpFeatureDecoupage Is Nothing Then
                'Créer la géométrie de découpage à partir de l'enveloppe du FeatureLayer de sélection
                pGeometryDecoupage = EnvelopeToPolygon(gpFeatureLayerSelection.AreaOfInterest)
            Else
                'Définir la géométrie de découpage
                pGeometryDecoupage = CType(gpFeatureDecoupage.ShapeCopy, IPolygon)
                pGeometryDecoupage.Project(TraiterElevationCourbe.SpatialReference)
            End If

            'Créer une grille de validation vide
            pGrilleValidation = New GeometryBag
            pGrilleValidation.SpatialReference = pGeometryDecoupage.SpatialReference
            'Vérifier comment on doit créer les lignes de la grile pour être plus rapide
            If pGeometryDecoupage.Envelope.Height > pGeometryDecoupage.Envelope.Width Then
                'Créer les lignes verticales dans la grille de validation pour chaque élément traité
                Call CreerLigneVerticaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            Else
                'Ajouter les lignes horizontales dans la grille de validation pour chaque élément traité
                Call CreerLigneHorizontaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            End If
            'Créer les lignes de la grille de validation dans un GeometryBag
            pGeomRelColl = CType(pGrilleValidation, IGeometryCollection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments en relation
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)
            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            'Interface contenant les nouveaux éléments à sélectionner
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de recherche des éléments en relation
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des intersections (" & gpFeatureLayerSelection.Name & ") ..."
            'Interface pour traiter la relation spatiale
            pRelOpNxM = CType(pGrilleValidation, IRelationalOperatorNxM)
            'Exécuter la recherche et retourner le résultat de la relation spatiale
            pRelResult = pRelOpNxM.Intersects(CType(pGeomSelColl, IGeometryBag))

            'Afficher le message de validation des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter les intersections : " + " (" + pRelResult.RelationElementCount.ToString + " relations) ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iRel, iSel)

                'Extraire l'élément traité
                pFeature = gpFeatureLayerSelection.FeatureClass.GetFeature(iOidSel(iSel))

                'Insérer dans la ligne de la grille les points d'intersection avec les courbes en relation
                Call InsererPointIntersectionLigne(pFeature, pGeomSelColl.Geometry(iSel), gnPosAttElevation, CType(pGeomRelColl.Geometry(iRel), IPolyline))

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next
            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

            'Afficher le message de validation des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Valider les élévations des courbes selon la grille : " + " (" + pGeomRelColl.GeometryCount.ToString + " Lignes) ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomRelColl.GeometryCount, pTrackCancel)
            'Traiter toutes les lignes de la grille
            For i = 0 To pGeomRelColl.GeometryCount - 1
                'Écrire la ligne de la grille dans la FeatureClass d'erreur
                'EcrireFeatureErreur("Grille " & i.ToString, pGeometryColl.Geometry(i))

                'Interface pour accéder aux sommets de la géométrie
                pPointColl = CType(pGeomRelColl.Geometry(i), IPointCollection)

                'Sortir si seulement 3 sommets ou moins car la ligne de la grille contenait 2 sommet au départ
                If pPointColl.PointCount > 3 Then
                    'Traiter tous les sommets de la ligne de la grille de validation
                    For j = 1 To pPointColl.PointCount - 3
                        'Vérifier si les OIDs sont différents
                        If pPointColl.Point(j).M <> pPointColl.Point(j + 1).M Then
                            'Calcul de la différence entre deux élévations consécutive
                            nDifference = CInt(Math.Abs(pPointColl.Point(j).Z - pPointColl.Point(j + 1).Z))

                            'Vérifier la différence d'élévation entre deux sommets consécutifs
                            If nDifference <> 0 And nDifference <> gnValeurCourbeIntermediaire And nDifference <> gnValeurCourbeIntercalaire Then
                                'Définir le résultat
                                sMessage = "#Équidistance invalide"
                                bSucces = False
                            Else
                                'Définir le résultat
                                sMessage = "#Équidistance valide"
                                bSucces = True
                            End If

                            'Vérifier si on doit sélectionner l'élément
                            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                'Définir la clé de l'incohérence contenant les deux OIDs d'éléments en erreur
                                If pPointColl.Point(j).M < pPointColl.Point(j + 1).M Then
                                    sOid = pPointColl.Point(j).M & "," & pPointColl.Point(j + 1).M
                                Else
                                    sOid = pPointColl.Point(j + 1).M & "," & pPointColl.Point(j).M
                                End If

                                'Vérifier si l'incohérence existe déja
                                If Not oIncoherences.Contains(sOid) Then
                                    'Ajouter la note de l'incohérence
                                    sMessage = "OID=" & sOid & sMessage _
                                            & " /Élévation:" & CInt(pPointColl.Point(j).Z).ToString & "," & CInt(pPointColl.Point(j + 1).Z).ToString _
                                            & " /" & Expression

                                    'Ajouter l'incohérence
                                    oIncoherences.Add(sMessage, sOid)

                                    'Ajouter les éléments dans la sélection
                                    pSelectionSet.Add(CInt(pPointColl.Point(j).M))
                                    pSelectionSet.Add(CInt(pPointColl.Point(j + 1).M))

                                    'Créer une nouvelle géométrie de type ligne
                                    pGeometry = New Polyline
                                    pGeometry.SpatialReference = pGeometryDecoupage.SpatialReference
                                    'Interface pour construire la géométrie de l'incohérence
                                    pPointCollErr = CType(pGeometry, IPointCollection)
                                    'Ajouter le premier point de l'incohérence
                                    pPointCollErr.AddPoint(pPointColl.Point(j))
                                    'Ajouter le deuxième point de l'incohérence
                                    pPointCollErr.AddPoint(pPointColl.Point(j + 1))

                                    'Ajouter l'enveloppe de l'élément sélectionné
                                    pGeomResColl.AddGeometry(pGeometry)

                                    'Écrire une erreur
                                    EcrireFeatureErreur(sMessage, pGeometry, CSng(pPointColl.Point(j).Z))
                                End If
                            End If
                        End If
                    Next j
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next i

            'Ajouter les éléments sélectionneés dans le FeatureLayer
            pFeatureSel.SelectionSet = pSelectionSet

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pGeomResColl = Nothing
            pGeometry = Nothing
            pGeometryDecoupage = Nothing
            pGrilleValidation = Nothing
            pFeature = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            oIncoherences = Nothing
            pPointColl = Nothing
            pPointCollErr = Nothing
            GC.Collect()
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'insérer dans la ligne de la grille les points d'intersection entre la ligne de la grille et les courbes en relation.
    ''' Le point d'intersection contient la valeur d'élévation dans la coordonnée Z et contient le OID de l'élément dans la coordonnée M. 
    '''</summary>
    '''
    '''<param name="pFeature">Interface contenant une courbes de niveaux à valider.</param>
    '''<param name="pGeometry">Interface contenant la géométrie de la courbe.</param>
    '''<param name="nPosAttElevation">Contient la position de l'attribut contenant l'élévation.</param>
    '''<param name="pPolyline">Interface contenant une ligne de la grille de validation.</param>
    ''' 
    Public Sub InsererPointIntersectionLigne(ByVal pFeature As IFeature, ByVal pGeometry As IGeometry, ByVal nPosAttElevation As Integer, ByRef pPolyline As IPolyline)
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour trouver les points d'intersection
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les points d'intersection dans la ligne de la grille
        Dim pLignePointColl As IPointCollection = Nothing   'Interface utilisé pour insérer les points d'intersection dans la ligne de la grille
        Dim pPoint As IPoint = Nothing                      'Interface contenant le sommet d'intersection
        Dim pHitTest As IHitTest = Nothing                  'Interface utilisé pour extraire la position du point d'intersection à insérer
        Dim pNewPoint As New Point                          'Interface contenant le nouveau point d'intersection (on ne l'utlise pas car pareil)
        Dim dDistance As Double = Nothing                   'Contient la distance entre le point et le nouveau point (toujours=0 dans ce cas-ci)
        Dim nPart As Integer = Nothing                      'Contient le numéro de partie (toujours=0 dans ce cas-ci)
        Dim nSegment As Integer = Nothing                   'Contient le numéro de segment du point d'intersection
        Dim bRightSide As Boolean = Nothing                 'Indique si le point d'intersection est du côté droit

        Try
            'Interface pour trouver les points d'intersection entre la ligne et la géométrie de l'élément
            pTopoOp = CType(pPolyline, ITopologicalOperator2)

            'Interface pour insérer les points d'intersection dans la ligne de la grille
            pLignePointColl = CType(pPolyline, IPointCollection)

            'Interface pour extraire les sommets de chaque géométrie d'intersection de point
            pPointColl = CType(pTopoOp.Intersect(pGeometry, esriGeometryDimension.esriGeometry0Dimension), IPointCollection)
            'Traiter tous les points d'intersection
            For i = 0 To pPointColl.PointCount - 1
                'Interface pour extraire la position du point à insérer dans la ligne de la grille
                pHitTest = CType(pLignePointColl, IHitTest)

                'Définir le point à traiter
                pPoint = pPointColl.Point(i)

                'Définir l'élévation dans la valeur Z du Point
                pPoint.Z = CDbl(pFeature.Value(nPosAttElevation))

                'Définir l'identifiant de l'élément dans la valeur M du Point
                pPoint.M = pFeature.OID

                'Insérer le point sur la géométrie de l'intervalle
                pHitTest.HitTest(pPoint, 0.1, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                 pNewPoint, dDistance, nPart, nSegment, bRightSide)

                'Insérer un nouveau sommet (insertion avant le numéro de point identifié, donc nSegement + 1)
                pLignePointColl.InsertPoints(nSegment + 1, 1, pPoint)
            Next

            'Interface pour extraire les sommets de chaque géométrie d'intersection de ligne
            pPointColl = CType(pTopoOp.Intersect(pGeometry, esriGeometryDimension.esriGeometry1Dimension), IPointCollection)
            'Traiter tous les points d'intersection
            For i = 0 To pPointColl.PointCount - 1
                'Interface pour extraire la position du point à insérer dans la ligne de la grille
                pHitTest = CType(pLignePointColl, IHitTest)

                'Définir le point à traiter
                pPoint = pPointColl.Point(i)

                'Définir l'élévation dans la valeur Z du Point
                pPoint.Z = CDbl(pFeature.Value(nPosAttElevation))

                'Définir l'identifiant de l'élément dans la valeur M du Point
                pPoint.M = pFeature.OID

                'Insérer le point sur la géométrie de l'intervalle
                pHitTest.HitTest(pPoint, 0.1, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                 pNewPoint, dDistance, nPart, nSegment, bRightSide)

                'Insérer un nouveau sommet (insertion avant le numéro de point identifié, donc nSegement + 1)
                pLignePointColl.InsertPoints(nSegment + 1, 1, pPoint)
            Next

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pPointColl = Nothing
            pLignePointColl = Nothing
            pPoint = Nothing
            pHitTest = Nothing
            pNewPoint = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui respecte ou non les équidistances d'élévation spécifiées.
    ''' L'équidistance correspond à une distance verticale séparant deux courbes de niveau.
    ''' Le calcul de la distance verticale est obtenue à l'aide d'une grille de lignes verticales.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les intersection entre les éléments sélectionnés et les lignes de la grille de validation qui respecte ou non les équidistances d'élévation spécifiées.</return>
    ''' 
    Private Function TraiterEquidistanceCourbe(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément traité.
        Dim pGeometryDecoupage As IPolygon = Nothing        'Interface contenant la géométrie de l'élément de découpage.
        Dim pGrilleValidation As IGeometryBag = Nothing     'Interface contenant la grille de validation
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface contenant les lignes de la grille de validation
        Dim oIncoherences As Collection = New Collection    'Objet contenant la liste des incohérences trouvées afin de ne pas faire de duplication
        Dim pPointColl As IPointCollection = Nothing        'Interface qui permet d'extraire les sommets d'une ligne de la grille
        Dim pPointCollErr As IPointCollection = Nothing     'Interface utilisé pour construire la géométrie de l'incohérence
        Dim nDifference As Integer = Nothing                'Contient la différence de valeur d'élévation
        Dim sOid As String = Nothing                        'Contient la valeur du OID de l'élément
        Dim sMessage As String = ""                         'Contient le message à écrire.
        Dim bSucces As Boolean = False                      'Indique si l'identifiant et la géométrie respecte l'élément de découpage.

        'Définir la géométrie par défaut
        TraiterEquidistanceCourbe = New GeometryBag
        TraiterEquidistanceCourbe.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

        Try
            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterEquidistanceCourbe, IGeometryCollection)

            'Vérifier si l'élément de découpage est valide
            If gpFeatureDecoupage Is Nothing Then
                'Créer la géométrie de découpage à partir de l'enveloppe du FeatureLayer de sélection
                pGeometryDecoupage = EnvelopeToPolygon(gpFeatureLayerSelection.AreaOfInterest)
            Else
                'Définir la géométrie de découpage
                pGeometryDecoupage = CType(gpFeatureDecoupage.ShapeCopy, IPolygon)
                pGeometryDecoupage.Project(TraiterEquidistanceCourbe.SpatialReference)
            End If

            'Créer la grille de validation contenant des lignes verticales de base
            'pGrilleValidation = CreerGrilleLigneVerticale(pGeometryDecoupage, pGeometryDecoupage.Envelope.Width / 10, pTrackCancel)

            'Créer une grille de validation vide
            pGrilleValidation = New GeometryBag
            pGrilleValidation.SpatialReference = pGeometryDecoupage.SpatialReference

            'Vérifier comment on doit créer les lignes de la grile pour être plus rapide
            If pGeometryDecoupage.Envelope.Height > pGeometryDecoupage.Envelope.Width Then
                'Créer les lignes verticales dans la grille de validation pour chaque élément traité
                Call CreerLigneVerticaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            Else
                'Ajouter les lignes horizontales dans la grille de validation pour chaque élément traité
                Call CreerLigneHorizontaleParElement(pGeometryDecoupage, pGrilleValidation, pTrackCancel)
            End If

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Enlever la sélection
            pFeatureSel.Clear()

            'Interface contenant les nouveaux éléments à sélectionner
            pSelectionSet = pFeatureSel.SelectionSet

            'Créer les lignes de la grille de validation dans un GeometryBag
            pGeometryColl = CType(pGrilleValidation, IGeometryCollection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Valider les équidistances des courbes selon la grille : " + " (" + pGeometryColl.GeometryCount.ToString + " Lignes) ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pGeometryColl.GeometryCount, pTrackCancel)

            'Traiter toutes les lignes de la grille de validation
            For i = 0 To pGeometryColl.GeometryCount - 1
                'Insérer dans la ligne de la grille les points d'intersection avec les courbes en relation
                Call InsererPointIntersectionLigneGrille(CType(pGeometryColl.Geometry(i), IPolyline), gpFeatureLayerSelection, gnPosAttElevation)

                'Écrire la ligne de la grille dans la FeatureClass d'erreur
                'EcrireFeatureErreur("Grille " & i.ToString, pGeometryColl.Geometry(i))

                'Interface pour accéder aux sommets de la géométrie
                pPointColl = CType(pGeometryColl.Geometry(i), IPointCollection)

                'Sortir si seulement 3 sommets ou moins car la ligne de la grille contenait 2 sommet au départ
                If pPointColl.PointCount > 3 Then
                    'Traiter tous les sommets de la ligne de la grille de validation
                    For j = 1 To pPointColl.PointCount - 3
                        'Vérifier si les OIDs sont différents
                        If pPointColl.Point(j).M <> pPointColl.Point(j + 1).M Then
                            'Calcul de la différence entre deux élévations consécutive
                            nDifference = CInt(Math.Abs(pPointColl.Point(j).Z - pPointColl.Point(j + 1).Z))

                            'Vérifier la différence d'élévation entre deux sommets consécutifs
                            If nDifference <> 0 And nDifference <> gnValeurCourbeIntermediaire And nDifference <> gnValeurCourbeIntercalaire Then
                                'Définir le résultat
                                sMessage = "#Équidistance invalide"
                                bSucces = False
                            Else
                                'Définir le résultat
                                sMessage = "#Équidistance valide"
                                bSucces = True
                            End If

                            'Vérifier si on doit sélectionner l'élément
                            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                                'Définir la clé de l'incohérence contenant les deux OIDs d'éléments en erreur
                                If pPointColl.Point(j).M < pPointColl.Point(j + 1).M Then
                                    sOid = pPointColl.Point(j).M & "," & pPointColl.Point(j + 1).M
                                Else
                                    sOid = pPointColl.Point(j + 1).M & "," & pPointColl.Point(j).M
                                End If

                                'Vérifier si l'incohérence existe déja
                                If Not oIncoherences.Contains(sOid) Then
                                    'Ajouter la note de l'incohérence
                                    sMessage = "OID=" & sOid & sMessage _
                                            & " /Élévation:" & CInt(pPointColl.Point(j).Z).ToString & "," & CInt(pPointColl.Point(j + 1).Z).ToString _
                                            & " /" & Expression

                                    'Ajouter l'incohérence
                                    oIncoherences.Add(sMessage, sOid)

                                    'Ajouter les éléments dans la sélection
                                    pSelectionSet.Add(CInt(pPointColl.Point(j).M))
                                    pSelectionSet.Add(CInt(pPointColl.Point(j + 1).M))

                                    'Créer une nouvelle géométrie de type ligne
                                    pGeometry = New Polyline
                                    pGeometry.SpatialReference = pGeometryDecoupage.SpatialReference
                                    'Interface pour construire la géométrie de l'incohérence
                                    pPointCollErr = CType(pGeometry, IPointCollection)
                                    'Ajouter le premier point de l'incohérence
                                    pPointCollErr.AddPoint(pPointColl.Point(j))
                                    'Ajouter le deuxième point de l'incohérence
                                    pPointCollErr.AddPoint(pPointColl.Point(j + 1))

                                    'Ajouter l'enveloppe de l'élément sélectionné
                                    pGeomSelColl.AddGeometry(pGeometry)

                                    'Écrire une erreur
                                    EcrireFeatureErreur(sMessage, pGeometry, CSng(pPointColl.Point(j).Z))
                                End If
                            End If
                        End If
                    Next j
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next i

            'Ajouter les éléments sélectionneés dans le FeatureLayer
            pFeatureSel.SelectionSet = pSelectionSet

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pGeomSelColl = Nothing
            pGeometry = Nothing
            pGeometryDecoupage = Nothing
            pGrilleValidation = Nothing
            pGeometryColl = Nothing
            oIncoherences = Nothing
            pPointColl = Nothing
            pPointCollErr = Nothing
            GC.Collect()
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'insérer dans la ligne de la grille les points d'intersection entre la ligne de la grille et les courbes en relation.
    ''' Le point d'intersection contient la valeur d'élévation dans la coordonnée Z et contient le OID de l'élément dans la coordonnée M. 
    '''</summary>
    '''
    '''<param name="pPolyline">Interface contenant une ligne de la grille de validation.</param>
    '''<param name="pFeatureLayer">Interface contenant les courbes de niveaux à valider.</param>
    '''<param name="nPosAttElevation">Contient la position de l'attribut contenant l'élévation.</param>
    ''' 
    Public Sub InsererPointIntersectionLigneGrille(ByRef pPolyline As IPolyline, ByRef pFeatureLayer As IFeatureLayer, ByVal nPosAttElevation As Integer)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisée pour sélectionner les éléments
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface contenant la relation spatiale de base.
        Dim pEnumId As IEnumIDs = Nothing                   'Interface qui permet d'extraire les Ids des éléments sélectionnés.
        Dim iOid As Integer = Nothing                       'Contient un Id d'un élément sélectionné.
        Dim pFeature As IFeature = Nothing                  'Interface contenant un élément en relation  avec la ligne de la grille
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour trouver les points d'intersection
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les points d'intersection dans la ligne de la grille
        Dim pLignePointColl As IPointCollection = Nothing   'Interface utilisé pour insérer les points d'intersection dans la ligne de la grille
        Dim pPoint As IPoint = Nothing                      'Interface contenant le sommet d'intersection
        Dim pHitTest As IHitTest = Nothing                  'Interface utilisé pour extraire la position du point d'intersection à insérer
        Dim pNewPoint As New Point                          'Interface contenant le nouveau point d'intersection (on ne l'utlise pas car pareil)
        Dim dDistance As Double = Nothing                   'Contient la distance entre le point et le nouveau point (toujours=0 dans ce cas-ci)
        Dim nPart As Integer = Nothing                      'Contient le numéro de partie (toujours=0 dans ce cas-ci)
        Dim nSegment As Integer = Nothing                   'Contient le numéro de segment du point d'intersection
        Dim bRightSide As Boolean = Nothing                 'Indique si le point d'intersection est du côté droit

        Try
            'Interface pour trouver les points d'intersection entre la ligne et la géométrie de l'élément
            pTopoOp = CType(pPolyline, ITopologicalOperator2)

            'Interface pour insérer les points d'intersection dans la ligne de la grille
            pLignePointColl = CType(pPolyline, IPointCollection)

            'Créer la requête spatiale
            pSpatialFilter = New SpatialFilterClass
            'Définir la relation d'intersection
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            'Définir la référence spatiale de sortie dans la requête spatiale
            pSpatialFilter.OutputSpatialReference(pFeatureLayer.FeatureClass.ShapeFieldName) = pPolyline.SpatialReference
            'Définir la géométrie utilisée pour la relation spatiale
            pSpatialFilter.Geometry = pPolyline

            'Interface qui permet la sélection des éléments
            pFeatureSel = CType(pFeatureLayer, IFeatureSelection)

            'Rechercher les éléments
            pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, False)

            'Interface pour extrire la liste des Ids
            pEnumId = pFeatureSel.SelectionSet.IDs

            'Trouver le premier élément en relation
            pEnumId.Reset()
            iOid = pEnumId.Next

            'Traiter tant qu'il y a des éléments en relation
            Do Until iOid = -1
                'Extraire l'élément en relation
                pFeature = pFeatureLayer.FeatureClass.GetFeature(iOid)

                'Définir la géométrie
                pGeometry = pFeature.ShapeCopy
                pGeometry.Project(pPolyline.SpatialReference)

                'Interface pour extraire les sommets de chaque géométrie d'intersection de point
                pPointColl = CType(pTopoOp.Intersect(pGeometry, esriGeometryDimension.esriGeometry0Dimension), IPointCollection)
                'Traiter tous les points d'intersection
                For i = 0 To pPointColl.PointCount - 1
                    'Interface pour extraire la position du point à insérer dans la ligne de la grille
                    pHitTest = CType(pLignePointColl, IHitTest)

                    'Définir le point à traiter
                    pPoint = pPointColl.Point(i)

                    'Définir l'élévation dans la valeur Z du Point
                    pPoint.Z = CDbl(pFeature.Value(nPosAttElevation))

                    'Définir l'identifiant de l'élément dans la valeur M du Point
                    pPoint.M = pFeature.OID

                    'Insérer le point sur la géométrie de l'intervalle
                    pHitTest.HitTest(pPoint, 0.1, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                     pNewPoint, dDistance, nPart, nSegment, bRightSide)

                    'Insérer un nouveau sommet (insertion avant le numéro de point identifié, donc nSegement + 1)
                    pLignePointColl.InsertPoints(nSegment + 1, 1, pPoint)
                Next

                'Interface pour extraire les sommets de chaque géométrie d'intersection de ligne
                pPointColl = CType(pTopoOp.Intersect(pGeometry, esriGeometryDimension.esriGeometry1Dimension), IPointCollection)
                'Traiter tous les points d'intersection
                For i = 0 To pPointColl.PointCount - 1
                    'Interface pour extraire la position du point à insérer dans la ligne de la grille
                    pHitTest = CType(pLignePointColl, IHitTest)

                    'Définir le point à traiter
                    pPoint = pPointColl.Point(i)

                    'Définir l'élévation dans la valeur Z du Point
                    pPoint.Z = CDbl(pFeature.Value(nPosAttElevation))

                    'Définir l'identifiant de l'élément dans la valeur M du Point
                    pPoint.M = pFeature.OID

                    'Insérer le point sur la géométrie de l'intervalle
                    pHitTest.HitTest(pPoint, 0.1, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                     pNewPoint, dDistance, nPart, nSegment, bRightSide)

                    'Insérer un nouveau sommet (insertion avant le numéro de point identifié, donc nSegement + 1)
                    pLignePointColl.InsertPoints(nSegment + 1, 1, pPoint)
                Next

                'Définir le prochain élément en relation
                iOid = pEnumId.Next
            Loop

            'Vider la sélection
            pFeatureSel.Clear()

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSpatialFilter = Nothing
            pEnumId = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
            pLignePointColl = Nothing
            pPoint = Nothing
            pHitTest = Nothing
            pNewPoint = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de créer des lignes verticales dans la grille de validation pour chaque élément traité.
    '''</summary>
    ''' 
    '''<param name="pGeometryDecoupage"> Interface contenant la géométrie de découpage.</param>
    '''<param name="pGrilleValidation"> Interface contenant les lignes verticale de la grille de validation.</param>
    '''<param name="pTrackCancel"> Interface qui permet d'annuler et afficher le traitement.</param>
    '''
    '''<return>Les intersection entre les éléments sélectionnés et les lignes de la grille de validation qui respecte ou non les équidistances d'élévation spécifiées.</return>
    ''' 
    Private Sub CreerLigneVerticaleParElement(ByVal pGeometryDecoupage As IPolygon, ByRef pGrilleValidation As IGeometryBag, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                        'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing          'Interface utilisé pour extraire les éléments à traiter.
        Dim pQueryFilter As IQueryFilter = Nothing              'Interface contenant la requête de tri.
        Dim pQueryFilterDef As IQueryFilterDefinition = Nothing 'Interface utilisé pour définir la requête de tri.
        Dim pFeature As IFeature = Nothing                      'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                    'Interface contenant la géométrie de l'élément traité.
        Dim sAttributLongueur As String = "Shape_Length"        'Nom de l'attribut contenant la longueur de la courbe pour le tri.

        Try
            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer les lignes verticales par élément (" & pSelectionSet.Count.ToString & " éléments) ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Traiter tous les attributs
            For i = 0 To gpFeatureLayerSelection.FeatureClass.Fields.FieldCount - 1
                'Vérifier si l'attribut est de type Double et non-éditable
                If gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Type = esriFieldType.esriFieldTypeDouble _
                And gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Editable = False Then
                    'Définir le nom de l'attribut de longueur
                    sAttributLongueur = gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Name
                    'Sortir
                    Exit For
                End If
            Next

            'Vérifier un nom d'attribut de longueur a été spécifié pour le tri
            If gpFeatureLayerSelection.FeatureClass.FindField(sAttributLongueur) >= 0 Then
                'Définir la requête de tri ascendant selon la longueur des courbes
                pQueryFilter = New QueryFilter
                pQueryFilterDef = CType(pQueryFilter, IQueryFilterDefinition)
                pQueryFilterDef.PostfixClause = "ORDER BY " & sAttributLongueur
                pQueryFilter.OutputSpatialReference(gpFeatureLayerSelection.FeatureClass.ShapeFieldName) = pGeometryDecoupage.SpatialReference
            End If

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(pQueryFilter, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Définir la géométrie
                pGeometry = pFeature.ShapeCopy
                pGeometry.Project(pGeometryDecoupage.SpatialReference)

                'Créer une ligne dans la grille de validation à partir de chaque géométrie si nécessaire
                Call CreerLigneVerticaleGeometrie(CType(pGeometry, IPolyline), pGeometryDecoupage, pGrilleValidation)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pQueryFilter = Nothing
            pQueryFilterDef = Nothing
            pFeature = Nothing
            pGeometry = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de créer des lignes horizontales dans la grille de validation pour chaque élément traité.
    '''</summary>
    ''' 
    '''<param name="pGeometryDecoupage"> Interface contenant la géométrie de découpage.</param>
    '''<param name="pGrilleValidation"> Interface contenant les lignes verticale de la grille de validation.</param>
    '''<param name="pTrackCancel"> Interface qui permet d'annuler et afficher le traitement.</param>
    '''
    '''<return>Les intersection entre les éléments sélectionnés et les lignes de la grille de validation qui respecte ou non les équidistances d'élévation spécifiées.</return>
    ''' 
    Private Sub CreerLigneHorizontaleParElement(ByVal pGeometryDecoupage As IPolygon, ByRef pGrilleValidation As IGeometryBag, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                        'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing          'Interface utilisé pour extraire les éléments à traiter.
        Dim pQueryFilter As IQueryFilter = Nothing              'Interface contenant la requête de tri.
        Dim pQueryFilterDef As IQueryFilterDefinition = Nothing 'Interface utilisé pour définir la requête de tri.
        Dim pFeature As IFeature = Nothing                      'Interface contenant l'élément à traiter.
        Dim pGeometry As IGeometry = Nothing                    'Interface contenant la géométrie de l'élément traité.
        Dim sAttributLongueur As String = "Shape_Length"        'Nom de l'attribut contenant la longueur de la courbe pour le tri.

        Try
            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer les lignes horizontales par élément (" & pSelectionSet.Count.ToString & " éléments) ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pSelectionSet.Count, pTrackCancel)

            'Traiter tous les attributs
            For i = 0 To gpFeatureLayerSelection.FeatureClass.Fields.FieldCount - 1
                'Vérifier si l'attribut est de type Double et non-éditable
                If gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Type = esriFieldType.esriFieldTypeDouble _
                And gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Editable = False Then
                    'Définir le nom de l'attribut de longueur
                    sAttributLongueur = gpFeatureLayerSelection.FeatureClass.Fields.Field(i).Name
                    'Sortir
                    Exit For
                End If
            Next

            'Vérifier un nom d'attribut de longueur a été spécifié pour le tri
            If gpFeatureLayerSelection.FeatureClass.FindField(sAttributLongueur) >= 0 Then
                'Définir la requête de tri ascendant selon la longueur des courbes
                pQueryFilter = New QueryFilter
                pQueryFilterDef = CType(pQueryFilter, IQueryFilterDefinition)
                pQueryFilterDef.PostfixClause = "ORDER BY " & sAttributLongueur
                pQueryFilter.OutputSpatialReference(gpFeatureLayerSelection.FeatureClass.ShapeFieldName) = pGeometryDecoupage.SpatialReference
            End If

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(pQueryFilter, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Définir la géométrie
                pGeometry = pFeature.ShapeCopy
                pGeometry.Project(pGeometryDecoupage.SpatialReference)

                'Créer une ligne dans la grille de validation à partir de chaque géométrie si nécessaire
                Call CreerLigneHorizontaleGeometrie(CType(pGeometry, IPolyline), pGeometryDecoupage, pGrilleValidation)

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pQueryFilter = Nothing
            pQueryFilterDef = Nothing
            pFeature = Nothing
            pGeometry = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de créer et retourner un "GeometryBag" contenant toutes les lignes verticales
    ''' de la grille de validation des courbes de niveau en fonction d'un polygon de découpage.
    ''' 
    ''' Toutes les lignes de la grille de validation sont des "Polyline" qui sont découpées en fonction du polygone de découpage.
    '''</summary>
    '''
    '''<param name="pPolygon">Interface contenant le polygone de découpage utilisé pour créer les lignes de la grille.</param>
    '''<param name="dDistance">Distance utilisée pour créer une ligne d'intervalle.</param>
    '''<param name="pTrackCancel">Interface qui permet d'afficher et annuler l'exécution du traitement.</param>
    '''
    '''<returns>"GeometryBag" contenant toutes les "Polyline" de la grille de validation des courbes de niveau.</returns>
    ''' 
    Private Function CreerGrilleLigneVerticale(ByVal pPolygon As IPolygon, ByVal dDistance As Double, ByRef pTrackCancel As ITrackCancel) As IGeometryBag
        'Déclarer les variables de travail
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter toutes les lignes de la grille de validation
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne de la grille de validation des courbes
        Dim pStepPro As IStepProgressor = Nothing           'Interface permettant d'afficher la progression du traitement
        Dim dOrigineX As Double = Nothing                   'Valeur X d'origine

        'Créer un nouveau GeometryBag
        CreerGrilleLigneVerticale = New GeometryBag
        CreerGrilleLigneVerticale.SpatialReference = pPolygon.SpatialReference

        Try
            'Sortir si aucune distance spécifié
            If dDistance > 0 Then
                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer les lignes verticales de base dans la grille de validation ..."
                'Définir les propriétés du Step Progressor
                pStepPro = CType(pTrackCancel.Progressor, IStepProgressor)
                pStepPro.MinRange = CInt(pPolygon.Envelope.XMin)
                pStepPro.MaxRange = CInt(pPolygon.Envelope.XMax)
                pStepPro.StepValue = CInt(dDistance)
                pStepPro.Position = CInt(pPolygon.Envelope.XMin)
                pStepPro.Show()

                'Définir la valeur de retour
                pGeometryColl = CType(CreerGrilleLigneVerticale, IGeometryCollection)

                'Définir la valeur X d'origine
                dOrigineX = pPolygon.Envelope.XMin

                'Créer toutes les intervalles
                Do Until dOrigineX = pPolygon.Envelope.XMax
                    'Vérifier si l'origine est plus grande que le maximum
                    If dOrigineX > pPolygon.Envelope.XMax Then
                        'Définir l'origine comme étant celle du maximum
                        dOrigineX = pPolygon.Envelope.XMax
                    End If

                    'Créer une nouvelle ligne verticale
                    pPolyline = CreerLigneVerticale(pPolygon, dOrigineX)

                    'Vérifier si l'origine est plus grande que le maximum
                    If dOrigineX <> pPolygon.Envelope.XMax Then
                        'Ajouter la distance de déplacement à la valeur X d'origine
                        dOrigineX = dOrigineX + dDistance
                    End If

                    'Ajouter la nouvelle ligne à la grille de validation
                    If Not pPolyline.IsEmpty Then pGeometryColl.AddGeometry(pPolyline)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do
                Loop

                'Fermer le Step Progressor
                pStepPro.Hide()
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryColl = Nothing
            pPolyline = Nothing
            dOrigineX = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de créer et retourner un "GeometryBag" contenant toutes les lignes horizontales
    ''' de la grille de validation des courbes de niveau en fonction d'un polygon de découpage.
    ''' 
    ''' Toutes les lignes de la grille de validation sont des "Polyline" qui sont découpées en fonction du polygone de découpage.
    '''</summary>
    '''
    '''<param name="pPolygon">Interface contenant le polygone de découpage utilisé pour créer les lignes de la grille.</param>
    '''<param name="dDistance">Distance utilisée pour créer une ligne d'intervalle.</param>
    '''<param name="pTrackCancel">Interface qui permet d'afficher et annuler l'exécution du traitement.</param>
    '''
    '''<returns>"GeometryBag" contenant toutes les "Polyline" de la grille de validation des courbes de niveau.</returns>
    ''' 
    Private Function CreerGrilleLigneHorizontale(ByVal pPolygon As IPolygon, ByVal dDistance As Double, ByRef pTrackCancel As ITrackCancel) As IGeometryBag
        'Déclarer les variables de travail
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter toutes les lignes de la grille de validation
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne de la grille de validation des courbes
        Dim pStepPro As IStepProgressor = Nothing           'Interface permettant d'afficher la progression du traitement
        Dim dOrigineY As Double = Nothing                   'Valeur Y d'origine

        'Créer un nouveau GeometryBag
        CreerGrilleLigneHorizontale = New GeometryBag
        CreerGrilleLigneHorizontale.SpatialReference = pPolygon.SpatialReference

        Try
            'Sortir si aucune distance spécifié
            If dDistance > 0 Then
                'Définir les propriétés du Step Progressor
                pStepPro = CType(pTrackCancel.Progressor, IStepProgressor)
                pStepPro.MinRange = CInt(pPolygon.Envelope.YMin)
                pStepPro.MaxRange = CInt(pPolygon.Envelope.YMax)
                pStepPro.StepValue = CInt(dDistance)
                pStepPro.Message = "Créer les lignes horizontales dans la grille de validation ..."
                pStepPro.Position = CInt(pPolygon.Envelope.YMin)
                pStepPro.Show()

                'Définir la valeur de retour
                pGeometryColl = CType(CreerGrilleLigneHorizontale, IGeometryCollection)

                'Définir la valeur X d'origine
                dOrigineY = pPolygon.Envelope.YMin

                'Créer toutes les intervalles
                Do Until dOrigineY = pPolygon.Envelope.YMax
                    'Vérifier si l'origine est plus grande que le maximum
                    If dOrigineY > pPolygon.Envelope.YMax Then
                        'Définir l'origine comme étant celle du maximum
                        dOrigineY = pPolygon.Envelope.YMax
                    End If

                    'Créer une nouvelle ligne horizontale
                    pPolyline = CreerLigneHorizontale(pPolygon, dOrigineY)

                    'Vérifier si l'origine est plus grande que le maximum
                    If dOrigineY <> pPolygon.Envelope.YMax Then
                        'Ajouter la distance de déplacement à la valeur X d'origine
                        dOrigineY = dOrigineY + dDistance
                    End If

                    'Ajouter la nouvelle ligne à la grille de validation
                    pGeometryColl.AddGeometry(pPolyline)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do
                Loop

                'Fermer le Step Progressor
                pStepPro.Hide()
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryColl = Nothing
            pPolyline = Nothing
            dOrigineY = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de créer une nouvelle ligne verticale dans la grille de validation en fonction de la géométrie 
    ''' et du polygone de découpage.
    ''' Une nouvelle ligne sera créée seulement si elle ne touche pas à aucune autre ligne. 
    ''' La grille de validation de type "GeometryBag" contient les lignes de type "Polyline".
    '''</summary>
    '''
    '''<param name="pGeometrie">Interface contenant une géométrie utilisé pour créer une ligne de la grille.</param>
    '''<param name="pPolygon">Interface contenant le polygone de découpage utilisé pour créer les lignes de la grille.</param>
    '''<param name="pGrilleValidation">Interface contenant les anciennes et les nouvelles lignes de la grille.</param>
    ''' 
    Private Sub CreerLigneVerticaleGeometrie(ByVal pGeometrie As IGeometry, ByVal pPolygon As IPolygon, ByRef pGrilleValidation As IGeometryBag)
        'Déclarer les variables de travail
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter les lignes à la grille de validation
        Dim pLigneVerticale As IPolyline = Nothing          'Interface contenant une ligne de la grille de validation
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier si la géométrie de l'élément intersecte la grille

        Try
            'Interface pour acceder à toutes les lignes de la grille
            pGeometryColl = CType(pGrilleValidation, IGeometryCollection)
            'Interface pour vérifier l'intersection
            pRelOp = CType(pGrilleValidation, IRelationalOperator)

            'Vérifier l'intersection
            If pRelOp.Disjoint(pGeometrie) Then
                'Créer une nouvelle ligne verticale
                pLigneVerticale = CreerLigneVerticale(pPolygon, ((pGeometrie.Envelope.XMin + pGeometrie.Envelope.XMax) / 2))

                'Ajouter la nouvelle ligne à la grille de validation
                pGeometryColl.AddGeometry(pLigneVerticale)
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryColl = Nothing
            pLigneVerticale = Nothing
            pRelOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de créer une nouvelle ligne horizontale dans la grille de validation en fonction de la géométrie 
    ''' et du polygone de découpage.
    ''' Une nouvelle ligne sera créée seulement si elle ne touche pas à aucune autre ligne. 
    ''' La grille de validation de type "GeometryBag" contient les lignes de type "Polyline".
    '''</summary>
    '''
    '''<param name="pGeometrie">Interface contenant une géométrie utilisé pour créer une ligne de la grille.</param>
    '''<param name="pPolygon">Interface contenant le polygone de découpage utilisé pour créer les lignes de la grille.</param>
    '''<param name="pGrilleValidation">Interface contenant les anciennes et les nouvelles lignes de la grille.</param>
    ''' 
    Private Sub CreerLigneHorizontaleGeometrie(ByVal pGeometrie As IGeometry, ByVal pPolygon As IPolygon, ByRef pGrilleValidation As IGeometryBag)
        'Déclarer les variables de travail
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter les lignes à la grille de validation
        Dim pLigneHorizontale As IPolyline = Nothing        'Interface contenant une ligne de la grille de validation
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier si la géométrie de l'élément intersecte la grille

        Try
            'Interface pour acceder à toutes les lignes de la grille
            pGeometryColl = CType(pGrilleValidation, IGeometryCollection)

            'Interface pour vérifier l'intersection
            pRelOp = CType(pGrilleValidation, IRelationalOperator)

            'Vérifier l'intersection
            If pRelOp.Disjoint(pGeometrie) Then
                'Créer une nouvelle ligne horizontale
                pLigneHorizontale = CreerLigneHorizontale(pPolygon, ((pGeometrie.Envelope.YMin + pGeometrie.Envelope.YMax) / 2))

                'Ajouter la nouvelle ligne à la grille de validation
                pGeometryColl.AddGeometry(pLigneHorizontale)
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryColl = Nothing
            pLigneHorizontale = Nothing
            pRelOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de créer et retourner une "Polyline" contenant une ligne verticale
    ''' de la grille de validation des courbes de niveau en fonction d'un polygon de découpage.
    ''' 
    ''' La ligne verticale de la grille de validation sont découpées en fonction du polygone de découpage.
    '''</summary>
    '''
    '''<param name="pPolygon">Interface contenant le polygone de découpage utilisé pour créer la ligne verticale de la grille.</param>
    '''<param name="dOrigineX">Valeur X d'origine de l'étendue à couvrir.</param>
    '''
    '''<returns>"IPolyline" contenant une ligne verticale de la grille de validation des courbes de niveau.</returns>
    ''' 
    Private Function CreerLigneVerticale(ByVal pPolygon As IPolygon, ByRef dOrigineX As Double) As IPolyline
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne de la grille de validation des courbes
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour découper une ligne selon le polygon de découpage
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour ajouter les sommets à une ligne de la grille
        Dim pPoint As IPoint = Nothing                      'Interface contenant un point d'une ligne de la grille de validation

        'Définir la valeur de retour par défaut
        CreerLigneVerticale = Nothing

        Try
            'Vérifier si l'origine est plus grande que le maximum
            If dOrigineX > pPolygon.Envelope.XMax Then
                'Définir l'origine comme étant celle du maximum
                dOrigineX = pPolygon.Envelope.XMax
            End If

            'Créer une nouvelle ligne correspondant à une intervalle
            pPolyline = New Polyline
            'Définir la référence spatiale de la nouvelle ligne
            pPolyline.SpatialReference = pPolygon.Envelope.SpatialReference
            'Interface pour ajouter des sommets à la nouvelle ligne
            pPointColl = CType(pPolyline, IPointCollection)

            'Créer un nouveau sommet vide
            pPoint = New Point
            'Définir la position du sommet inférieur
            pPoint.X = dOrigineX
            pPoint.Y = pPolygon.Envelope.YMin
            'Ajouter le sommet inférieur à la nouvelle ligne
            pPointColl.AddPoint(pPoint)

            'Créer un nouveau sommet vide
            pPoint = New Point
            'Définir la position du sommet inférieur
            pPoint.X = dOrigineX
            pPoint.Y = pPolygon.Envelope.YMax
            'Ajouter le sommet inférieur à la nouvelle ligne
            pPointColl.AddPoint(pPoint)

            'Découper la ligne de la grille selon le polygon de découpage
            pTopoOp = CType(pPolyline, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()
            pPolyline = CType(pTopoOp.Intersect(pPolygon, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

            'Retourner la ligne verticale
            CreerLigneVerticale = pPolyline

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
            pPoint = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de créer et retourner une "Polyline" contenant une ligne Horizontale
    ''' de la grille de validation des courbes de niveau en fonction d'un polygon de découpage.
    ''' 
    ''' La ligne Horizontale de la grille de validation sont découpées en fonction du polygone de découpage.
    '''</summary>
    '''
    '''<param name="pPolygon">Interface contenant le polygone de découpage utilisé pour créer la ligne Horizontale de la grille.</param>
    '''<param name="dOrigineY">Valeur Y d'origine de l'étendue à couvrir.</param>
    '''
    '''<returns>"IPolyline" contenant une ligne Horizontale de la grille de validation des courbes de niveau.</returns>
    ''' 
    Private Function CreerLigneHorizontale(ByVal pPolygon As IPolygon, ByRef dOrigineY As Double) As IPolyline
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne de la grille de validation des courbes
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour découper une ligne selon le polygon de découpage
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour ajouter les sommets à une ligne de la grille
        Dim pPoint As IPoint = Nothing                      'Interface contenant un point d'une ligne de la grille de validation

        'Définir la valeur de retour par défaut
        CreerLigneHorizontale = Nothing

        Try
            'Vérifier si l'origine est plus grande que le maximum
            If dOrigineY > pPolygon.Envelope.YMax Then
                'Définir l'origine comme étant celle du maximum
                dOrigineY = pPolygon.Envelope.YMax
            End If

            'Créer une nouvelle ligne correspondant à une intervalle
            pPolyline = New Polyline
            'Définir la référence spatiale de la nouvelle ligne
            pPolyline.SpatialReference = pPolygon.Envelope.SpatialReference
            'Interface pour ajouter des sommets à la nouvelle ligne
            pPointColl = CType(pPolyline, IPointCollection)

            'Créer un nouveau sommet vide
            pPoint = New Point
            'Définir la position du sommet inférieur
            pPoint.Y = dOrigineY
            pPoint.X = pPolygon.Envelope.XMin
            'Ajouter le sommet inférieur à la nouvelle ligne
            pPointColl.AddPoint(pPoint)

            'Créer un nouveau sommet vide
            pPoint = New Point
            'Définir la position du sommet inférieur
            pPoint.Y = dOrigineY
            pPoint.X = pPolygon.Envelope.XMax
            'Ajouter le sommet inférieur à la nouvelle ligne
            pPointColl.AddPoint(pPoint)

            'Découper la ligne de la grille selon le polygon de découpage
            pTopoOp = CType(pPolyline, ITopologicalOperator2)
            pPolyline = CType(pTopoOp.Intersect(pPolygon, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

            'Retourner la ligne verticale
            CreerLigneHorizontale = pPolyline

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
            pPoint = Nothing
        End Try
    End Function
#End Region
End Class
