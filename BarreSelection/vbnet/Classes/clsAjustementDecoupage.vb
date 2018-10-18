Imports System
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

'**
'Nom de la composante : clsAjustementDecoupage.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont l’ajustement des éléments à la limite d'un découpage respecte ou non
''' la continuité de ces derniers.
''' 
''' La classe permet de trouver les éléments qui respectent ou non:
''' -La distance de précision avec la limite de découpage.
''' -L'adjacence des éléments à la limite du décpupage.
''' -la correspondance des valeurs d'attributs spécifiés des éléments à la limite du découpage.
''' 
''' Note : Cette contrainte ne peut être traitée si aucun polygone de découpage n'est spécifié.     
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 22 Novembre 2016
'''</remarks>
''' 
Public Class clsAjustementDecoupage
    Inherits clsValeurAttribut

    '''<summary>Contient la position de l'attribut de découpage dans la classe de découpage.</summary>
    Protected giPosAttributDecoupage As Integer = -1
    '''<summary>Contient l'identifiant de l'élément de découpage.</summary>
    Protected gsIdentifiantDecoupage As String = "DATASET_NAME"
    '''<summary>Liste des attributs à traiter.</summary>
    Protected gsListeAttribut As String = "CODE_SPEC"
    ''' <summary>Interface ESRI contenant les limites communes de découpage avec les points d'adjacence.</summary>
    Protected gpLimiteDecoupageAvecPoint As IPolyline = Nothing
    ''' <summary>Objet contenant la liste des attributs d'adjacence à valider.</summary>
    Protected gpAttributAdjacence As Collection = Nothing
    ''' <summary>Interface ESRI contenant la liste des points d'adjacence.</summary>
    Protected gpListePointAdjacence As IGeometryCollection = Nothing
    ''' <summary>Objet contenant la liste des éléments à traiter.</summary>
    Protected gpListeElementTraiter As Collection = Nothing
    ''' <summary>Objet contenant la liste des éléments aux points d'adjacence.</summary>
    Protected gpListeElementPointAdjacent As Collection = Nothing
    ''' <summary>Objet contenant les points d'adjacence qui possède des erreurs.</summary>
    Protected gpListeErreurPointAdjacent As Collection = Nothing
    ''' <summary>Objet contenant les éléments d'erreurs de précision.</summary>
    Protected gpErreurFeaturePrecision As Collection = Nothing
    ''' <summary>Objet contenant les éléments d'erreurs d'adjacence.</summary>
    Protected gpErreurFeatureAdjacence As Collection = Nothing
    ''' <summary>Objet contenant les éléments d'erreurs d'attributs.</summary>
    Protected gpErreurFeatureAttribut As Collection = Nothing
    '''<summary>Interface contenant la référence spatiale projeter par défaut.</summary>
    Protected gpSpatialReferenceProj As ISpatialReference = Nothing

    '''<summary> Contient la tolérance d'adjacence utilisée pour corriger les éléments adjacents.</summary>
    Protected gdTolAdjacence As Double = 3.0
    '''<summary> Contient la tolérance d'adjacence utilisée pour corriger les éléments adjacents.</summary>
    Protected gdTolAdjacenceOri As Double = 3.0
    '''<summary> Contient la tolérance de recherche des éléments à traiter.</summary>
    Protected gdTolRecherche As Double = 1.0
    '''<summary> Contient la tolérance de recherche des éléments à traiter.</summary>
    Protected gdTolRechercheOri As Double = 1.0

    '''<summary> Indiquer si la correspondance aux points d'adjacence doit être unique.</summary>
    Protected gbAdjacenceUnique As Boolean = False
    '''<summary> Indiquer si on permet d'avoir des classes différentes entre les éléments adjacents.</summary>
    Protected gbClasseDifferente As Boolean = False
    '''<summary> Indiquer si on ne doit pas tenir compte des identifiants.</summary>
    Protected gbSansIdentifiant As Boolean = False

    ''' <summary>Structure contenant un élément à traiter.</summary>
    Public Structure Structure_Element_Traiter
        Dim OID As Integer
        Dim FeatureClass As IFeatureClass
    End Structure

    ''' <summary>Structure contenant les points d'adjacence.</summary>
    Public Structure Structure_Point_Adjacence
        Dim Point As IPoint
        Dim FeatureColl As Collection
    End Structure

    ''' <summary>Structure contenant les erreurs.</summary>
    Public Structure Structure_Erreur
        'Description de l'erreur
        Dim Description As String
        'Distance d'erreur
        Dim Distance As Double

        'Définition de l'information pour le point A
        Dim PointA As IPoint
        Dim FeatureA As IFeature
        Dim IdentifiantA As String
        Dim ValeurA As String
        Dim PosAttA As Integer

        'Définition de l'information pour le point B
        Dim PointB As IPoint
        Dim FeatureB As IFeature
        Dim IdentifiantB As String
        Dim ValeurB As String
        Dim PosAttB As Integer
    End Structure

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "DATASET_NAME"
            Expression = "CODE_SPEC 1.0 3.0"

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
    '''<param name="bAdjacenceMultiple"> Indique si on permet l'adjacence multiple.</param>
    '''
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String, ByVal sExpression As String,
                   Optional ByVal bAdjacenceMultiple As Boolean = False)
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
    '''</summary>
    ''' 
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    '''<param name="pFeatureDecoupage"> Élément de la classe de découpage à traiter.</param>
    ''' 
    Public Sub New(ByRef pFeatureLayerSelection As IFeatureLayer, ByVal sParametres As String, Optional ByVal pFeatureDecoupage As IFeature = Nothing)
        Try
            'Définir les valeurs par défaut
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres

            'Vérifier si l'élément de découpage est spécifié
            If pFeatureDecoupage IsNot Nothing Then
                'Sélection les éléments à la limite du découpage
                Call SelectionnerElementLimiteDecoupage(CType(pFeatureDecoupage.Shape, IPolygon))
            End If

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
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        giPosAttributDecoupage = Nothing

        gpLimiteDecoupageAvecPoint = Nothing
        gpAttributAdjacence = Nothing
        gpListePointAdjacence = Nothing
        gpListeElementTraiter = Nothing
        gpListeElementPointAdjacent = Nothing
        gpListeErreurPointAdjacent = Nothing
        gpErreurFeaturePrecision = Nothing
        gpErreurFeatureAdjacence = Nothing
        gpErreurFeatureAttribut = Nothing
        gpSpatialReferenceProj = Nothing

        gsIdentifiantDecoupage = Nothing
        gsListeAttribut = Nothing

        gdTolRecherche = Nothing
        gdTolAdjacence = Nothing
        gdTolRechercheOri = Nothing
        gdTolAdjacenceOri = Nothing

        gbAdjacenceUnique = Nothing
        gbClasseDifferente = Nothing
        gbSansIdentifiant = Nothing

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
            Nom = "AjustementDecoupage"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la ligne de paramètre à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides Property Parametres() As String
        Get
            'Retourner la valeur des paramètres
            Parametres = gsNomAttribut & " " & gsExpression
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
                gsNomAttribut = params(0)
                gsExpression = params(1)

                'Traiter tous les paramètres de l'expression
                For i = 2 To params.Length - 1
                    'Définir l'expression
                    gsExpression = gsExpression & " " & params(i)
                Next

                'Définir la liste des paramètres de l'expression
                params = gsExpression.Split(CChar(" "))
                'Définir la liste des attributs
                gsListeAttribut = params(0)
                'Définir la tolérance de recherche
                If params.Length > 1 Then
                    'Convertir la tolérance de recherche
                    If IsNumeric(params(1)) Then gdTolRecherche = CDbl(params(1))
                    gdTolRechercheOri = gdTolRecherche
                End If
                'Définir la tolérance d'adjacence
                If params.Length > 2 Then
                    'Convertir la tolérance d'adjacence
                    If IsNumeric(params(2)) Then gdTolAdjacence = CDbl(params(2))
                    gdTolAdjacenceOri = gdTolAdjacence
                End If

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
        'Déclarer les variables de travail
        Dim pFields As IFields = Nothing            'Interface contenant les attributs de la FeatureClass.
        Dim sAttributs As String = ""               'Liste des attributs de la Featureclass de sélection.

        Try
            'Définir la liste des paramètres par défaut
            ListeParametres = New Collection

            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Définir l'interface contenant les attributs de la Featureclass
                pFields = gpFeatureLayerSelection.FeatureClass.Fields

                'Traiter tous les attributs
                For i = 0 To pFields.FieldCount - 1
                    'Vérifier si l'attribut est éditable et ne correspond pas à la géométrie
                    If pFields.Field(i).Editable And pFields.Field(i).Name <> gsNomAttribut _
                    And pFields.Field(i).Name <> gpFeatureLayerSelection.FeatureClass.ShapeFieldName Then
                        'Vérifier si c'est le premier attribut
                        If sAttributs = "" Then
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributs = pFields.Field(i).Name
                        Else
                            'Ajouter le nom de l'attribut dans la liste
                            sAttributs = sAttributs & "," & pFields.Field(i).Name
                        End If
                    End If
                Next

                'Définir l'attribut de découpage et valider la valeur de l'attribut CODE_SPEC
                ListeParametres.Add(gsNomAttribut & " CODE_SPEC 1.0 3.0")

                'Définir l'attribut de découpage, valider la valeur de l'attribut CODE_SPEC et seul la correspondance UNIQUE aux points d'adjacence est permise
                ListeParametres.Add(gsNomAttribut & " CODE_SPEC 1.0 3.0 UNIQUE")

                'Définir l'attribut de découpage, valider la valeur de l'attribut CODE_SPEC et ne pas tenir compte des identifiants
                ListeParametres.Add(gsNomAttribut & " CODE_SPEC 1.0 3.0 SANS_ID")

                'Définir l'attribut de découpage et valider la valeur de tous les attributs éditable
                ListeParametres.Add(gsNomAttribut & " " & sAttributs & " 1.0 3.0")

                'Définir l'attribut de découpage, valider la valeur de tous les attributs éditable, 
                'seul la correspondance UNIQUE aux points d'adjacence est permise et ne pas tenir compte des identifiants
                ListeParametres.Add(gsNomAttribut & " " & sAttributs & " 1.0 3.0 UNIQUE")

                'Définir l'attribut de découpage, valider la valeur de tous les attributs éditable et ne pas tenir compte des identifiants
                ListeParametres.Add(gsNomAttribut & " " & sAttributs & " 1.0 3.0 SANS_ID")

                'Définir l'attribut de découpage, valider la valeur de tous les attributs éditable, 
                'seul la correspondance UNIQUE aux points d'adjacence est permise et ne pas tenir compte des identifiants
                ListeParametres.Add(gsNomAttribut & " " & sAttributs & " 1.0 3.0 UNIQUE SANS_ID")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFields = Nothing
            sAttributs = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si la FeatureClass est valide.
    '''</summary>
    '''
    '''<return>Boolean qui indique si la FeatureClass est valide.</return>
    ''' 
    Public Overloads Overrides Function FeatureClassValide() As Boolean
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface pour vérifier la sélection.

        Try
            'Définir la valeur par défaut, la contrainte est valide.
            FeatureClassValide = True
            gsMessage = "La contrainte est valide."

            'Vérifier si la FeatureClass est valide
            If gpFeatureLayerSelection.FeatureClass IsNot Nothing Then
                'Vérifier si la FeatureClass est de type Point, MultiPoint, Polyline ou Polygon
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'Vérifier si aucun élément de découpage n'est spécifié
                    If gpFeatureDecoupage Is Nothing Then
                        'Vérifier si le FeatureLayer est valide
                        If gpFeatureLayerDecoupage IsNot Nothing Then
                            'Vérifier si la FeatureClass est valide
                            If gpFeatureLayerDecoupage.FeatureClass IsNot Nothing Then
                                'Vérifier si la FeatureClass est de type Polygon
                                If gpFeatureLayerDecoupage.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                                    'Interface pour vérifier la sélection
                                    pFeatureSel = CType(gpFeatureLayerDecoupage, IFeatureSelection)
                                    'Vérifier si aucun élément de découpage n'est sélectionné
                                    If pFeatureSel.SelectionSet.Count = 0 Then
                                        'Définir le message d'erreur
                                        FeatureClassValide = False
                                        gsMessage = "ERREUR : Aucun élément de découpage n'est sélectionné dans la classe de découpage."
                                    End If

                                    'Si le type de la classe de découpage est invalide
                                Else
                                    'Définir le message d'erreur
                                    FeatureClassValide = False
                                    gsMessage = "ERREUR : La classe de découpage n'est pas de type Polygon."
                                End If

                                'Si la classe de découpage est invalide
                            Else
                                'Définir le message d'erreur
                                FeatureClassValide = False
                                gsMessage = "ERREUR : La classe de découpage est invalide."
                            End If

                            'Si le Layer de découpage est invalide
                        Else
                            'Définir le message d'erreur
                            FeatureClassValide = False
                            gsMessage = "ERREUR : Le Layer de découpage n'est pas spécifié."
                        End If
                    End If

                    'Si la classe de sélection est invalide
                Else
                    'Définir le message d'erreur
                    FeatureClassValide = False
                    gsMessage = "ERREUR : La classe de sélection n'est pas de type Point, MultiPoint, Polyline ou Polygon."
                End If

                'Si la classe de sélection est invalide
            Else
                'Définir le message d'erreur
                FeatureClassValide = False
                gsMessage = "ERREUR : La classe de sélection est invalide."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si l'attribut est valide.
    ''' 
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overridable Overloads Function AttributValide() As Boolean
        'déclarer les variables de travail
        Dim iPosAtt As Integer = -1     'Contient la position de l'attribut

        Try
            'Définir la valeur par défaut, la contrainte est valide.
            AttributValide = True
            gsMessage = "La contrainte est valide."

            'Définir la position de l'attribut dans la classe de sélection
            giAttribut = gpFeatureLayerSelection.FeatureClass.Fields.FindField(gsNomAttribut)

            'Vérifier si l'attribut est invalide
            If giAttribut = -1 Then
                'La contrainte est invalide
                AttributValide = False
                gsMessage = "ERREUR : L'attribut de la classe de sélection est invalide : " & gsNomAttribut
            End If

            'Définir la position de l'attribut dans la classe de découpage
            iPosAtt = gpFeatureLayerDecoupage.FeatureClass.Fields.FindField(gsNomAttribut)

            'Vérifier si l'attribut est invalide
            If iPosAtt = -1 Then
                'La contrainte est invalide
                AttributValide = False
                gsMessage = "ERREUR : L'attribut de la classe de découpage est invalide : " & gsNomAttribut
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            iPosAtt = Nothing
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
        Dim sParams() As String = Nothing       'Contient la liste des paramètres de l'expression.
        Dim sNomAttribut As String = ""         'Contient le nom de l'attribut.
        Dim sListeAttribut As String = ""       'Contient la liste des attributs à valider.
        Dim sTolRecherche As String = ""        'Contient la tolérance de recherche.
        Dim sTolAdjacence As String = ""        'Contient la tolérance d'adjacence.

        'Définir la valeur par défaut
        ExpressionValide = True
        gsMessage = "L'expression est valide"

        Try
            'Vérifier si l'expression est vide
            If Len(gsExpression) = 0 Then
                'Définir le message d'erreur
                gsMessage = "ERREUR : L'expression n'est pas spécifiée!"
                Return False
            End If

            'Définir la liste des paramètres de l'expression
            sParams = gsExpression.Split(CChar(" "))
            'Définir la liste des attributs
            sListeAttribut = sParams(0)
            'Définir la tolérance de recherche
            If sParams.Length > 1 Then sTolRecherche = sParams(1)
            'Définir la tolérance d'adjacence
            If sParams.Length > 2 Then sTolAdjacence = sParams(2)

            'Créer une nouvelle liste d'attribut d'adjacence
            gpAttributAdjacence = New Collection

            'Valider tous les attributs de la liste
            For Each sNomAttribut In sListeAttribut.Split(CChar(","))
                'Vérifier si l'attribut est invalide
                If gpFeatureLayerSelection.FeatureClass.Fields.FindField(sNomAttribut) = -1 Then
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : L'attribut est invalide : " & sNomAttribut
                    Return False
                End If
                'Ajouter l'attribut dans la liste
                gpAttributAdjacence.Add(sNomAttribut)
            Next

            'Vérifier si la tolérance de recherche n'est pas numérique
            If Not IsNumeric(sTolRecherche) Then
                'Définir le message d'erreur
                gsMessage = "ERREUR : La tolérance de recherche n'est pas numérique : " & sTolRecherche.ToString
                Return False
            End If

            'Vérifier si la tolérance de recherche n'est pas valide
            If CDbl(sTolRecherche) <= 0 Then
                'Définir le message d'erreur
                gsMessage = "ERREUR : La tolérance de recherche n'est pas valide : " & sTolRecherche.ToString
                Return False
            End If

            'Vérifier si la tolérance d'adjacence n'est pas numérique
            If Not IsNumeric(sTolAdjacence) Then
                'Définir le message d'erreur
                gsMessage = "ERREUR : La tolérance d'adjacence n'est pas numérique : " & sTolAdjacence.ToString
                Return False
            End If

            'Vérifier si la tolérance d'adjacence n'est pas valide
            If CDbl(sTolAdjacence) < CDbl(sTolRecherche) Then
                'Définir le message d'erreur
                gsMessage = "ERREUR : La tolérance d'adjacence n'est pas valide : " & sTolAdjacence.ToString
                Return False
            End If

            'Définir les paramètres de l'expression
            gsListeAttribut = sListeAttribut
            gdTolRecherche = CDbl(sTolRecherche)
            gdTolRechercheOri = gdTolRecherche
            gdTolAdjacence = CDbl(sTolAdjacence)
            gdTolAdjacenceOri = gdTolAdjacence
            gbAdjacenceUnique = gsExpression.ToUpper.Contains("UNIQUE")
            gbSansIdentifiant = gsExpression.ToUpper.Contains("SANS_ID")

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            sParams = Nothing
            sNomAttribut = Nothing
            sListeAttribut = Nothing
            sTolRecherche = Nothing
            sTolAdjacence = Nothing
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
        Dim pFeatureSel As IFeatureSelection = Nothing          'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.

        Try
            'Sortir si la contrainte est invalide
            If Me.EstValide() = False Then Err.Raise(1, , Me.Message)

            'Définir la géométrie par défaut
            Selectionner = New GeometryBag
            Selectionner.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Conserver le type de sélection à effectuer
            gbEnleverSelection = bEnleverSelection

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

            'Traiter le FeatureLayer selon pour l'ajustement d'un découpage spécifique
            Selectionner = TraiterAjustementDecoupage(pTrackCancel, bEnleverSelection)

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
            pSelectionSet = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont l'ajustement des éléments à la limite du découpage
    ''' respecte ou non les paramètres spécifiés.
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
    '''<return>Les géométries des éléments qui respectent ou non les relations spatiales.</return>
    ''' 
    Private Function TraiterAjustementDecoupage(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.

        Try
            'Définir la géométrie par défaut
            TraiterAjustementDecoupage = New GeometryBag
            TraiterAjustementDecoupage.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterAjustementDecoupage, IGeometryCollection)

            'Définir la limite de découpage à traiter à partir des éléments de découpage sélectionnés
            Call DefinirLimiteDecoupage(pTrackCancel)

            'Identifier tous les points d'adjacence
            Call IdentifierAdjacenceDecoupage(pTrackCancel)

            'Vérifier si on doit écrire les erreurs d'ajustement invalide
            If gbEnleverSelection Then
                'Écrire les erreurs de précision
                Call EcrireErreurPrecision(pGeomResColl)

                'Écrire les erreurs d'adjacence
                Call EcrireErreurAdjacence(pGeomResColl)

                'Écrire les erreurs d'attribut
                Call EcrireErreurAttribut(pGeomResColl)
            End If

            'Sélectionner les élément selon le type de sélection
            Call SelectionnerElement(pTrackCancel, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de définir les limites communes entre les éléments de la classe de découpage sélectionnés. 
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    '''
    Public Sub DefinirLimiteDecoupage(ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour vérifier la présence des éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface utilisé pour extraire les éléments sélectionnés.
        Dim pCursor As ICursor = Nothing                'Interface utilisé pour extraire les éléments sélectionnés.
        Dim pFeature As IFeature = Nothing              'Interface contenant un élément sélectionné.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter des géométries.
        Dim pGeomBag As IGeometryCollection = Nothing   'Interface contenant toutes les géométries des éléments.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface utilisé pour pour calculer l'intersection des limites commune.
        Dim i As Integer = Nothing                      'Compteur
        Dim j As Integer = Nothing                      'Compteur

        Try
            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Définition de la limite du découpage en cours ..."

            'Interface pour extraire les éléments sélectionnés
            pFeatureSel = CType(gpFeatureLayerDecoupage, IFeatureSelection)

            'Vérifier si un élément de découpage est spécifié
            If gpFeatureDecoupage IsNot Nothing Then
                'Transformer les tolérances en géographique au besoin
                pGeometry = CType(gpFeatureDecoupage.Shape, IGeometry)
                Call TransformerTolerances(pGeometry.Envelope)

                'Interface pour extraire la limite de la géométrie
                pTopoOp = CType(pGeometry, ITopologicalOperator2)

                'Définir la limite de découpage
                gpLimiteDecoupage = CType(pTopoOp.Boundary, IPolyline)

                'Vérifier le nombre d'éléments sélectionné
            ElseIf pFeatureSel.SelectionSet.Count > 1 Then
                'Définir un nouveau Bag
                pGeometry = New GeometryBag
                pGeometry.SpatialReference = gpSpatialReference
                pGeomBag = CType(pGeometry, IGeometryCollection)

                'Extraire les éléments sélectionnés
                pFeatureSel.SelectionSet.Search(Nothing, True, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature

                'Traiter tous les éléments
                Do Until pFeature Is Nothing
                    'Définir la géométrie de découpage
                    pGeometry = pFeature.ShapeCopy
                    'Projeter la géométrie
                    pGeometry.Project(gpSpatialReference)
                    'Ajouter la géométrie dans la Bag
                    pGeomBag.AddGeometry(pGeometry)

                    'Extraire le premier élément
                    pFeature = pFeatureCursor.NextFeature
                Loop

                'Transformer les tolérances en géographique au besoin
                pGeometry = CType(pGeomBag, IGeometry)
                Call TransformerTolerances(pGeometry.Envelope)

                'Définir la géométrie contenant les limites de découpage
                gpLimiteDecoupage = CType(New Polyline, IPolyline)
                gpLimiteDecoupage.SpatialReference = gpSpatialReference
                'Interface pour ajouter les limites de découpage
                pGeomColl = CType(gpLimiteDecoupage, IGeometryCollection)

                'Traiter toutes les géométries contenues dans le Bag
                For i = 0 To pGeomBag.GeometryCount - 2
                    'Interface pour trouver l'intersection
                    pTopoOp = CType(pGeomBag.Geometry(i), ITopologicalOperator2)
                    'Traiter toutes les autres géométries contenues dans le Bag
                    For j = i + 1 To pGeomBag.GeometryCount - 1
                        'Trouver l'intersection
                        pGeometry = pTopoOp.Intersect(pGeomBag.Geometry(j), esriGeometryDimension.esriGeometry1Dimension)
                        'Vérifier si une intersection a été trouvée
                        If Not pGeometry.IsEmpty Then
                            'Ajouter les limites communes
                            pGeomColl.AddGeometryCollection(CType(pGeometry, IGeometryCollection))
                        End If
                    Next
                Next

                'Vérifier si la limite commune est présente
                If gpLimiteDecoupage.IsEmpty Then
                    'Afficher un message
                    Throw New Exception("ERREUR : Aucune limite de découpage commune!")
                Else
                    'Interface pour simplifier les limites communes
                    pTopoOp = CType(gpLimiteDecoupage, ITopologicalOperator2)
                    'Simplifier les limites communes
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()
                End If

                'Si seulement un éléments de découpage
            ElseIf pFeatureSel.SelectionSet.Count = 1 Then
                'Extraire les éléments sélectionnés
                pFeatureSel.SelectionSet.Search(Nothing, True, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature

                'Transformer les tolérances en géographique au besoin
                pGeometry = CType(pFeature.Shape, IGeometry)
                Call TransformerTolerances(pGeometry.Envelope)

                'Interface pour extraire la limite de la géométrie
                pTopoOp = CType(pGeometry, ITopologicalOperator2)

                'Définir la limite de découpage
                gpLimiteDecoupage = CType(pTopoOp.Boundary, IPolyline)

                'Si aucun élément de découpage n'est trouvé
            Else
                'Afficher un message
                Throw New Exception("ERREUR : Aucun élément de découpage n'est sélectionné!")
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pGeomBag = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier et valider tous les points d'adjacence aux limites communes de découpage 
    ''' qui sont présents dans la fenêtre graphique.
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    ''' 
    ''' <remarks>
    ''' À la fin du traitement :
    ''' -Les points d'adjacence sont identifiés.
    ''' -Les erreurs de précision sont identifiées.
    ''' -Les erreurs d'adjacence sont identifiées.
    ''' -Les erreurs d'attribut sont identifiées.
    ''' </remarks>
    ''' 
    Public Sub IdentifierAdjacenceDecoupage(ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing            'Interface contenant une géométrie.
        Dim bSelection As Boolean = True                'Indice de sélection des éléments de la classe de découpage.

        Try
            'Définir la nouvelle liste des éléments pour chaque point d'adjacence
            gpListeElementPointAdjacent = New Collection
            'Définir la nouvelle liste des points d'adjacence en erreur
            gpListeErreurPointAdjacent = New Collection
            'Définir une nouvelle Collection pour les éléments en erreur
            gpErreurFeaturePrecision = New Collection
            gpErreurFeatureAdjacence = New Collection
            gpErreurFeatureAttribut = New Collection

            'Définir un nouveau Bag pour les points d'adjacence
            pGeometry = New GeometryBag
            pGeometry.SpatialReference = gpSpatialReference
            'Interface pour ajouter les points adjacents
            gpListePointAdjacence = CType(pGeometry, IGeometryCollection)

            'Désactiver la sélection du layer de découpage
            bSelection = gpFeatureLayerDecoupage.Selectable
            gpFeatureLayerDecoupage.Selectable = False

            'Définition des points d'adjacence virtuels
            Call DefinirPointAdjacenceVirtuel(pTrackCancel)

            'Trouver les points d'adjacence et les éléments qui y sont liés
            Call TrouverListePointAdjacence(pTrackCancel)

            'Valider la liste des points d'adjacence
            Call ValiderDistanceListePointAdjacence(pTrackCancel)

            'Valider la liste des points d'adjacence
            Call ValiderAdjacenceListePointAdjacence(pTrackCancel)

            'Trouver l'élément adjacent pour tous les éléments isolés, lorsque possible
            'Call TrouverElementAdjacent(pTrackCancel)

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Remettre la sélection du layer de découpage
            gpFeatureLayerDecoupage.Selectable = bSelection
            'Vider la mémoire
            pGeometry = Nothing
            bSelection = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de définir les points d'adjacence virtuels.
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    ''' 
    Public Sub DefinirPointAdjacenceVirtuel(ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pAdjColl As IPointCollection = Nothing      'Interface ESRI utilisé pour extraire un sommet d'adjacence.
        Dim pCloneLimite As IClone = Nothing            'Interface ESRI qui permet de cloner une géométrie.
        Dim pPointIdAware As IPointIDAware = Nothing    'Interface ESRI pour gérer les IDs de point.
        Dim pPoint As IPoint = Nothing                  'Interface ESRI contenant le sommet traité.
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Définition des points d'adjacence virtuels ..."

            'Projeter les limites communes selon la projection de travail
            gpLimiteDecoupage.Project(gpSpatialReference)

            'Cloner la limite de découpage
            pCloneLimite = CType(gpLimiteDecoupage, IClone)
            gpLimiteDecoupageAvecPoint = CType(pCloneLimite.Clone, IPolyline)

            'Interface pour traiter les sommets
            pAdjColl = CType(gpLimiteDecoupageAvecPoint, IPointCollection)
            'Permettre d'ajouter des identifiant de point
            pPointIdAware = CType(pAdjColl, IPointIDAware)
            pPointIdAware.PointIDAware = True
            'Traiter tous les Ids de point
            For i = 0 To pAdjColl.PointCount - 1
                'Initialiser l'identifiant du point virtuel
                pPoint = pAdjColl.Point(i)
                pPoint.ID = 0
                pAdjColl.UpdatePoint(i, pPoint)
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pAdjColl = Nothing
            pCloneLimite = Nothing
            pPointIdAware = Nothing
            pPoint = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de trouver tous tous les points d'adjacence à la limite du découpage.
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    ''' 
    ''' <remarks>
    ''' À la fin du traitement :
    ''' -Les points d'adjacence sont identifiés ainsi que les éléments qui y sont liés.
    ''' </remarks>
    ''' 
    Public Sub TrouverListePointAdjacence(ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pSpatialFilter As ISpatialFilter = Nothing  'Interface utiliser pour extraire les éléments à la limites du découpage.
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour sélectionner les éléments à la limite du découpage.
        Dim pCursor As ICursor = Nothing                'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface utilisé pour extraire les éléments à traiter.
        Dim qElementTraiter As Structure_Element_Traiter = Nothing 'Contient un élément à traiter.
        Dim pFeature As IFeature = Nothing              'Interface ESRI contenant un élément sélectionné.
        Dim pGeometry As IGeometry = Nothing            'Interface ESRI contenant la géométrie d'un élément.
        Dim pAdjacence As IGeometry = Nothing           'Interface ESRI contenant la géométrie d'adjacence.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface ESRI utilisé pour ajouter des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface ESRI utilisé pour pour calculer la topologie.
        Dim pBufferLimite As IGeometry = Nothing        'Interface contenant le buffer de la limite de découpage.
        Dim pPointColl As IPointCollection = Nothing    'Interface ESRI utilisé pour traiter tous les sommets d'adjacence
        Dim pHitTest As IHitTest = Nothing              'Interface ESRI pour tester la présence du sommet recherché
        Dim pPoint As IPoint = Nothing                  'Interface ESRI contenant le sommet traité
        Dim sIdentifiant As String = Nothing            'Contient l'identifiant de découpage de l'élément
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments à la limite du découpage en cours ..."

            'Créer la nouvelle collection des éléments à traiter
            gpListeElementTraiter = New Collection

            'Interface utilisé pour extraire le point correspondant exactement aux limites communes
            pHitTest = CType(gpLimiteDecoupageAvecPoint, IHitTest)

            'Interface qui permet d'extraire la composante trouvée
            pGeomColl = CType(gpLimiteDecoupageAvecPoint, IGeometryCollection)

            'Interface pour créer le buffer
            pTopoOp = CType(gpLimiteDecoupageAvecPoint, ITopologicalOperator2)
            'Créer le buffer de la limite
            pBufferLimite = pTopoOp.Buffer(gdTolRecherche)
            'Simplifier le buffer de la limite
            pTopoOp = CType(pBufferLimite, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Définir la requête spatiale
            pSpatialFilter = New SpatialFilter
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            pSpatialFilter.OutputSpatialReference(gpFeatureLayerSelection.FeatureClass.ShapeFieldName) = pBufferLimite.SpatialReference
            pSpatialFilter.Geometry = pBufferLimite

            'Interface pour sélectionner les éléments à la limite du découpage
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            'Vérifier si une sélection est présente
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionner tous les éléments à la limite du découpage
                pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            Else
                'Sélectionner tous les éléments déjà sélectionnés à la limite du découpage
                pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultAnd, False)
            End If

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Trouver la liste des points d'adjacence en cours ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pFeatureSel.SelectionSet.Count, pTrackCancel)

            'Interfaces pour extraire les éléments sélectionnés
            pFeatureSel.SelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments
            Do Until pFeature Is Nothing
                'Créer un lien vers un élément à traiter
                qElementTraiter = New Structure_Element_Traiter
                qElementTraiter.OID = pFeature.OID
                qElementTraiter.FeatureClass = CType(pFeature.Class, IFeatureClass)
                'Ajouter l'élément dans la liste à traiter avec un numéro d'index
                gpListeElementTraiter.Add(qElementTraiter, (gpListeElementTraiter.Count + 1).ToString)

                'Définir la valeur de l'identifiant
                sIdentifiant = fsIdentifiantDecoupage(pFeature, gsIdentifiantDecoupage)

                'Interface contenant la géométrie de l'élément
                pGeometry = pFeature.ShapeCopy
                'Projeter la géométrie
                pGeometry.Project(pBufferLimite.SpatialReference)

                'Vérifier si la géométrie est une surface
                If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Convertir la géométrie en multipoint
                    pAdjacence = GeometrieToMultiPoint(pGeometry)
                    'Simplifier la géométrie
                    pTopoOp = CType(pAdjacence, ITopologicalOperator2)
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()
                    'Extraire la géométrie d'adjacence de l'élément
                    pAdjacence = pTopoOp.Intersect(pBufferLimite, esriGeometryDimension.esriGeometry0Dimension)
                    'Interface pour traiter tous les sommets
                    pPointColl = CType(pAdjacence, IPointCollection)
                    'Traiter tous les sommets
                    For i = 0 To pPointColl.PointCount - 1
                        'Interface contenant le point
                        pPoint = pPointColl.Point(i)
                        'Identifier un point d'adjacence
                        Call IdentifierPointAdjacence(sIdentifiant, pPoint, pFeature, pHitTest, pGeomColl)
                    Next

                    'Vérifier si la géométrie est une ligne
                ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Convertir la géométrie en multipoint
                    pAdjacence = GeometrieToMultiPoint(pGeometry)
                    'Simplifier la géométrie
                    pTopoOp = CType(pAdjacence, ITopologicalOperator2)
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()
                    'Extraire la géométrie d'adjacence de l'élément
                    pAdjacence = pTopoOp.Intersect(pBufferLimite, esriGeometryDimension.esriGeometry0Dimension)
                    'Interface pour traiter tous les sommets
                    pPointColl = CType(pAdjacence, IPointCollection)
                    'Traiter tous les sommets
                    For i = 0 To pPointColl.PointCount - 1
                        'Interface contenant le point
                        pPoint = pPointColl.Point(i)
                        'Identifier un point d'adjacence
                        Call IdentifierPointAdjacence(sIdentifiant, pPoint, pFeature, pHitTest, pGeomColl)
                    Next

                    'Vérifier si la géométrie est un point multiple
                ElseIf pFeature.Shape.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                    'Simplifier la géométrie
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()
                    'Extraire la géométrie d'adjacence de l'élément
                    pAdjacence = pTopoOp.Intersect(pBufferLimite, esriGeometryDimension.esriGeometry0Dimension)
                    'Interface pour traiter tous les sommets
                    pPointColl = CType(pAdjacence, IPointCollection)
                    'Traiter tous les sommets
                    For i = 0 To pPointColl.PointCount - 1
                        'Interface contenant le point
                        pPoint = pPointColl.Point(i)
                        'Identifier un point d'adjacence
                        Call IdentifierPointAdjacence(sIdentifiant, pPoint, pFeature, pHitTest, pGeomColl)
                    Next

                    'Vérifier si la géométrie est un point
                ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                    'Interface contenant le point
                    pPoint = CType(pGeometry, IPoint)
                    'Identifier un point d'adjacence
                    Call IdentifierPointAdjacence(sIdentifiant, pPoint, pFeature, pHitTest, pGeomColl)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()
            Loop

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pSpatialFilter = Nothing
            pFeatureSel = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            qElementTraiter = Nothing
            pBufferLimite = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pAdjacence = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
            pHitTest = Nothing
            pPointColl = Nothing
            pPoint = Nothing
            sIdentifiant = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier un point d'adjacence spécifique par rapport aux limites communes de découpage
    ''' et de valider la précision du point d'adjacence identifié.
    '''</summary>
    '''
    '''<param name="sIdentifiant">Contient l'identifiant de découpage de l'élément.</param>
    '''<param name="pPoint">Interface ESRI contenant le sommet traité.</param>
    '''<param name="pFeature">Interface ESRI contenant un élément sélectionné.</param>
    '''<param name="pHitTest">Interface ESRI pour tester la présence du sommet recherché.</param>
    '''<param name="pGeomColl">Interface ESRI utilisé pour ajouter des géométries.</param>
    ''' 
    Private Sub IdentifierPointAdjacence(ByRef sIdentifiant As String, ByRef pPoint As IPoint, ByRef pFeature As IFeature, _
                                         ByRef pHitTest As IHitTest, ByRef pGeomColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim oErreur As Structure_Erreur = Nothing   'Structure contenant une erreur de précision
        Dim qElementColl As Collection = Nothing    'Objet qui contient la liste des éléments adjacents
        Dim pAdjColl As IPointCollection = Nothing  'Interface ESRI utilisé pour extraire un sommet d'adjacence
        Dim pNewPoint As IPoint = Nothing           'Interface contenant le sommet trouvé
        Dim pProxOp As IProximityOperator = Nothing 'Interface pour calculer la distance
        Dim dDistance As Double = Nothing           'Interface contenant la distance calculée entre le point de recherche et le sommet trouvé
        Dim nNumeroPartie As Integer = Nothing      'Numéro de partie trouvée
        Dim nNumeroSommet As Integer = Nothing      'Numéro de sommet de la partie trouvée
        Dim bCoteDroit As Boolean = Nothing         'Indiquer si le point trouvé est du côté droit de la géométrie
        Dim dTolerance As Double = Nothing          'Contient la tolérance de recherche

        Try
            'Créer un nouveau point vide
            pNewPoint = New Point
            pNewPoint.SpatialReference = pPoint.SpatialReference

            'Vérifier si seulement la correspondance UNIQUE aux points d'adjacence est permise
            If gbAdjacenceUnique Then
                'Définir la tolérance selon celle de recherche
                dTolerance = gdTolRecherche
                'Si l'adjacence multiple est permise
            Else
                'Définir la tolérance selon celle d'adjacence
                dTolerance = gdTolAdjacence
            End If

            'Vérifier la présence d'un sommet existant selon une tolérance
            If pHitTest.HitTest(pPoint, dTolerance, esriGeometryHitPartType.esriGeometryPartVertex, _
                                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                'Définir la référence spatiale du point trouvé
                pNewPoint.SpatialReference = pPoint.SpatialReference
                'Vérifier si c'est un point non identifié
                If pNewPoint.ID = 0 Then
                    'Ajouter la géométrie d'adjacence
                    gpListePointAdjacence.AddGeometry(pNewPoint)
                    'Définir le ID du sommet
                    pNewPoint.ID = gpListePointAdjacence.GeometryCount * -1
                    'Mettre à jour le ID du point
                    pAdjColl = CType(pGeomColl.Geometry(nNumeroPartie), IPointCollection)
                    pAdjColl.UpdatePoint(nNumeroSommet, pNewPoint)
                    'Ajouter l'élément adjacent
                    qElementColl = New Collection
                    qElementColl.Add(gpListeElementTraiter.Count)
                    'Ajouter le point d'adjacence
                    gpListeElementPointAdjacent.Add(qElementColl, pNewPoint.ID.ToString)
                Else
                    'Définir la liste des éléments adjacents
                    qElementColl = CType(gpListeElementPointAdjacent.Item(pNewPoint.ID.ToString), Collection)
                    qElementColl.Add(gpListeElementTraiter.Count)
                End If

                'Vérifier si la distance est plus grande que la précision
                If dDistance > gdPrecision Then
                    'Ajouter une erreur de précision
                    pPoint.ID = pNewPoint.ID

                    'Projeter les points
                    pPoint.Project(gpSpatialReferenceProj)
                    pNewPoint.Project(gpSpatialReferenceProj)
                    'Calculer la distance
                    pProxOp = CType(pPoint, IProximityOperator)
                    dDistance = pProxOp.ReturnDistance(pNewPoint)

                    'Vérifier si la distance est plus grande que la tolérance de recherche
                    If dDistance > gdTolRechercheOri Then
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + gdTolAdjacenceOri.ToString("0.0##") + ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.FeatureA = pFeature
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        gpErreurFeaturePrecision.Add(oErreur)

                    Else
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Ne touche pas à un point existant (" & dDistance.ToString("0.0##") & "<=" & gdTolRechercheOri.ToString("0.0##") & ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.FeatureA = pFeature
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        gpErreurFeaturePrecision.Add(oErreur)
                    End If

                    'Conserver l'dentifiant du point d'adjacence en erreur
                    If Not gpListeErreurPointAdjacent.Contains(pNewPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pNewPoint.ID, pNewPoint.ID.ToString)
                End If

                'Vérifier la présence d'un sommet sur une droite selon une tolérance
            ElseIf pHitTest.HitTest(pPoint, gdTolRecherche, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                'Définir la référence spatiale du point trouvé
                pNewPoint.SpatialReference = pPoint.SpatialReference
                'Ajouter la géométrie d'adjacence
                gpListePointAdjacence.AddGeometry(pNewPoint)
                'Interface pour ajouter un sommet
                pAdjColl = CType(pGeomColl.Geometry(nNumeroPartie), IPointCollection)
                'Définir le ID du sommet
                pNewPoint.ID = gpListePointAdjacence.GeometryCount
                'Insérer un nouveau sommet
                pAdjColl.InsertPoints(nNumeroSommet + 1, 1, pNewPoint)
                'Ajouter l'élément adjacent
                qElementColl = New Collection
                qElementColl.Add(gpListeElementTraiter.Count)
                'Ajouter le point d'adjacence
                gpListeElementPointAdjacent.Add(qElementColl, pNewPoint.ID.ToString)

                'Vérifier si la distance est plus grande que la précision
                If dDistance > gdPrecision Then
                    'Ajouter une erreur de précision
                    pPoint.ID = pNewPoint.ID

                    'Projeter les points
                    pPoint.Project(gpSpatialReferenceProj)
                    pNewPoint.Project(gpSpatialReferenceProj)
                    'Calculer la distance
                    pProxOp = CType(pPoint, IProximityOperator)
                    dDistance = pProxOp.ReturnDistance(pNewPoint)

                    'Vérifier si la distance est plus grande que la tolérance de recherche
                    If dDistance > gdTolRechercheOri Then
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + gdTolAdjacenceOri.ToString("0.0#######") + ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.FeatureA = pFeature
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        gpErreurFeaturePrecision.Add(oErreur)

                    Else
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Ne touche pas à la limite (" & dDistance.ToString("0.0##") & "<=" & gdTolRechercheOri.ToString("0.0#######") & ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.FeatureA = pFeature
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        gpErreurFeaturePrecision.Add(oErreur)
                    End If

                    'Conserver l'dentifiant du point d'adjacence en erreur
                    If Not gpListeErreurPointAdjacent.Contains(pNewPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pNewPoint.ID, pNewPoint.ID.ToString)
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            oErreur = Nothing
            qElementColl = Nothing
            pAdjColl = Nothing
            pNewPoint = Nothing
            pProxOp = Nothing
            dDistance = Nothing
            nNumeroPartie = Nothing
            nNumeroSommet = Nothing
            bCoteDroit = Nothing
            dTolerance = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de valider la distance entre les points d'adjacence.
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    ''' 
    ''' <remarks>
    ''' À la fin du traitement :
    ''' Les erreurs suivantes sont identifiées :
    '''    -Élément adjacent à déplacer
    ''' </remarks>
    ''' 
    Public Sub ValiderDistanceListePointAdjacence(ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pPointColl As IPointCollection = Nothing    'Interface ESRI utilisé pour accéder aux sommets de la limite de découpage avec points.
        Dim oErreur As Structure_Erreur = Nothing       'Structure contenant une erreur d'adjacence.
        Dim pProxOp As IProximityOperator = Nothing     'Interface ESRI utilisé pour calculer la distance.
        Dim qElementTraiterA As Structure_Element_Traiter = Nothing 'Contient un élément à traiter.
        Dim qElementTraiterB As Structure_Element_Traiter = Nothing 'Contient un élément à traiter.
        Dim qElementCollA As Collection = Nothing       'Objet qui contient la liste des éléments adjacents.
        Dim qElementCollB As Collection = Nothing       'Objet qui contient la liste des éléments adjacents.
        Dim pFeatureA As IFeature = Nothing             'Interface contenant un élément.
        Dim pFeatureB As IFeature = Nothing             'Interface contenant un élément.
        Dim pPointA As IPoint = Nothing                 'Interface contenant le point en erreur.
        Dim pPointB As IPoint = Nothing                 'Interface contenant le point en erreur.
        Dim sIdentifiantA As String = Nothing           'Contient la valeur de l'identifiant.
        Dim sIdentifiantB As String = Nothing           'Contient la valeur de l'identifiant.
        Dim dDistance As Double = Nothing               'Contient la distance d'adjacence.
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Interface utilisé pour accéder aux points d'adjacence sur la limite de découpage
            pPointColl = CType(gpLimiteDecoupageAvecPoint, IPointCollection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Valider la distance entre les points d'adjacence ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pPointColl.PointCount - 2, pTrackCancel)

            'Traiter tous les points de découpage
            For i = 0 To pPointColl.PointCount - 2
                'Définir le premier point
                pPointA = pPointColl.Point(i)
                'Définir le prochain point
                pPointB = pPointColl.Point(i + 1)

                'Vérifier si les sommets ne sont pas virtuel
                If pPointA.ID > 0 And pPointB.ID > 0 Then
                    'Projeter les points
                    pPointA.Project(gpSpatialReferenceProj)
                    pPointB.Project(gpSpatialReferenceProj)

                    'Interface pour calculer la distance d'adjacence
                    pProxOp = CType(pPointA, IProximityOperator)
                    'Calculer la distance d'adjacence
                    dDistance = pProxOp.ReturnDistance(pPointB)

                    'Sortir si la distance est plus grande que la tolérance d'adjacence
                    If dDistance <= gdTolAdjacenceOri Then
                        'Définir la liste des sommets d'éléments adjacents
                        qElementCollA = CType(gpListeElementPointAdjacent.Item(Math.Abs(pPointA.ID)), Collection)
                        'Définir l'élément
                        qElementTraiterA = CType(gpListeElementTraiter.Item(qElementCollA.Item(1)), Structure_Element_Traiter)
                        pFeatureA = qElementTraiterA.FeatureClass.GetFeature(qElementTraiterA.OID)
                        'Définir la valeur de l'identifiant
                        sIdentifiantA = fsIdentifiantDecoupage(pFeatureA, gsIdentifiantDecoupage)

                        'Définir la liste des sommets d'éléments adjacents
                        qElementCollB = CType(gpListeElementPointAdjacent.Item(Math.Abs(pPointB.ID)), Collection)
                        'Définir l'élément
                        qElementTraiterB = CType(gpListeElementTraiter.Item(qElementCollB.Item(1)), Structure_Element_Traiter)
                        pFeatureB = qElementTraiterB.FeatureClass.GetFeature(qElementTraiterB.OID)
                        'Définir la valeur de l'identifiant
                        sIdentifiantB = fsIdentifiantDecoupage(pFeatureB, gsIdentifiantDecoupage)

                        'Vérifier si une erreur de déplacement doit être identifiée
                        'Si l'adjacence est unique, on peut déplacer si les deux sommets ont seulement 1 élément
                        If gbAdjacenceUnique = False Or (gbAdjacenceUnique And qElementCollA.Count = 1 And qElementCollB.Count = 1) Then
                            'Conserver l'erreur dans une structure
                            oErreur = New Structure_Erreur
                            oErreur.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + gdTolAdjacenceOri.ToString("0.0##") + ")"
                            oErreur.IdentifiantA = sIdentifiantA
                            oErreur.PointA = pPointA
                            oErreur.FeatureA = pFeatureA
                            oErreur.IdentifiantB = sIdentifiantB
                            oErreur.PointB = pPointB
                            oErreur.FeatureB = pFeatureB
                            oErreur.Distance = dDistance

                            'Ajouter la structure d'erreur dans la collection
                            gpErreurFeatureAdjacence.Add(oErreur)
                            'Conserver l'dentifiant du point d'adjacence en erreur
                            If Not gpListeErreurPointAdjacent.Contains(pPointA.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPointA.ID, pPointA.ID.ToString)
                            'Conserver l'dentifiant du point d'adjacence en erreur
                            If Not gpListeErreurPointAdjacent.Contains(pPointB.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPointB.ID, pPointB.ID.ToString)
                        End If
                    End If
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pPointColl = Nothing
            oErreur = Nothing
            pProxOp = Nothing
            qElementTraiterA = Nothing
            qElementTraiterB = Nothing
            qElementCollA = Nothing
            qElementCollB = Nothing
            pPointA = Nothing
            pPointB = Nothing
            pFeatureA = Nothing
            pFeatureB = Nothing
            sIdentifiantA = Nothing
            sIdentifiantB = Nothing
            dDistance = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de valider l'adjacence des éléments de la liste des points d'adjacence.
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    ''' 
    ''' <remarks>
    ''' À la fin du traitement :
    '''  Les erreurs possibles sont:
    '''    -Plusieurs éléments adjacents
    '''    -Aucun élément adjacent
    '''    -Valeurs différentes d'attribut    
    ''' </remarks>
    ''' 
    Private Sub ValiderAdjacenceListePointAdjacence(ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim qElementTraiter As Structure_Element_Traiter = Nothing 'Contient un élément à traiter.
        Dim oErreur As Structure_Erreur = Nothing       'Structure contenant une erreur d'attribut
        Dim pFeatureA As IFeature = Nothing             'Interface contenant un élément.
        Dim pFeatureB As IFeature = Nothing             'Interface contenant un élément.
        Dim pFeatureAdj As IFeature = Nothing           'Interface contenant un élément ajacent.
        Dim pDatasetA As IDataset = Nothing             'Interface contenant le Path de la Geodatabase de l'élément.
        Dim pDatasetB As IDataset = Nothing             'Interface contenant le Path de la Geodatabase de l'élément.
        Dim qElementColl As Collection = Nothing        'Objet qui contient la liste des éléments adjacents.
        Dim pPoint As IPoint = Nothing                  'Interface contenant le sommet traité.
        Dim sIdentifiantA As String = Nothing           'Contient la valeur de l'identifiant.
        Dim sIdentifiantB As String = Nothing           'Contient la valeur de l'identifiant.
        Dim sIdentifiantAdj As String = Nothing         'Contient la valeur de l'identifiant de l'élément adjacent.
        Dim sClef As String = Nothing                   'Contient la valeur de la clef de recherche de l'erreur.
        Dim nNbAdjacent As Integer = Nothing            'Contient le nombre d'éléments adjacents de même classe.
        Dim pPolyline As IPolyline = Nothing            'Interface pour extraire les extrémités de ligne.
        Dim i As Integer = Nothing                      'Compteur
        Dim j As Integer = Nothing                      'Compteur
        Dim k As Integer = Nothing                      'Compteur

        Try
            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Valider la liste des points d'adjacence en cours ..."
            'Afficher la barre de progression
            InitBarreProgression(0, gpListeElementPointAdjacent.Count, pTrackCancel)

            'Traiter tous les points d'adjacence
            For i = 1 To gpListeElementPointAdjacent.Count
                'Définir le point d'adjacence
                pPoint = CType(gpListePointAdjacence.Geometry(i - 1), IPoint)

                'Si le point d'adjacence n'a pas déjà une erreur
                If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then
                    'Définir la liste des éléments adjacents
                    qElementColl = CType(gpListeElementPointAdjacent.Item(i), Collection)

                    'Traiter tous les éléments adjacents pour un point d'adjacence
                    For j = 1 To qElementColl.Count
                        'Définir l'élément
                        qElementTraiter = CType(gpListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)
                        pFeatureA = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                        pDatasetA = CType(pFeatureA.Class, IDataset)

                        'Définir la valeur de l'identifiant
                        sIdentifiantA = fsIdentifiantDecoupage(pFeatureA, gsIdentifiantDecoupage)

                        'Initialiser le nombre d'éléments adjacents
                        nNbAdjacent = 0
                        'Vérifier l'adjacence de tous les éléments
                        For k = 1 To qElementColl.Count
                            'Définir l'élément adjacent
                            qElementTraiter = CType(gpListeElementTraiter.Item(qElementColl.Item(k)), Structure_Element_Traiter)
                            pFeatureB = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                            pDatasetB = CType(pFeatureB.Class, IDataset)

                            'Définir la valeur de l'identifiant adjacent
                            sIdentifiantB = fsIdentifiantDecoupage(pFeatureB, gsIdentifiantDecoupage)

                            'Vérifier si ce n'est pas le même OBJECTID ou si ce n'est pas la même classe, ou si ce n'est pas la même Géodatabase
                            If pFeatureA.OID <> pFeatureB.OID Or pFeatureA.Class.AliasName <> pFeatureB.Class.AliasName Or pDatasetA.Workspace.PathName <> pDatasetB.Workspace.PathName Then
                                'Vérifier si c'est la même classe ou si si on permet une classe différente
                                If pFeatureA.Class.AliasName = pFeatureB.Class.AliasName Or gbClasseDifferente = True Then
                                    'Vérifier si l'identifiant est différent de celui de son adjacent
                                    If sIdentifiantA <> sIdentifiantB Then
                                        'Vérifier si les éléments sont adjacents au point d'adjacence
                                        If EstAdjacent(pPoint, pFeatureA, pFeatureB) Then
                                            'Compter le nombre d'éléments adjacents d'identifiant différent
                                            nNbAdjacent = nNbAdjacent + 1
                                            'Définir l'élément adjacent
                                            pFeatureAdj = pFeatureB
                                            'Définir l'identifiant adjacent
                                            sIdentifiantAdj = sIdentifiantB

                                            'Valider les valeurs des attributs d'adjacence
                                            Call ValiderAttributAdjacence(pPoint, pFeatureA, pFeatureB, sIdentifiantA, sIdentifiantB)
                                        End If
                                    End If
                                End If
                            End If
                        Next

                        'Vérifier si l'élément possède plusieurs éléments adjacents
                        If nNbAdjacent > 1 Then
                            'Vérifier si seulement l'adjacence unique est permise
                            If gbAdjacenceUnique Then
                                'Conserver l'erreur dans une structure
                                oErreur = New Structure_Erreur
                                oErreur.Description = "Plusieurs éléments adjacents (Nb éléments=" & qElementColl.Count.ToString & ", Nb adjacents=" & nNbAdjacent.ToString & ")"
                                oErreur.IdentifiantA = sIdentifiantA
                                oErreur.PointA = pPoint
                                oErreur.FeatureA = pFeatureA
                                oErreur.ValeurA = nNbAdjacent.ToString
                                oErreur.IdentifiantB = sIdentifiantAdj
                                oErreur.PointB = pPoint
                                oErreur.FeatureB = pFeatureAdj
                                oErreur.Distance = 0

                                'Vérifier si le OID de l'élément est plus petit
                                If pFeatureA.OID < pFeatureAdj.OID Then
                                    'Définir la clef de recherche de l'erreur
                                    sClef = pPoint.ID.ToString & ":" & pFeatureA.OID.ToString & "/" & pFeatureAdj.OID.ToString & ":" & pFeatureA.Class.AliasName
                                Else
                                    'Définir la clef de recherche de l'erreur
                                    sClef = pPoint.ID.ToString & ":" & pFeatureAdj.OID.ToString & "/" & pFeatureA.OID.ToString & ":" & pFeatureA.Class.AliasName
                                End If

                                'Ajouter la structure d'erreur dans la collection
                                If Not gpErreurFeatureAdjacence.Contains(sClef) Then gpErreurFeatureAdjacence.Add(oErreur, sClef)
                                'Conserver l'dentifiant du point d'adjacence en erreur
                                If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                            End If

                            'Vérifier si l'élément possède un seul élément adjacent
                        ElseIf nNbAdjacent = 1 Then
                            'Vérifier si plusieurs éléments adjacents et si seulement l'adjacence unique est permise
                            If qElementColl.Count > 2 And gbAdjacenceUnique Then
                                'Conserver l'erreur dans une structure
                                oErreur = New Structure_Erreur
                                oErreur.Description = "Un seul élément adjacent mais plusieurs éléments présents (Nb éléments=" & qElementColl.Count.ToString & ", Nb adjacents=" & nNbAdjacent.ToString & ")"
                                oErreur.IdentifiantA = sIdentifiantA
                                oErreur.PointA = pPoint
                                oErreur.FeatureA = pFeatureA
                                oErreur.ValeurA = nNbAdjacent.ToString
                                oErreur.IdentifiantB = sIdentifiantAdj
                                oErreur.PointB = pPoint
                                oErreur.FeatureB = pFeatureAdj
                                oErreur.Distance = 0

                                'Vérifier si le OID de l'élément est plus petit
                                If pFeatureA.OID < pFeatureAdj.OID Then
                                    'Définir la clef de recherche de l'erreur
                                    sClef = pPoint.ID.ToString & ":" & pFeatureA.OID.ToString & "/" & pFeatureAdj.OID.ToString & ":" & pFeatureA.Class.AliasName
                                Else
                                    'Définir la clef de recherche de l'erreur
                                    sClef = pPoint.ID.ToString & ":" & pFeatureAdj.OID.ToString & "/" & pFeatureA.OID.ToString & ":" & pFeatureA.Class.AliasName
                                End If

                                'Ajouter la structure d'erreur dans la collection
                                If Not gpErreurFeatureAdjacence.Contains(sClef) Then gpErreurFeatureAdjacence.Add(oErreur, sClef)
                                'Conserver l'dentifiant du point d'adjacence en erreur
                                If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                            End If

                            'Vérifier si aucun élément adjacent
                        ElseIf nNbAdjacent = 0 Then
                            'Vérifier si aucun élément adjacent
                            If qElementColl.Count = 1 Then
                                'Si la géométrie de l'élément est une ligne
                                If pFeatureA.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                                    'Interface pour extraire les extrémités de la ligne
                                    pPolyline = CType(pFeatureA.Shape, IPolyline)
                                    'Vérifier si le point d'adjacence correspond à une extrémité de la ligne
                                    If pPolyline.FromPoint.Compare(pPoint) = 0 Or pPolyline.ToPoint.Compare(pPoint) = 0 Then
                                        'Conserver l'erreur dans une structure
                                        oErreur = New Structure_Erreur
                                        oErreur.Description = "Aucun sommet d'élément adjacent"
                                        oErreur.IdentifiantA = sIdentifiantA
                                        oErreur.PointA = pPoint
                                        oErreur.FeatureA = pFeatureA
                                        oErreur.ValeurA = nNbAdjacent.ToString
                                        oErreur.Distance = 0

                                        'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                        If Not gpErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then gpErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                        'Conserver l'dentifiant du point d'adjacence en erreur
                                        If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)

                                        'Si le point d'adjacence ne correspond pas une extrémité de la ligne et si doit tenir compte des identifiants
                                    ElseIf gbSansIdentifiant = False Then
                                        'Conserver l'erreur dans une structure
                                        oErreur = New Structure_Erreur
                                        oErreur.Description = "Aucun élément adjacent mais n'est pas une extrémité de ligne"
                                        oErreur.IdentifiantA = sIdentifiantA
                                        oErreur.PointA = pPoint
                                        oErreur.FeatureA = pFeatureA
                                        oErreur.ValeurA = nNbAdjacent.ToString
                                        oErreur.Distance = 0

                                        'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                        If Not gpErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then gpErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                        'Conserver l'dentifiant du point d'adjacence en erreur
                                        If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                                    End If

                                    'Si la géométrie de l'élément n'est pas une ligne
                                Else
                                    'Conserver l'erreur dans une structure
                                    oErreur = New Structure_Erreur
                                    oErreur.Description = "Aucun sommet d'élément adjacent"
                                    oErreur.IdentifiantA = sIdentifiantA
                                    oErreur.PointA = pPoint
                                    oErreur.FeatureA = pFeatureA
                                    oErreur.ValeurA = nNbAdjacent.ToString
                                    oErreur.Distance = 0

                                    'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                    If Not gpErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then gpErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                    'Conserver l'dentifiant du point d'adjacence en erreur
                                    If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                                End If

                                'Si deux éléments
                            ElseIf qElementColl.Count = 2 Then
                                'Si on doit tenir compte des identifiants
                                If gbSansIdentifiant = False Then
                                    'Conserver l'erreur dans une structure
                                    oErreur = New Structure_Erreur
                                    oErreur.Description = "Aucun élément adjacent mais un autre élément est présent (Nb éléments=" & qElementColl.Count.ToString & ")"
                                    oErreur.IdentifiantA = sIdentifiantA
                                    oErreur.PointA = pPoint
                                    oErreur.FeatureA = pFeatureA
                                    oErreur.ValeurA = nNbAdjacent.ToString
                                    oErreur.Distance = 0

                                    'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                    If Not gpErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then gpErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                    'Conserver l'dentifiant du point d'adjacence en erreur
                                    If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                                End If

                                'Si plusieurs éléments
                            ElseIf qElementColl.Count > 2 Then
                                'Si on doit tenir compte des identifiants ou si seulement l'adjacence unique est permise 
                                If gbSansIdentifiant = False Or gbAdjacenceUnique Then
                                    'Conserver l'erreur dans une structure
                                    oErreur = New Structure_Erreur
                                    oErreur.Description = "Aucun élément adjacent mais plusieurs éléments présents (Nb éléments=" & qElementColl.Count.ToString & ")"
                                    oErreur.IdentifiantA = sIdentifiantA
                                    oErreur.PointA = pPoint
                                    oErreur.FeatureA = pFeatureA
                                    oErreur.ValeurA = nNbAdjacent.ToString
                                    oErreur.Distance = 0

                                    'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                    If Not gpErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then gpErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                    'Conserver l'dentifiant du point d'adjacence en erreur
                                    If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                                End If
                            End If
                        End If
                    Next
                End If

                'Récupération de la mémoire disponible
                GC.Collect()

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            qElementTraiter = Nothing
            oErreur = Nothing
            pFeatureA = Nothing
            pFeatureB = Nothing
            pFeatureAdj = Nothing
            pDatasetA = Nothing
            pDatasetB = Nothing
            qElementColl = Nothing
            pPoint = Nothing
            pPolyline = Nothing
            sIdentifiantA = Nothing
            sIdentifiantB = Nothing
            sIdentifiantAdj = Nothing
            sClef = Nothing
            nNbAdjacent = Nothing
            i = Nothing
            j = Nothing
            k = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de valider tous les attributs d'adjacence pour un point d'adjacence. 
    ''' Le ListBox des erreurs d'attribut est rempli.
    '''</summary>
    '''
    '''<param name="pPoint">Interface contenant la géométrie du point d'adjacence.</param>
    '''<param name="pFeatureA">Interface ESRI contenant l'élément de base.</param>
    '''<param name="pFeatureB">Interface ESRI contenant l'élément à comparer.</param>
    ''' 
    '''<returns>True si les éléments sont adjacents, False sinon.</returns>
    ''' 
    Private Function EstAdjacent(pPoint As IPoint, pFeatureA As IFeature, pFeatureB As IFeature) As Boolean
        'Déclarer les variables de travail
        Dim pHitTest As IHitTest = Nothing                  'Interface ESRI pour tester la présence du sommet recherché
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire un sommet d'une géométrie.
        Dim pGeometryColl As IGeometryCollection = Nothing  'Interface pour extraire une composante d'une géométrie.
        Dim pProxOp As IProximityOperator = Nothing         'Interface pour vérifier si le sommet précédent ou suivant touche l'élément adjacent.
        Dim pNewPoint As IPoint = Nothing           'Interface contenant le sommet trouvé.
        Dim dDistance As Double = Nothing           'Interface contenant la distance calculée entre le point de recherche et le sommet trouvé.
        Dim nNumeroPartie As Integer = Nothing      'Numéro de partie trouvée.
        Dim nNumeroSommet As Integer = Nothing      'Numéro de sommet de la partie trouvée.
        Dim bCoteDroit As Boolean = Nothing         'Indiquer si le point trouvé est du côté droit de la géométrie.

        'Par défaut, les éléments ne sont pas adjacents
        EstAdjacent = False

        Try
            'Vérifier si la géométrie est une surface
            If pFeatureA.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour extraire le numéro de sommet
                pHitTest = CType(pFeatureA.Shape, IHitTest)

                'Vérifier la présence d'un sommet existant selon une tolérance
                If pHitTest.HitTest(pPoint, gdTolRecherche, esriGeometryHitPartType.esriGeometryPartVertex, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Interface pour extraire une composante de la géométrie
                    pGeometryColl = CType(pFeatureA.Shape, IGeometryCollection)

                    'Interface pour extraire un sommet d'une géométrie
                    pPointColl = CType(pGeometryColl.Geometry(nNumeroPartie), IPointCollection)

                    'Extraire le point précédant
                    If nNumeroSommet = 0 Then
                        'Extraire le point précédant
                        pNewPoint = pPointColl.Point(pPointColl.PointCount - 2)
                    Else
                        'Extraire le point précédant
                        pNewPoint = pPointColl.Point(nNumeroSommet - 1)
                    End If
                    'Définir le centre de la ligne
                    pNewPoint.X = (pNewPoint.X + pPoint.X) / 2
                    pNewPoint.Y = (pNewPoint.Y + pPoint.Y) / 2

                    'Extraire le point précédant
                    pProxOp = CType(pNewPoint, IProximityOperator)

                    'Vérifier si le point touche la surface
                    If pProxOp.ReturnDistance(pFeatureB.Shape) <= gdTolRecherche Then
                        'Les éléments sont adjacents
                        Return True
                    End If

                    'Extraire le point suivant
                    If nNumeroSommet = pPointColl.PointCount Then
                        'Extraire le point suivant
                        pNewPoint = pPointColl.Point(1)
                    Else
                        'Extraire le point suivant
                        pNewPoint = pPointColl.Point(nNumeroSommet + 1)
                    End If
                    'Définir le centre de la ligne
                    pNewPoint.X = (pNewPoint.X + pPoint.X) / 2
                    pNewPoint.Y = (pNewPoint.Y + pPoint.Y) / 2

                    'Extraire le point précédant
                    pProxOp = CType(pNewPoint, IProximityOperator)

                    'Vérifier si le point touche la surface
                    If pProxOp.ReturnDistance(pFeatureB.Shape) <= gdTolRecherche Then
                        'Les éléments sont adjacents
                        Return True
                    End If
                End If

                'Si la géométrie n'est pas une surface
            Else
                'Les éléments sont adjacents
                EstAdjacent = True
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pHitTest = Nothing
            pPointColl = Nothing
            pGeometryColl = Nothing
            pProxOp = Nothing
            pNewPoint = Nothing
            dDistance = Nothing
            nNumeroPartie = Nothing
            nNumeroSommet = Nothing
            bCoteDroit = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de valider tous les attributs d'adjacence pour un point d'adjacence. 
    ''' Le ListBox des erreurs d'attribut est rempli.
    '''</summary>
    '''
    '''<param name="pPoint">Interface contenant la géométrie du point d'adjacence.</param>
    '''<param name="pFeature">Interface ESRI contenant l'élément de base.</param>
    '''<param name="pFeatureAdj">Interface ESRI contenant l'élément à comparer.</param>
    '''<param name="sIdentifiant">Identifiant de l'élément de base.</param>
    '''<param name="sIdentifiantAdj">Identifiant de l'élément à comparer.</param>
    ''' 
    Private Sub ValiderAttributAdjacence(ByVal pPoint As IPoint, ByVal pFeature As IFeature, ByVal pFeatureAdj As IFeature, _
                                         ByVal sIdentifiant As String, ByVal sIdentifiantAdj As String)
        'Déclarer les variables de travail
        Dim oErreur As Structure_Erreur = Nothing   'Structure contenant une erreur d'attribut
        Dim sNomAttribut As String = Nothing        'Contient le nom de l'attribut à traiter
        Dim sClef As String = Nothing               'Contient la valeur de la clef de recherche de l'erreur
        Dim sValeurA As String = Nothing            'Contient la valeur de l'attribut pour l'élément de base
        Dim sValeurB As String = Nothing            'Contient la valeur de l'attribut pour l'élément à comparer
        Dim nPosAttA As Integer = Nothing           'Contient la position de l'attribut pour l'élément de base
        Dim nPosAttB As Integer = Nothing           'Contient la position de l'attribut pour l'élément à comparer
        Dim i As Integer = Nothing                  'Compteur

        Try
            'Traiter tous les attributs d'adjacence
            For i = 1 To gpAttributAdjacence.Count
                'Définir le nom de l'attribt à traiter
                sNomAttribut = CType(gpAttributAdjacence.Item(i), String)

                'Définir la valeur de l'attribut de base
                sValeurA = "."
                'Définir la position de l'attribut de base
                nPosAttA = pFeature.Fields.FindField(sNomAttribut)
                'Vérifier si l'attribut de l'élément de base est présent
                If nPosAttA >= 0 Then
                    'Définir la valeur de l'attribut de base
                    sValeurA = pFeature.Value(nPosAttA).ToString
                End If

                'Définir la valeur de l'attribut à comparer
                sValeurB = "."
                'Définir la position de l'attribut de base
                nPosAttB = pFeatureAdj.Fields.FindField(sNomAttribut)
                'Vérifier si l'attribut de l'élément à comparer est présent
                If nPosAttB >= 0 Then
                    'Définir la valeur de l'attribut à comparer
                    sValeurB = pFeatureAdj.Value(nPosAttB).ToString
                End If

                'Vérifier si l'attribut de l'élément de base est différent de celui à comparer
                If sValeurA <> sValeurB Then
                    'Vérifier si le OID de l'élément est plus petit
                    If pFeature.OID < pFeatureAdj.OID Then
                        'Définir la clef de recherche de l'erreur
                        sClef = pPoint.ID.ToString & ":" & pFeature.OID.ToString & "/" & pFeatureAdj.OID.ToString & ":" & pFeature.Class.AliasName
                    Else
                        'Définir la clef de recherche de l'erreur
                        sClef = pPoint.ID.ToString & ":" & pFeatureAdj.OID.ToString & "/" & pFeature.OID.ToString & ":" & pFeature.Class.AliasName
                    End If

                    'Vérifier si l'erreur est déja présente
                    If Not gpErreurFeatureAttribut.Contains(sClef) Then
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Valeurs différentes d'attribut (" & sNomAttribut & "=" & sValeurA & "<>" & sValeurB & ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pPoint
                        oErreur.FeatureA = pFeature
                        oErreur.ValeurA = sValeurA
                        oErreur.PosAttA = nPosAttA
                        oErreur.PointB = pPoint
                        oErreur.IdentifiantB = sIdentifiantAdj
                        oErreur.FeatureB = pFeatureAdj
                        oErreur.ValeurB = sValeurB
                        oErreur.PosAttB = nPosAttB
                        oErreur.Distance = 0

                        'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                        gpErreurFeatureAttribut.Add(oErreur, sClef)
                        'Conserver l'dentifiant du point d'adjacence en erreur
                        If Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then gpListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                    End If
                End If
            Next

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            oErreur = Nothing
            sNomAttribut = Nothing
            sClef = Nothing
            sValeurA = Nothing
            sValeurB = Nothing
            nPosAttA = Nothing
            nPosAttB = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de trouver et identifier l'élément adjacent pour tous les éléments isolés.
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    ''' 
    ''' <remarks>
    ''' À la fin du traitement :
    ''' Les erreurs "Aucun élément adjacent" identifiées préalablement sont éliminées si une corresponde est trouvée.
    ''' Les erreurs suivantes sont identifiées :
    '''    -Élément adjacent à déplacer
    ''' </remarks>
    ''' 
    Public Sub TrouverElementAdjacentIsole(ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim oErreurA As Structure_Erreur = Nothing  'Structure contenant une erreur d'adjacence
        Dim oErreurB As Structure_Erreur = Nothing  'Structure contenant une erreur d'adjacence
        Dim pPointColl As IPointCollection = Nothing 'Interface ESRI utilisé pour accéder aux sommets de la limite de découpage avec points
        Dim pProxOp As IProximityOperator = Nothing 'Interface ESRI utilisé pour calculer la distance
        Dim pPointA As IPoint = Nothing             'Interface contenant le point en erreur
        Dim pPointB As IPoint = Nothing             'Interface contenant le point en erreur
        Dim dDistance As Double = Nothing           'Contient la distance d'adjacence
        Dim i As Integer = Nothing                  'Compteur
        Dim j As Integer = Nothing                  'Compteur
        Dim k As Integer = Nothing                  'Compteur

        Try
            'Interface utilisé pour accéder aux points d'adjacence sur la limite de découpage
            pPointColl = CType(gpLimiteDecoupageAvecPoint, IPointCollection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Trouver les éléments adjacents aux points d'adjacence isolés en cours ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pPointColl.PointCount - 2, pTrackCancel)

            'Traiter tous les points de découpage
            For i = 0 To pPointColl.PointCount - 2
                'Définir le point en erreur
                pPointA = pPointColl.Point(i)

                'Vérifier si le point est en erreur
                If gpErreurFeatureAdjacence.Contains(pPointA.ID.ToString) Then
                    'Définir l'erreur A
                    oErreurA = CType(gpErreurFeatureAdjacence.Item(pPointA.ID.ToString), Structure_Erreur)

                    'Trouver l'élément adjacent
                    For j = i + 1 To pPointColl.PointCount - 1
                        'Définir le point d'adjacence correspondant
                        pPointB = pPointColl.Point(j)

                        'Projeter les points
                        pPointA.Project(gpSpatialReferenceProj)
                        pPointB.Project(gpSpatialReferenceProj)
                        'Interface pour calculer la distance d'adjacence
                        pProxOp = CType(pPointA, IProximityOperator)
                        'Calculer la distance d'adjacence
                        dDistance = pProxOp.ReturnDistance(pPointB)

                        'Sortir si la distance est plus grande que la tolérance d'adjacence
                        'Debug.Print(dDistance.ToString("0.0#######"))
                        If dDistance > gdTolAdjacenceOri Then Exit For

                        'Vérifier si le point correspondant est aussi en erreur
                        If gpErreurFeatureAdjacence.Contains(pPointB.ID.ToString) Then
                            'Définir l'erreur B
                            oErreurB = CType(gpErreurFeatureAdjacence.Item(pPointB.ID.ToString), Structure_Erreur)

                            'Vérifier si c'est la même classe
                            If oErreurA.FeatureA.Class.AliasName = oErreurB.FeatureA.Class.AliasName Or gbClasseDifferente Then
                                'Ajouter le lien d'adjacence dans l'erreur
                                oErreurA.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + gdTolAdjacenceOri.ToString("0.0##") + ")"
                                oErreurA.IdentifiantB = oErreurB.IdentifiantA
                                oErreurA.FeatureB = oErreurB.FeatureA
                                oErreurA.PointB = oErreurB.PointA
                                oErreurA.Distance = dDistance

                                'Retirer l'élément A d'adjacence isolé en erreur
                                gpErreurFeatureAdjacence.Remove(pPointA.ID.ToString)

                                'Retirer l'élément B d'adjacence isolé en erreur
                                gpErreurFeatureAdjacence.Remove(pPointB.ID.ToString)

                                'Ajouter la structure d'erreur dans la collection
                                gpErreurFeatureAdjacence.Add(oErreurA)

                                'Sortir de la boucle lorsque l'élément est trouvé
                                Exit For
                            End If
                        End If
                    Next
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            oErreurA = Nothing
            oErreurB = Nothing
            pPointColl = Nothing
            pProxOp = Nothing
            pPointA = Nothing
            pPointB = Nothing
            dDistance = Nothing
            i = Nothing
            j = Nothing
            k = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'écrire les erreurs de précision. 
    '''</summary>
    '''
    '''<param name="pGeomResColl">Interface contenant les géométries d'erreurs.</param>
    ''' 
    Public Sub EcrireErreurPrecision(ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryErr As IMultipoint = Nothing   'Interface contenant la géométrie d'erreur.
        Dim pPointColl As IPointCollection = Nothing 'Interface utilisé pour ajouter un point.
        Dim oErreur As Structure_Erreur             'Structure contenant une erreur de précision.
        Dim sDesc As String = ""                    'Description de l'erreur.
        Dim i As Integer = 0                        'Compteur

        Try
            'Traiter toutes les erreurs de précision
            For i = 1 To gpErreurFeaturePrecision.Count
                'Définir l'erreur de précision
                oErreur = CType(gpErreurFeaturePrecision.Item(i), Structure_Erreur)

                'Définir la description
                sDesc = "Erreur de précision : " & oErreur.Description & " /PointID=" & oErreur.PointA.ID.ToString _
                        & ", Identifiant=" & oErreur.IdentifiantA & ", OID=" & oErreur.FeatureA.OID.ToString _
                        & ", Classe=" & oErreur.FeatureA.Class.AliasName

                'Définir la géométrie d'erreur
                pGeometryErr = GeometrieToMultiPoint(oErreur.PointA)
                'Vérifier si un point correspondant est trouvé
                If oErreur.PointB IsNot Nothing Then
                    'Ajouter le point correspondant
                    pPointColl = CType(pGeometryErr, IPointCollection)
                    pPointColl.AddPoint(oErreur.PointB)
                End If
                'Ajouter la géométrie d'erreur
                pGeomResColl.AddGeometry(pGeometryErr)

                'Écrire une erreur
                EcrireFeatureErreur(sDesc, pGeometryErr, CSng(oErreur.Distance))
            Next

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryErr = Nothing
            oErreur = Nothing
            sDesc = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'écrire les erreurs d'adjacence. 
    '''</summary>
    '''
    '''<param name="pGeomResColl">Interface contenant les géométries d'erreurs.</param>
    ''' 
    Public Sub EcrireErreurAdjacence(ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryErr As IMultipoint = Nothing   'Interface contenant la géométrie d'erreur.
        Dim pPointColl As IPointCollection = Nothing 'Interface utilisé pour ajouter un point.
        Dim oErreur As Structure_Erreur             'Structure contenant une erreur d'adjacence.
        Dim sDesc As String = ""                    'Description de l'erreur.
        Dim i As Integer = 0                        'Compteur

        Try
            'Traiter toutes les erreurs d'adjacence
            For i = 1 To gpErreurFeatureAdjacence.Count
                'Définir l'erreur d'adjacence
                oErreur = CType(gpErreurFeatureAdjacence.Item(i), Structure_Erreur)

                'Définir la description
                sDesc = "Erreur d'adjacence : " & oErreur.Description & " /PointID=" & oErreur.PointA.ID.ToString _
                        & ", Identifiant=" & oErreur.IdentifiantA & ", OID=" & oErreur.FeatureA.OID.ToString _
                        & ", Classe=" & oErreur.FeatureA.Class.AliasName

                'Définir la géométrie d'erreur
                pGeometryErr = GeometrieToMultiPoint(oErreur.PointA)
                'Vérifier si un point correspondant est trouvé
                If oErreur.PointB IsNot Nothing Then
                    'Ajouter le point correspondant
                    pPointColl = CType(pGeometryErr, IPointCollection)
                    pPointColl.AddPoint(oErreur.PointB)
                    'Définir la description
                    sDesc = sDesc & " /PointID=" & oErreur.PointB.ID.ToString
                    'Vérifier si l'élément est spécifié
                    If oErreur.FeatureB IsNot Nothing Then
                        'Compléter la description
                        sDesc = sDesc & ", Identifiant=" & oErreur.IdentifiantB & ", OID=" & oErreur.FeatureB.OID.ToString _
                                & ", Classe=" & oErreur.FeatureB.Class.AliasName
                    End If
                End If
                'Ajouter la géométrie d'erreur
                pGeomResColl.AddGeometry(pGeometryErr)

                'Écrire une erreur
                EcrireFeatureErreur(sDesc, pGeometryErr, CSng(oErreur.Distance))
            Next

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryErr = Nothing
            oErreur = Nothing
            sDesc = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'écrire les erreurs d'attribut. 
    '''</summary>
    '''
    '''<param name="pGeomResColl">Interface contenant les géométries d'erreurs.</param>
    ''' 
    Public Sub EcrireErreurAttribut(ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryErr As IMultipoint = Nothing   'Interface contenant la géométrie d'erreur.
        Dim oErreur As Structure_Erreur             'Structure contenant une erreur d'attribut
        Dim sDesc As String = ""                    'Description de l'erreur.
        Dim i As Integer = 0                        'Compteur

        Try
            'Traiter toutes les erreurs d'attribut
            For i = 1 To gpErreurFeatureAttribut.Count
                'Définir l'erreur d'attribut
                oErreur = CType(gpErreurFeatureAttribut.Item(i), Structure_Erreur)

                'Définir la description
                sDesc = "Erreur d'attribut : " & oErreur.Description & " /PointID=" & oErreur.PointA.ID.ToString _
                        & ", Identifiant=" & oErreur.IdentifiantA & ", OID=" & oErreur.FeatureA.OID.ToString _
                        & ", Classe=" & oErreur.FeatureA.Class.AliasName & ", Valeur=" & oErreur.ValeurA

                'Vérifier si l'élément correspondant est trouvé
                If Not oErreur.PointB Is Nothing Then
                    'Définir la description
                    sDesc = sDesc & " /PointID=" & oErreur.PointB.ID.ToString _
                            & ", Identifiant=" & oErreur.IdentifiantB & ", OID=" & oErreur.FeatureB.OID.ToString _
                            & ", Classe=" & oErreur.FeatureB.Class.AliasName & ", Valeur=" & oErreur.ValeurB
                End If

                'Définir la géométrie d'erreur
                pGeometryErr = GeometrieToMultiPoint(oErreur.PointA)

                'Ajouter la géométrie d'erreur
                pGeomResColl.AddGeometry(pGeometryErr)

                'Écrire une erreur
                EcrireFeatureErreur(sDesc, pGeometryErr, CSng(oErreur.Distance))
            Next

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryErr = Nothing
            oErreur = Nothing
            sDesc = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments selon le type de sélection Conserver ou Enlever. 
    '''</summary>
    '''
    '''<param name="pTrackCancel">Interface qui permet d'annuler la sélection avec la touche ESC.</param>
    '''<param name="pGeomResColl">Interface ESRI contenant les géométries résultantes trouvées..</param>
    ''' 
    Public Sub SelectionnerElement(ByRef pTrackCancel As ITrackCancel, ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour vérifier la présence des éléments sélectionnés.
        Dim qElementTraiter As Structure_Element_Traiter = Nothing 'Contient un élément à traiter.
        Dim pFeature As IFeature = Nothing              'Interface contenant un élément.
        Dim qElementColl As Collection = Nothing        'Objet qui contient la liste des éléments adjacents.
        Dim pPoint As IPoint = Nothing                  'Interface contenant le sommet traité.
        Dim i As Integer = Nothing                      'Compteur
        Dim j As Integer = Nothing                      'Compteur

        Try
            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments en cours ..."
            'Afficher la barre de progression
            InitBarreProgression(0, gpListeElementPointAdjacent.Count, pTrackCancel)

            'Interface pour sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            'Enlever la sélection
            pFeatureSel.Clear()

            'Traiter tous les points d'adjacence
            For i = 1 To gpListeElementPointAdjacent.Count
                'Définir le point d'adjacence
                pPoint = CType(gpListePointAdjacence.Geometry(i - 1), IPoint)

                'Vérifier si une erreur est présente et qu'on sélectionne les éléments en erreur
                If gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) And gbEnleverSelection Then
                    'Définir la liste des éléments adjacents
                    qElementColl = CType(gpListeElementPointAdjacent.Item(i), Collection)

                    'Traiter tous les éléments adjacents pour un point d'adjacence
                    For j = 1 To qElementColl.Count
                        'Définir l'élément
                        qElementTraiter = CType(gpListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)

                        'Ajouter l'élément dans la sélection
                        pFeatureSel.SelectionSet.Add(qElementTraiter.OID)
                    Next

                    'Si on sélectionne les éléments sans erreur
                ElseIf Not gpListeErreurPointAdjacent.Contains(pPoint.ID.ToString) And gbEnleverSelection = False Then
                    'Définir la liste des éléments adjacents
                    qElementColl = CType(gpListeElementPointAdjacent.Item(i), Collection)

                    'Traiter tous les éléments adjacents pour un point d'adjacence
                    For j = 1 To qElementColl.Count
                        'Définir l'élément
                        qElementTraiter = CType(gpListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)

                        'Ajouter l'élément dans la sélection
                        pFeatureSel.SelectionSet.Add(qElementTraiter.OID)
                    Next

                    'Ajouter la géométrie du point d'adjacence dans les points trouvées
                    pGeomResColl.AddGeometry(pPoint)

                    'Écrire une erreur dont le point d'adjacence est sans ereur
                    Call EcrireErreurValide(pPoint, qElementTraiter.OID, qElementColl.Count)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            qElementTraiter = Nothing
            qElementColl = Nothing
            pPoint = Nothing
            i = Nothing
            j = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'écrire les erreurs pour un point d'adjacence sans erreur. 
    '''</summary>
    '''
    '''<param name="pPoint">Interface contenant la géométrie du point d'adjacence valide.</param>
    '''<param name="iOid">OBJECTID du premier élément valide du point d'adjacence.</param>
    '''<param name="iNbElem">Nombre élément du point d'adjacence.</param>
    ''' 
    Public Sub EcrireErreurValide(ByVal pPoint As IPoint, ByVal iOid As Integer, ByVal iNbElem As Integer)
        'Déclarer les variables de travail
        Dim sDesc As String = ""                    'Description de l'erreur.

        Try
            'Définir la description
            sDesc = "Point d'adjacence valide /PointID=" & pPoint.ID.ToString _
                    & ", OID=" & iOid.ToString & ", Nb élément=" & iNbElem.ToString

            'Écrire une erreur de point d'adjacence valide
            EcrireFeatureErreur(sDesc, GeometrieToMultiPoint(pPoint), iNbElem)

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            sDesc = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de retourner le nom de l'identifiant de découpage.
    ''' L'identifiant de découpage extrait de deux façons.:
    ''' -La première est contenue dans la valeur de l'attribut d'identifiant de découpage s'il est présent.
    ''' -La deuxième est contenue dans le nom du fichier sans extension dans laquelle la classe est contenue.
    ''' La première façon a préséance sur la deuxième.
    '''</summary>
    '''
    '''<param name="pFeature">Interface ESRI contenant l'élément traité.</param>
    '''<param name="sIdentifiantDecoupage">Contient le nom de l'attribut d'identifiant de découpage.</param>
    ''' 
    '''<returns>La fonction va retourner un "String" contenant l'identifiant de découpage. Sinon "Nothing".</returns>
    '''
    Public Function fsIdentifiantDecoupage(ByRef pFeature As IFeature, ByVal sIdentifiantDecoupage As String) As String
        'Déclarer les variables de travail
        Dim pDataset As IDataset = Nothing  'Interface contenant le nom du Dataset
        Dim nPosAtt As Integer = Nothing    'Position de l'attribut d'identifiant

        'Définir la valeur par défaut
        fsIdentifiantDecoupage = ""

        Try
            'Définir la position de l'attribut d'identitifant
            nPosAtt = pFeature.Fields.FindField(sIdentifiantDecoupage)
            'Vérifier l'attribut est présent
            If nPosAtt >= 0 Then
                'Définir la valeur de l'identifiant à partir de la valeur de l'attribut
                fsIdentifiantDecoupage = pFeature.Value(nPosAtt).ToString
            Else
                'Définir la valeur de l'identifiant à partir du nom sans extension du fichier dans laquelle se trouve la classe
                pDataset = CType(pFeature.Class, IDataset)
                fsIdentifiantDecoupage = System.IO.Path.GetFileNameWithoutExtension(pDataset.Workspace.PathName)
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pDataset = Nothing
            nPosAtt = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer qui intersecte le buffer de recherche de la limite du découpage.
    '''</summary>
    '''
    '''<param name="pPolygon"> Géométrie du polygone de découpage.</param>
    ''' 
    Private Sub SelectionnerElementLimiteDecoupage(ByVal pPolygon As IPolygon)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface utiliser pour extraire les éléments à la limites du découpage.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface ESRI utilisé pour pour calculer la topologie.
        Dim pBufferLimite As IGeometry = Nothing            'Interface contenant le buffer de la limite de découpage.

        Try
            'Projeter le Polygone du découpage
            pPolygon.Project(gpSpatialReference)

            'Transformer les tolérances en géographique au besoin
            Call TransformerTolerances(pPolygon.Envelope)

            'Interface pour créer la limite du découpage
            pTopoOp = CType(pPolygon, ITopologicalOperator2)
            'Définir la limite du découpage
            gpLimiteDecoupage = CType(pTopoOp.Boundary, IPolyline)

            'Interface pour créer le buffer de la limite de découpage
            pTopoOp = CType(gpLimiteDecoupage, ITopologicalOperator2)
            'Créer le buffer de la limite
            pBufferLimite = pTopoOp.Buffer(gdTolRecherche)
            'Simplifier le buffer de la limite
            pTopoOp = CType(pBufferLimite, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Définir la requête spatiale
            pSpatialFilter = New SpatialFilter
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            pSpatialFilter.OutputSpatialReference(gpFeatureLayerSelection.FeatureClass.ShapeFieldName) = pBufferLimite.SpatialReference
            pSpatialFilter.Geometry = pBufferLimite

            'Interface pour sélectionner les éléments à la limite du découpage
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Sélectionner tous les éléments à la limite du découpage
            pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, False)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pSpatialFilter = Nothing
            pTopoOp = Nothing
            pBufferLimite = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de transformer les tolrérances de recherche et d'adjacence en géographique à partir d'un enveloppe d'une de zone travail.
    '''</summary>
    '''
    '''<param name="pEnvelope"> Envelope correspondant à la zone de travail.</param>
    ''' 
    Private Sub TransformerTolerances(ByVal pEnvelope As IEnvelope)
        'Déclarer les variables de travail
        Dim pClone As IClone = Nothing                              'Interface utilisé pour pour cloner une géométrie.
        Dim pEnvelopeProj As IEnvelope = Nothing                    'Interface contenant l'enveloppe projeté.
        Dim pSpatialRefFact As ISpatialReferenceFactory2 = Nothing  'Interface pour extraire une référence spatiale existante.
        Dim pProximityOp As IProximityOperator = Nothing    'Interface utilisé pour calculer la distance.
        Dim dDistanceGeog As Double = Nothing               'Distance géographique.
        Dim dDistanceProj As Double = Nothing               'Distance projetée.

        Try
            'Vérifier si la référence spatiale est géographique
            If TypeOf (pEnvelope.SpatialReference) Is GeographicCoordinateSystem Then
                'Cloner l'enveloppe
                pClone = CType(pEnvelope, IClone)
                pEnvelopeProj = CType(pClone.Clone, IEnvelope)

                'Interface pour extraire la référence spatiale
                pSpatialRefFact = New SpatialReferenceEnvironment

                'Définir la référence spatiale projeter LCC NAD83 CSRS:3979
                gpSpatialReferenceProj = pSpatialRefFact.CreateSpatialReference(3979)

                'Projeter la limite dans la référence spatiale LCC NAD83 CSRS:3979
                pEnvelopeProj.Project(gpSpatialReferenceProj)

                'Définir la distance projetée
                pProximityOp = CType(pEnvelopeProj.LowerLeft, IProximityOperator)
                dDistanceProj = pProximityOp.ReturnDistance(pEnvelopeProj.UpperRight)

                'Définir la distance géographique
                pProximityOp = CType(pEnvelope.LowerLeft, IProximityOperator)
                dDistanceGeog = pProximityOp.ReturnDistance(pEnvelope.UpperRight)

                'Redéfinir les tolérances
                gdTolRecherche = (gdTolRechercheOri * dDistanceGeog) / dDistanceProj
                gdTolAdjacence = (gdTolAdjacenceOri * dDistanceGeog) / dDistanceProj
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pClone = Nothing
            pEnvelopeProj = Nothing
            pSpatialRefFact = Nothing
            pProximityOp = Nothing
            dDistanceGeog = Nothing
            dDistanceProj = Nothing
        End Try
    End Sub
#End Region
End Class
