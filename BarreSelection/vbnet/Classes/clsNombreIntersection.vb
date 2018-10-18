Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt

'**
'Nom de la composante : clsNombreIntersection.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont le nombre d’intersection de géométrie de type point, ligne ou surface 
''' avec des éléments en relation respecte ou non l’expression régulière spécifiée.
''' 
''' Chaque géométrie des éléments traités est comparée avec ses éléments en relations.
''' 
''' La classe permet de traiter les quatre attributs de nombre d'intersection POINT, LIGNE, SURFACE et ELEMENT.
''' 
''' POINT : Le nombre d'intersection de type point (dimension=0) entre deux géométries.
''' LIGNE : Le nombre d'intersection de type ligne (dimension=1) entre deux géométries.
''' SURFACE : Le nombre d'intersection de type surface (dimension=2) entre deux géométries.
''' INTERIEURE : Le nombre d'intersection intérieures entre deux géométries.
''' EDGE : Le nombre d'intersection contenus dans un EDGE de topologie pour les éléments traités.
''' ELEMENT : Le nombre d'éléments qui intersecte la géométrie des éléments traités.
''' 
''' Note : Le nombre d'intersection correspond au nombre de composantes contenues dans la géométrie résultante de l'intersection.
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 23 avril 2015
'''</remarks>
''' 
Public Class clsNombreIntersection
    Inherits clsValeurAttribut

    '''<summary>Indique que les limites des géométries seront à traiter.</summary>
    Protected gbLimite As Boolean = False
    '''<summary>Indique que les limites des géométries en relation seront à traiter.</summary>
    Protected gbLimiteRel As Boolean = False
    '''<summary>Indique que chaque élément en relation seront à traiter de façon unique.</summary>
    Protected gbUnique As Boolean = False
    '''<summary>Indique si on veut simplifier les géométries d'intersection.</summary>
    Protected gbSimplifier As Boolean = False
    '''<summary>Contient les noms d'attributs dont les valeurs doivent être identiques lors d'une intersection.</summary>
    Protected gqValeurAttribut As Collection = Nothing

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "POINT"
            Expression = "^[1-2]$"
            gbLimite = False
            gbLimiteRel = False
            gbUnique = False
            gbSimplifier = False
            gqValeurAttribut = Nothing
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
                   ByVal sNomAttribut As String, ByVal sExpression As String,
                   Optional ByVal bLimite As Boolean = False, Optional ByVal bUnique As Boolean = False)
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression
            gbLimite = bLimite
            gbLimiteRel = False
            gbUnique = bUnique
            gbSimplifier = False
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
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres
            gpFeatureLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        gbLimite = Nothing
        gbLimiteRel = Nothing
        gbUnique = Nothing
        gbSimplifier = Nothing
        gqValeurAttribut = Nothing
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
            Nom = "NombreIntersection"
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
            If gbLimite And gbLimiteRel Then
                Parametres = Parametres & " LIMITE-LIMITE"
            ElseIf gbLimite Then
                Parametres = Parametres & " LIMITE"
            ElseIf gbLimiteRel Then
                Parametres = Parametres & " -LIMITE"
            End If
            If gbUnique Then Parametres = Parametres & " UNIQUE"
            If gbSimplifier Then Parametres = Parametres & " SIMPLIFIER"
            'Définir le paramètres des valeurs d'attributs à comparer
            If gqValeurAttribut IsNot Nothing Then
                'définir le début des paramètres de noms d'attributs à comparer
                Parametres = Parametres & " ATTRIBUT="
                'Traiter tous les noms d'attributs
                For i = 1 To gqValeurAttribut.Count
                    'Vérifier si c'est le premier attribut
                    If i = 1 Then
                        'Défirnir l'attribut
                        Parametres = Parametres & gqValeurAttribut.Item(i).ToString
                        'Si ce n'est pas le premier attribut
                    Else
                        'Défirnir l'attribut
                        Parametres = Parametres & "," & gqValeurAttribut.Item(i).ToString
                    End If
                Next
            End If
        End Get
        Set(ByVal value As String)
            'Déclarer les variables de travail
            Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression régulière, 2:LIMITE ou UNIQUE, 3:UNIQUE ou LIMITE

            Try
                'Extraire les paramètres
                params = value.Split(CChar(" "))
                'Vérifier si les deux paramètres sont présents
                If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION")

                'Définir les valeurs par défaut
                gsNomAttribut = params(0)
                gsExpression = params(1)
                gbLimite = False
                gbUnique = False
                gbSimplifier = False
                gqValeurAttribut = Nothing

                'Vérifier si les d'autres paramètres sont présents
                For i = 2 To params.Length - 1
                    'Valider la présence du paramètre LIMITE
                    If params(i) = "LIMITE" Then
                        gbLimite = True
                        'Valider la présence du paramètre -LIMITE
                    ElseIf params(i) = "-LIMITE" Then
                        gbLimiteRel = True
                        'Valider la présence du paramètre LIMITE-LIMITE
                    ElseIf params(i) = "LIMITE-LIMITE" Then
                        gbLimite = True
                        gbLimiteRel = True
                        'Valider la présence du paramètre UNIQUE
                    ElseIf params(i) = "UNIQUE" Then
                        gbUnique = True
                        'Valider la présence du paramètre SIMPLIFIER
                    ElseIf params(i) = "SIMPLIFIER" Then
                        gbSimplifier = True
                        'Valider la présence du paramètre ATTRIBUT=
                    ElseIf params(i).Contains("ATTRIBUT=") Then
                        'Permet d'indiquer la présence des valeurs d'attributs à comparer
                        gqValeurAttribut = New Collection
                        'Définir la liste des attributs à comparer
                        Dim atts() As String = params(i).Replace("ATTRIBUT=", "").Split(CChar(","))
                        'Traiter tous les attributs
                        For j = 0 To atts.Length - 1
                            'Ajouter l'attribut dans la liste des attributs à comparer
                            gqValeurAttribut.Add(atts(j))
                        Next
                        'Vider la mémoire
                        atts = Nothing
                        'Si le paramètre est invalide
                    Else
                        Err.Raise(1, , "Seul les paramètres LIMITE, -LIMITE, LIMITE-LIMITE, UNIQUE, SIMPLIFIER ou ATTRIBUT= sont possibles : " & params(i))
                    End If
                Next

            Catch ex As Exception
                'Retourner l'erreur
                Throw ex
            Finally
                'Vider la mémoire
                params = Nothing
            End Try
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner si les limites des géométries seront à traiter.
    '''</summary>
    ''' 
    Public Property Limite() As Boolean
        Get
            Limite = gbLimite
        End Get
        Set(ByVal value As Boolean)
            gbLimite = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner si les limites des géométries en relation seront à traiter.
    '''</summary>
    ''' 
    Public Property LimiteRel() As Boolean
        Get
            LimiteRel = gbLimiteRel
        End Get
        Set(ByVal value As Boolean)
            gbLimiteRel = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner si chaque élément en relation seront à traiter de façon unique.
    '''</summary>
    ''' 
    Public Property Unique() As Boolean
        Get
            Unique = gbUnique
        End Get
        Set(ByVal value As Boolean)
            gbUnique = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner si les géométries d'intersection sont simplifiées.
    '''</summary>
    ''' 
    Public Property Simplifier() As Boolean
        Get
            Simplifier = gbSimplifier
        End Get
        Set(ByVal value As Boolean)
            gbSimplifier = value
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"
    '''<summary>
    ''' Fonction qui permet de retourner la liste des paramètres possibles.
    '''</summary>
    ''' 
    Public Overloads Overrides Function ListeParametres() As Collection
        Try
            'Définir la liste des paramètres par défaut
            ListeParametres = New Collection

            'Vérifier si FeatureLayer est valide
            If gpFeatureLayerSelection IsNot Nothing Then
                'Définir le paramètre pour trouver les éléments sans intersection avec d'autres éléments en relation
                ListeParametres.Add("ELEMENT ^0$")
                'Définir le paramètre pour trouver les éléments sans intersection avec d'autres éléments en relation en utilisant la limite de la géométrie
                ListeParametres.Add("ELEMENT ^0$ LIMITE")
                'Définir le paramètre pour trouver au moins 2 intersections pour chaque NODE d'un élément
                ListeParametres.Add("NODE ^[2-9]$")
                'Définir le paramètre pour trouver au moins 2 intersections pour chaque EDGE d'un élément
                ListeParametres.Add("EDGE ^[2-9]$")
                'Définir le paramètre pour trouver aucune intersection intérieure
                ListeParametres.Add("INTERIEURE ^0$")
                'Définir le paramètre pour trouver aucune intersection intérieure avec la même valeur d'attribut NAMEID
                ListeParametres.Add("INTERIEURE ^0$ ATTRIBUT=NAMEID")
                'Définir le paramètre pour trouver aucune intersection extérieure
                ListeParametres.Add("EXTERIEURE ^0$")
                'Définir le paramètre pour trouver une seule intersection de type point
                ListeParametres.Add("POINT ^1$")
                'Définir le paramètre pour trouver une seule intersection de type point simplifié
                ListeParametres.Add("POINT ^1$ SIMPLIFIER")
                'Définir le paramètre pour trouver une seule intersection de type point pour chaque élément unique en relation
                ListeParametres.Add("POINT ^1$ UNIQUE")
                'Définir le paramètre pour trouver une seule intersection de type point en traitant les limites des éléments à traiter
                ListeParametres.Add("POINT ^1$ LIMITE")
                'Définir le paramètre pour trouver une seule intersection de type point en traitant les limites des éléments en relation
                ListeParametres.Add("POINT ^1$ -LIMITE")
                'Définir le paramètre pour trouver une seule intersection de type point en traitant les limites des éléments à traiter et en relation
                ListeParametres.Add("POINT ^1$ LIMITE-LIMITE")
                'Définir le paramètre pour trouver une seule intersection de type point 
                'pour chaque élément unique en relation et en traitant les limites des éléments à traiter
                ListeParametres.Add("POINT ^1$ LIMITE UNIQUE SIMPLIFIER")
                'Définir le paramètre pour trouver une seule intersection de type ligne
                ListeParametres.Add("LIGNE ^1$")
                'Définir le paramètre pour trouver une seule intersection de type ligne
                ListeParametres.Add("LIGNE ^1$ SIMPLIFIER")
                'Définir le paramètre pour trouver une seule intersection de type ligne pour chaque élément unique en relation
                ListeParametres.Add("LIGNE ^1$ UNIQUE")
                'Définir le paramètre pour trouver une seule intersection de type ligne en traitant les limites des éléments à traiter
                ListeParametres.Add("LIGNE ^1$ LIMITE")
                'Définir le paramètre pour trouver une seule intersection de type ligne en traitant les limites des éléments en relation
                ListeParametres.Add("LIGNE ^1$ -LIMITE")
                'Définir le paramètre pour trouver une seule intersection de type ligne en traitant les limites des éléments à traiter et en relation
                ListeParametres.Add("LIGNE ^1$ LIMITE-LIMITE")
                'Définir le paramètre pour trouver une seule intersection de type ligne 
                'pour chaque élément unique en relation et en traitant les limites des éléments à traiter
                ListeParametres.Add("LIGNE ^1$ LIMITE UNIQUE SIMPLIFIER")
                'Définir le paramètre pour trouver une seule intersection de type surface
                ListeParametres.Add("SURFACE ^1$")
                'Définir le paramètre pour trouver une seule intersection de type surface
                ListeParametres.Add("SURFACE ^1$ SIMPLIFIER")
                'Définir le paramètre pour trouver une seule intersection de type surface pour chaque élément unique en relation
                ListeParametres.Add("SURFACE ^1$ UNIQUE")
                'Définir le paramètre pour trouver une seule intersection de type surface en traitant les limites des éléments à traiter
                ListeParametres.Add("SURFACE ^1$ LIMITE")
                'Définir le paramètre pour trouver une seule intersection de type surface 
                'pour chaque élément unique en relation et en traitant les limites des éléments à traiter
                ListeParametres.Add("SURFACE ^1$ LIMITE UNIQUE SIMPLIFIER")
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
    '''
    Public Overloads Overrides Function FeatureClassValide() As Boolean
        Try
            'La contrainte est valide par défaut
            FeatureClassValide = True
            gsMessage = "La contrainte est valide."

            'Vérifier si la FeatureClass est valide
            If gpFeatureLayerSelection.FeatureClass IsNot Nothing Then
                'Vérifier si la FeatureClass est de type Polyline et Polygon
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'Vérifier si l'attribut est SURFACE
                    If gsNomAttribut = "SURFACE" Then
                        'Vérifier si le type de géométrie n'est pas Polygon
                        If gpFeatureLayerSelection.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPolygon Then
                            'Indiquer que la contrainte est invalide
                            FeatureClassValide = False
                            gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polygon."
                            Exit Function
                        End If

                        'Vérifier si l'attribut est LIGNE ou EDGE
                    ElseIf (gsNomAttribut = "LIGNE" Or gsNomAttribut = "EDGE") Then
                        'Vérifier si le type de géométrie est Point ou Multipoint 
                        If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                        Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint Then
                            'Indiquer que la contrainte est invalide
                            FeatureClassValide = False
                            gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Polyline ou Polygon."
                            Exit Function
                        End If
                    End If

                    'Vérifier si on doit traiter le paramètre LIMITE
                    If gbLimite Then
                        'Vérifier sile type est un Point ou un Multipoint
                        If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                        Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint Then
                            'Indiquer que la contrainte est invalide
                            FeatureClassValide = False
                            gsMessage = "ERREUR : Le paramètre LIMITE est invalide si le type de la FeatureClass est de type Point ou Multipoint."
                            Exit Function
                        End If
                    End If
                Else
                    'Indiquer que la contrainte est invalide
                    FeatureClassValide = False
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Point, MultiPoint, Polyline ou Polygon."
                    Exit Function
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
            AttributValide = True
            gsMessage = "La contrainte est valide."

            'Vérifier si l'attribut est valide
            If gsNomAttribut = "POINT" Or gsNomAttribut = "LIGNE" Or gsNomAttribut = "SURFACE" _
            Or gsNomAttribut = "NODE" Or gsNomAttribut = "EDGE" Or gsNomAttribut = "ELEMENT" Or gsNomAttribut = "INTERIEURE" Or gsNomAttribut = "EXTERIEURE" Then
                'Vérifier si l'attribut est SURFACE et si on doit traiter le paramètre LIMITE
                'If gsNomAttribut = "SURFACE" And gbLimite Then
                '    'Erreur de combinaison invalide
                '    gsMessage = "ERREUR : L'attribut SURFACE est invalide avec le paramètre LIMITE."
                '    AttributValide = False
                'End If
                'Vérifier si  l'attribut est EXTERIEURE et le type de géométrie est un point
                If gsNomAttribut = "EXTERIEURE" And gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                    'Erreur de combinaison invalide
                    gsMessage = "ERREUR : L'attribut EXTERIEURE est invalide avec un type de géométrie POINT."
                    AttributValide = False
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection entre deux géométries
    ''' respecte ou non l'expression régulière spécifiée.
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
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.
        Dim pSRTolerance As ISpatialReferenceTolerance = Nothing    'Interface contenant la tolérance de précision de la référence spatiale.

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

            'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer
            pSelectionSet = pFeatureSel.SelectionSet

            'Vérifier si des éléments sont sélectionnés
            If pSelectionSet.Count = 0 Then
                'Sélectionnées tous les éléments du FeatureLayer
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                pSelectionSet = pFeatureSel.SelectionSet
            End If

            'Vérifier si l'attribut est POINT
            If gsNomAttribut = "POINT" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                'Traiter le FeatureLayer et retourner les intersections points trouvées
                Selectionner = TraiterNombreIntersectionPoint(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est LIGNE
            ElseIf gsNomAttribut = "LIGNE" Then
                'Vérifier si on doit traiter les limites des géométries
                If gbLimite Or gbLimiteRel Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                    'Traiter le FeatureLayer et retourner les intersections points trouvées
                    Selectionner = TraiterNombreIntersectionPoint(pTrackCancel, bEnleverSelection)
                Else
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                    'Traiter le FeatureLayer et retourner les intersections lignes trouvées
                    Selectionner = TraiterNombreIntersectionLigne(pTrackCancel, bEnleverSelection)
                End If

                'Vérifier si l'attribut est SURFACE
            ElseIf gsNomAttribut = "SURFACE" Then
                'Vérifier si on doit traiter les limites des géométries
                If gbLimite Or gbLimiteRel Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                    'Traiter le FeatureLayer et retourner les intersections lignes trouvées
                    Selectionner = TraiterNombreIntersectionLigne(pTrackCancel, bEnleverSelection)
                Else
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolygon)

                    'Traiter le FeatureLayer et retourner les intersections surface trouvées
                    Selectionner = TraiterNombreIntersectionSurface(pTrackCancel, bEnleverSelection)
                End If

                'Vérifier si l'attribut est INTERIEURE
            ElseIf gsNomAttribut = "INTERIEURE" Then
                'Vérifier si la classe est un point
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                    'Vérifier si on traite les limites de la classe de sélection
                ElseIf gbLimite Then
                    'Vérifier si la géométrie est une surface
                    If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)
                        'Vérifier si la géométrie est une ligne
                    ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)
                        'Sinon
                    Else
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)
                    End If
                    'Si on ne traite pas la limite
                Else
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)
                End If

                'Traiter le FeatureLayer et retourner les intersections Intérieures trouvées
                Selectionner = TraiterNombreIntersectionInterieure(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est EXTERIEURE
            ElseIf gsNomAttribut = "EXTERIEURE" Then
                'Vérifier si la classe est un point
                If gbLimite Then
                    'Vérifier si la géométrie est une surface
                    If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)
                        'Vérifier si la géométrie est une ligne
                    ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)
                        'Sinon
                    Else
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)
                    End If
                    'Si on ne traite pas la limite
                Else
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)
                End If

                'Traiter le FeatureLayer et retourner les intersections Extérieures trouvées
                Selectionner = TraiterNombreIntersectionExterieure(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est NODE
            ElseIf gsNomAttribut = "NODE" Then
                'Afficher le message d'identification des points trouvés
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si la Layer de sélection est absent dans les Layers relations
                If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                    'Ajouter le layer de sélection dans les layers en relation
                    gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                End If
                'Interface pour extraire la tolérance de précision de la référence spatiale
                pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                'Création de la topologie
                pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)

                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)

                'Traiter le FeatureLayer et retourner les intersections trouvées
                Selectionner = TraiterNombreIntersectionNode(pTopologyGraph, pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est EDGE
            ElseIf gsNomAttribut = "EDGE" Then
                'Afficher le message d'identification des points trouvés
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si la Layer de sélection est absent dans les Layers relations
                If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                    'Ajouter le layer de sélection dans les layers en relation
                    gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                End If
                'Interface pour extraire la tolérance de précision de la référence spatiale
                pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                'Création de la topologie
                pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)

                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPolyline)

                'Traiter le FeatureLayer et retourner les intersections trouvées
                Selectionner = TraiterNombreIntersectionEdge(pTopologyGraph, pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est ELEMENT
            ElseIf gsNomAttribut = "ELEMENT" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                'Traiter le FeatureLayer et retourner les intersections trouvées
                Selectionner = TraiterNombreIntersectionElement(pTrackCancel, bEnleverSelection)
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
            pSelectionSet = Nothing
            pTopologyGraph = Nothing
            pSRTolerance = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'éléments qui intersecte chaque EDGE de la topologie n'est pas respecté.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterNombreIntersectionEdge(ByRef pTopologyGraph As ITopologyGraph4,
                                                   ByRef pTrackCancel As ITrackCancel,
                                                   Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un edge. 
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pZAware As IZAware = Nothing                    'Interface utilisé pour enlever le Z. 
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant les EDGES en erreur. 
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionEdge = New GeometryBag
            TraiterNombreIntersectionEdge.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet
            pFeatureSel.Clear()

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterNombreIntersectionEdge, IGeometryCollection)

            'Vérifier si la géométrie est de type Polyline ou Polygon
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
            Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Initialiser l'ajout
                    bAjouter = False

                    'Interface pour extraire les composantes
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                    'Extraire la première composante
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter toutes les composantes
                    Do Until pTopoEdge Is Nothing
                        'Interface pour extraire le nombre d'intersections
                        pEnumTopoParent = pTopoEdge.Parents()

                        'Valider le nombre d'intersections selon l'expression régulière
                        oMatch = oRegEx.Match(pEnumTopoParent.Count.ToString)

                        'Vérifier si on doit sélectionner l'élément
                        If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                            'Initialiser l'ajout
                            bAjouter = True
                            'Conserver le EDGE en erreur
                            pPolylineErr = CType(pTopoEdge.Geometry, IPolyline)
                            'Enlever le Z du EDGE
                            pZAware = CType(pPolylineErr, IZAware)
                            pZAware.DropZs()
                            pZAware.ZAware = False
                            'Ajouter les EDGES trouvés de l'élément sélectionné
                            pGeomSelColl.AddGeometry(CType(pPolylineErr, IGeometry))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & gsNomAttribut & " /NbInt=" & pEnumTopoParent.Count.ToString & " /ExpReg=" & gsExpression, _
                                                pPolylineErr, pEnumTopoParent.Count)
                        End If

                        'Extraire la prochaine composante
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop

                    'Vérifier s'il faut ajouter l'élément dans la sélection
                    If bAjouter Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                    End If

                    'Récupération de la mémoire disponble
                    pFeature = Nothing
                    pEnumTopoEdge = Nothing
                    pTopoEdge = Nothing
                    pEnumTopoParent = Nothing
                    pPolylineErr = Nothing
                    pZAware = Nothing
                    oMatch = Nothing
                    GC.Collect()

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément à traiter
                    pFeature = pFeatureCursor.NextFeature()
                Loop
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
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pEnumTopoParent = Nothing
            pZAware = Nothing
            pPolylineErr = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre de EDGE qui intersecte chaque NODE de la topologie n'est pas respecté.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    Private Function TraiterNombreIntersectionNode(ByRef pTopologyGraph As ITopologyGraph4,
                                                   ByRef pTrackCancel As ITrackCancel,
                                                   Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim oMatch As Match = Nothing                       'Object qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface utilisé pour extraire les Nodes.
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface contenant les EDGES du Node traité.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node. 
        Dim pMultiPointColl As IPointCollection = Nothing   'Interface contenant tous les nodes en erreur d'un élément.
        Dim pMultiPointErr As IMultipoint = Nothing         'Interface contenant un node en erreur.
        Dim bAjouter As Boolean = False                     'Indique si l'élément doit être ajouté dans la sélection.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionNode = New GeometryBag
            TraiterNombreIntersectionNode.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet
            pFeatureSel.Clear()

            'Interfaces pour extraire les éléments sélectionnés
            pSelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterNombreIntersectionNode, IGeometryCollection)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments sélectionnés du FeatureLayer
            Do While Not pFeature Is Nothing
                'Initialiser l'ajout
                bAjouter = False

                'Interface pour extraire les composantes
                pEnumTopoNode = pTopologyGraph.GetParentNodes(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                'Extraire la première composante
                pTopoNode = pEnumTopoNode.Next

                'Traiter toutes les composantes
                Do Until pTopoNode Is Nothing
                    'Interface pour extraire le nombre d'intersections
                    pEnumNodeEdge = pTopoNode.Edges(True)

                    'Valider le nombre d'intersections selon l'expression régulière
                    oMatch = oRegEx.Match(pEnumNodeEdge.Count.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Initialiser l'ajout
                        bAjouter = True
                        'Créer un multipoint vide
                        pMultiPointErr = New Multipoint
                        pMultiPointErr.SpatialReference = TraiterNombreIntersectionNode.SpatialReference
                        pMultiPointColl = CType(pMultiPointErr, IPointCollection)
                        'Ajouter le point trouvé
                        pMultiPointColl.AddPoint(CType(pTopoNode.Geometry, IPoint))
                        'Ajouter les points trouvés de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pMultiPointErr)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #" & gsNomAttribut & " /NbInt=" & pEnumNodeEdge.Count.ToString & " /ExpReg=" & gsExpression, _
                                            pMultiPointErr, pEnumNodeEdge.Count)
                    End If

                    'Extraire la prochaine composante
                    pTopoNode = pEnumTopoNode.Next
                Loop

                'Vérifier s'il faut ajouter l'élément dans la sélection
                If bAjouter Then
                    'Ajouter l'élément dans la sélection
                    pFeatureSel.Add(pFeature)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Récupération de la mémoire disponble
                pFeature = Nothing
                pEnumTopoNode = Nothing
                pEnumNodeEdge = Nothing
                pTopoNode = Nothing
                pMultiPointErr = Nothing
                pMultiPointColl = Nothing
                oMatch = Nothing
                GC.Collect()

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

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
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pEnumNodeEdge = Nothing
            pMultiPointColl = Nothing
            pMultiPointErr = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection intérieure respecte ou non 
    ''' l'expression spécifiée.
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
    '''<return>Les intersections intérieures entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    Private Function TraiterNombreIntersectionInterieure(ByRef pTrackCancel As ITrackCancel,
                                                         Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(1) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(1) As Integer                           'Vecteur des OIds des éléments en relation.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionInterieure = New GeometryBag
            TraiterNombreIntersectionInterieure.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterNombreIntersectionInterieure, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, gbLimite)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pSelectionSet

            'Afficher le message d'identification des points trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Initialiser les géométries des surfaces d'intersection (" & gpFeatureLayerSelection.Name & ") ..."
            'Initialiser le GeometryBag contenant les géométries d'intersections résultantes.
            InitialiserIntersection(pTrackCancel, pGeomSelColl, pGeomResColl)

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection And (gbLimite = gbLimiteRel) Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel, gbLimiteRel)
                End If

                'Afficher le message de recherche des éléments en relation
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.RelationEx(CType(pGeomRelColl, IGeometryBag), esriSpatialRelationEnum.esriSpatialRelationInteriorIntersection)

                'Vérifier si on traite les relation de façon unique
                If gbUnique Then
                    'Afficher le message d'identification des surfaces d'intersection unique
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des intérieures d'intersection unique (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreInterieureUnique(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, pTrackCancel, pGeomResColl)
                Else
                    'Afficher le message d'identification des surfaces d'intersection multiple
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des intérieures d'intersection multiple (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreInterieureMultiple(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, pTrackCancel, pGeomResColl)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments trouvés
            SelectionnerNombreIntersection(bEnleverSelection, oRegEx, iOidSel, pGeomSelColl, pTrackCancel, pGeomResColl, pNewSelectionSet)

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

            'Définir le résultat des géométries sélectionnées
            TraiterNombreIntersectionInterieure = CType(pGeomResColl, IGeometryBag)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            bMemeClasse = Nothing
            bMemeDefinition = Nothing
            bMemeSelection = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection extérieure respecte ou non l'expression spécifiée.
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
    '''<return>Les intersections extérieures entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    Private Function TraiterNombreIntersectionExterieure(ByRef pTrackCancel As ITrackCancel,
                                                         Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(1) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(1) As Integer                           'Vecteur des OIds des éléments en relation.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionExterieure = New GeometryBag
            TraiterNombreIntersectionExterieure.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterNombreIntersectionExterieure, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, gbLimite)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pSelectionSet

            'Afficher le message d'identification des points trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Initialiser les géométries des surfaces d'intersection (" & gpFeatureLayerSelection.Name & ") ..."
            'Initialiser le GeometryBag contenant les géométries d'intersections résultantes.
            InitialiserIntersection(pTrackCancel, pGeomSelColl, pGeomResColl)

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection And (gbLimite = gbLimiteRel) Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel, gbLimiteRel)
                End If

                'Afficher le message de recherche des éléments en relation
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))
                pRelResult.Subtract(pRelOpNxM.RelationEx(CType(pGeomRelColl, IGeometryBag), esriSpatialRelationEnum.esriSpatialRelationWithin))

                'Vérifier si on traite les relation de façon unique
                If gbUnique Then
                    'Afficher le message d'identification des surfaces d'intersection unique
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des extérieures d'intersection unique (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreExterieureUnique(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                Else
                    'Afficher le message d'identification des surfaces d'intersection multiple
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des extérieures d'intersection multiple (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreExterieureMultiple(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments trouvés
            SelectionnerNombreIntersection(bEnleverSelection, oRegEx, iOidSel, pGeomSelColl, pTrackCancel, pGeomResColl, pNewSelectionSet)

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

            'Définir le résultat des géométries sélectionnées
            TraiterNombreIntersectionExterieure = CType(pGeomResColl, IGeometryBag)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            bMemeClasse = Nothing
            bMemeDefinition = Nothing
            bMemeSelection = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'identifier les intersections extérieures uniques entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreExterieureUnique(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                                 ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                                 ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection extérieure avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionExterieure(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans la géométrie résultante
                    pGeomColl.AddGeometry(pGeometryInt)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
            iSel = Nothing
            iRel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections Extérieures multiples entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreExterieureMultiple(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                                   ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                                   ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionExterieure(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir le GeometryBag contenu dans la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Définir l'intersection contenu dans le GeometryBag
                    pGeomColl = CType(pGeomColl.Geometry(0), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans le Polygone
                    pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
            iSel = Nothing
            iRel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire l'intersection extérieure entre la géométrie de l'élément à traiter et la géométrie d'un élément en relation.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface contenant la géométrie de l'élément à traiter.</param>
    '''<param name="pGeometryRel">Interface contenant la géométrie d'un élément en relation.</param>
    ''' 
    '''<returns>"IGeometry" contenant les intersections intérieures entre la géométrie de l'élément traité et les géométries des éléments en relation.</returns>
    '''
    Private Function ExtraireIntersectionExterieure(ByRef pGeometry As IGeometry, ByRef pGeometryRel As IGeometry) As IGeometry
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing                 'Interface contenant la géométrie d'intersection.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Interface pour extraire l'intersection entre les deux géométries
            pTopoOp = CType(pGeometry, ITopologicalOperator2)

            'Projeter la géométrie de l'élément en relation
            pGeometryRel.Project(pGeometry.SpatialReference)

            'Trouver l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
            pGeometryInt = pTopoOp.Difference(pGeometryRel)

            'Définir l'intersection intérieure
            ExtraireIntersectionExterieure = pGeometryInt

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'initialiser le GeometryBag contenant les intersections résultantes.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub InitialiserIntersection(ByRef pTrackCancel As ITrackCancel, ByRef pGeomSelColl As IGeometryCollection, _
                                                  ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie d'intersection intérieure.
        Dim pGeometryBag As IGeometryBag = Nothing      'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Vérifier si les points d'intersection sont unique
            If gbUnique Then
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir la valeur de retour par défaut
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            Else
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir le GeometryBag vide
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Définir la géométrie vide
                    If pGeomSelColl.Geometry(i).Dimension = esriGeometryDimension.esriGeometry0Dimension Then
                        'Créer un multipoint vide
                        pGeometry = New Multipoint
                    ElseIf pGeomSelColl.Geometry(i).Dimension = esriGeometryDimension.esriGeometry1Dimension Then
                        'Créer une polyligne vide
                        pGeometry = New Polyline
                    ElseIf pGeomSelColl.Geometry(i).Dimension = esriGeometryDimension.esriGeometry2Dimension Then
                        'Créer un polygone vide
                        pGeometry = New Polygon
                    End If
                    'Définir la référence spatiale
                    pGeometry.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Interface pour ajouter le polygon dans le GeometryBag
                    pGeomColl = CType(pGeometryBag, IGeometryCollection)
                    'Ajouter le polygon vide dans le GeometryBag
                    pGeomColl.AddGeometry(pGeometry)
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            End If

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections intérieures uniques entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="pFeatureLayerSel"> Interface contenant la classe d'éléments à traiter.</param>
    '''<param name="pFeatureLayerRel"> Interface contenant la classe d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreInterieureUnique(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                                 ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal pFeatureLayerSel As IFeatureLayer, ByVal pFeatureLayerRel As IFeatureLayer, _
                                                 ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeature = Nothing           'Interface contenant un élément de sélection.
        Dim pFeatureRel As IFeature = Nothing           'Interface contenant un élément en relation.
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Vérifier la présence des attributs à comparer
                    If gqValeurAttribut IsNot Nothing Then
                        'Définir l'élément de sélection
                        pFeatureSel = pFeatureLayerSel.FeatureClass.GetFeature(iOidSel(iSel))
                        'Définir l'élément en relation
                        pFeatureRel = pFeatureLayerRel.FeatureClass.GetFeature(iOidRel(iRel))
                        'Vérifier si les attributs sont identiques
                        If ValeurAttributIdentique(pFeatureSel, pFeatureRel, gqValeurAttribut) Then
                            'Extraire la géométrie d'intersection avec l'élément en relation
                            pGeometryInt = ExtraireIntersectionInterieure(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                            'Définir la géométrie résultante
                            pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                            'Ajouter la géométrie d'intersection dans la géométrie résultante
                            pGeomColl.AddGeometry(pGeometryInt)
                        End If

                        'Si aucun attribut à comparer
                    Else
                        'Extraire la géométrie d'intersection avec l'élément en relation
                        pGeometryInt = ExtraireIntersectionInterieure(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                        'Définir la géométrie résultante
                        pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                        'Ajouter la géométrie d'intersection dans la géométrie résultante
                        pGeomColl.AddGeometry(pGeometryInt)
                    End If
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pFeatureRel = Nothing
            pGeometryInt = Nothing
            pGeomColl = Nothing
            iSel = Nothing
            iRel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections intérieures multiples entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="pFeatureLayerSel"> Interface contenant la classe d'éléments à traiter.</param>
    '''<param name="pFeatureLayerRel"> Interface contenant la classe d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreInterieureMultiple(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                                   ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal pFeatureLayerSel As IFeatureLayer, ByVal pFeatureLayerRel As IFeatureLayer, _
                                                   ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeature = Nothing           'Interface contenant un élément de sélection.
        Dim pFeatureRel As IFeature = Nothing           'Interface contenant un élément en relation.
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Vérifier la présence des attributs à comparer
                    If gqValeurAttribut IsNot Nothing Then
                        'Définir l'élément de sélection
                        pFeatureSel = pFeatureLayerSel.FeatureClass.GetFeature(iOidSel(iSel))
                        'Définir l'élément en relation
                        pFeatureRel = pFeatureLayerRel.FeatureClass.GetFeature(iOidRel(iRel))
                        'Vérifier si les attributs sont identiques
                        If ValeurAttributIdentique(pFeatureSel, pFeatureRel, gqValeurAttribut) Then
                            'Extraire la géométrie d'intersection avec l'élément en relation
                            pGeometryInt = ExtraireIntersectionInterieure(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                            'Définir le GeometryBag contenu dans la géométrie résultante
                            pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                            'Définir l'intersection contenu dans le GeometryBag
                            pGeomColl = CType(pGeomColl.Geometry(0), IGeometryCollection)
                            'Ajouter la géométrie d'intersection dans le Polygone
                            pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))
                        End If

                        'Si aucun attribut à comparer
                    Else
                        'Extraire la géométrie d'intersection avec l'élément en relation
                        pGeometryInt = ExtraireIntersectionInterieure(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                        'Définir le GeometryBag contenu dans la géométrie résultante
                        pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                        'Définir l'intersection contenu dans le GeometryBag
                        pGeomColl = CType(pGeomColl.Geometry(0), IGeometryCollection)
                        'Ajouter la géométrie d'intersection dans le Polygone
                        pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))
                    End If
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pFeatureRel = Nothing
            pGeometryInt = Nothing
            pGeomColl = Nothing
            iSel = Nothing
            iRel = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'indiquer si les valeurs d'attributs sont identiques.
    '''</summary>
    ''' 
    '''<param name="pFeatureSel">Interface contenant un élément de sélection.</param>
    '''<param name="pFeatureRel">Interface contenant un élément en relation.</param>
    '''<param name="qCollAttribut">Collection des attributs à comparer.</param>
    ''' 
    '''<returns> Boolean indiquant si les valeurs d'attributs sont identiques.</returns>
    ''' 
    Private Function ValeurAttributIdentique(ByVal pFeatureSel As IFeature, ByVal pFeatureRel As IFeature, ByVal qCollAttribut As Collection) As Boolean
        'Définir les variables de travail
        Dim iPosSel As Integer = -1     'Position de l'attribut dans l'élément de sélection.
        Dim iPosRel As Integer = -1     'Position de l'attribut dans l'élément en relation.
        Dim sNomAttribut As String = "" 'Nom de l'attribut à valider.

        'Par défaut, les valeurs d'attributs sont identiques
        ValeurAttributIdentique = True

        Try
            'Traiter tous les attributs
            For i = 1 To qCollAttribut.Count
                'Définir le nom de l'attribut à valider
                sNomAttribut = qCollAttribut.Item(i).ToString

                'Extraire la position de l'attribut dans l'élément de sélection.
                iPosSel = pFeatureSel.Fields.FindField(sNomAttribut)
                'Sortir si l'attribut n'existe pas
                If iPosSel = -1 Then Return False

                'Extraire la position de l'attribut dans l'élément en relation.
                iPosRel = pFeatureRel.Fields.FindField(sNomAttribut)
                'Sortir si l'attribut n'existe pas
                If iPosRel = -1 Then Return False

                'Sortir si les valeurs d'attributs sont différentes
                If pFeatureSel.Value(iPosSel).ToString <> pFeatureRel.Value(iPosRel).ToString Then Return False
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte ou non le nombre d'intersections spécifiés.
    '''</summary>
    ''' 
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''<param name="oRegEx"> Objet utilisé pour vérifier si la valeur respecte l'expression régulière.</param>
    '''<param name="iOidSel"> Vecteur des OIds des éléments à traiter.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    '''<param name="pNewSelectionSet"> Interface pour sélectionner les éléments.</param>
    ''' 
    Private Sub SelectionnerNombreIntersection(ByVal bEnleverSelection As Boolean, ByRef oRegEx As Regex, ByRef iOidSel() As Integer, _
                                               ByRef pGeomSelColl As IGeometryCollection, ByRef pTrackCancel As ITrackCancel, _
                                               ByRef pGeomResColl As IGeometryCollection, ByRef pNewSelectionSet As ISelectionSet)
        'Déclarer les variables de travail
        Dim oMatch As Match = Nothing                       'Objet qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter l'intersection des géométries.
        Dim pNewGeomColl As IGeometryCollection = Nothing   'Interface pour ajouter l'intersection des géométries.
        Dim pIntersectColl As IGeometryCollection = Nothing 'Interface pour extraire le nombre d'intersection des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Définir la valeur de retour par défaut
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            'Définir le nouveau GeometryBag des géométries résultantes
            pNewGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter toutes les gométries d'intersection
            For i = 0 To pGeomResColl.GeometryCount - 1
                'Définir le GeometryBag contenu dans la géométrie résultante
                pGeomColl = CType(pGeomResColl.Geometry(i), IGeometryCollection)

                'Si la géométrie des intersections est vide
                If pGeomColl.GeometryCount = 0 Then
                    If gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Ajouter un polygon vide 
                        pGeomColl.AddGeometry(New Polygon)
                    ElseIf gpFeatureClassErreur.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Ajouter une polyline vide 
                        pGeomColl.AddGeometry(New Polyline)
                    Else
                        'Ajouter un multipoint vide 
                        pGeomColl.AddGeometry(New Multipoint)
                    End If
                End If

                'Traiter toutes les géométries du GeometryBag
                For j = 0 To pGeomColl.GeometryCount - 1
                    'Définir l'intersection intérieure contenu dans le GeometryBag
                    pIntersectColl = CType(pGeomColl.Geometry(j), IGeometryCollection)

                    'Vérifier si on doit simplifier
                    If gbSimplifier Then
                        'Simplifier la géométrie
                        pTopoOp = CType(pIntersectColl, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()
                    End If

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pIntersectColl.GeometryCount.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pNewSelectionSet.Add(iOidSel(i))
                        'Vérifier si la géométrie d'intersection est vide
                        If pIntersectColl.GeometryCount = 0 Then
                            'Ajouter l'enveloppe de la géométrie de l'élément
                            pNewGeomColl.AddGeometry(pGeomSelColl.Geometry(i))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbIntersection=" & pIntersectColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                pGeomSelColl.Geometry(i), pIntersectColl.GeometryCount)
                        Else
                            'Ajouter la géométrie d'intersection
                            pNewGeomColl.AddGeometry(CType(pIntersectColl, IGeometry))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbIntersection=" & pIntersectColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                CType(pIntersectColl, IGeometry), pIntersectColl.GeometryCount)
                        End If
                    End If
                Next

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Remplacer les géométries résultantes
            pGeomResColl = pNewGeomColl

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oMatch = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
            pNewGeomColl = Nothing
            pIntersectColl = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire l'intersection intérieure entre la géométrie de l'élément à traiter et la géométrie d'un élément en relation.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface contenant la géométrie de l'élément à traiter.</param>
    '''<param name="pGeometryRel">Interface contenant la géométrie d'un élément en relation.</param>
    ''' 
    '''<returns>"IGeometry" contenant les intersections intérieures entre la géométrie de l'élément traité et les géométries des éléments en relation.</returns>
    '''
    Private Function ExtraireIntersectionInterieure(ByRef pGeometry As IGeometry, ByRef pGeometryRel As IGeometry) As IGeometry
        'Déclarer les variables de travail
        Dim pClone As IClone = Nothing                          'Interface pourcloner une géométrie.
        Dim pGeometryInt As IGeometry = Nothing                 'Interface contenant la géométrie d'intersection.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Interface pour cloner la géométrie à traiter
            pClone = CType(pGeometry, IClone)
            'Cloner la géométrie à traiter
            ExtraireIntersectionInterieure = CType(pClone.Clone, IGeometry)
            'Par défaut la géométrie résultant est vide
            ExtraireIntersectionInterieure.SetEmpty()

            'Vérifier si la géométrie est un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Créer un nouveau multipoint vide
                ExtraireIntersectionInterieure = GeometrieToMultiPoint(pGeometry)

                'Si la géométrie n'est pas un point
            Else
                'Interface pour extraire l'intersection entre les deux géométries
                pTopoOp = CType(pGeometry, ITopologicalOperator2)

                'Projeter la géométrie de l'élément en relation
                pGeometryRel.Project(pGeometry.SpatialReference)

                Try
                    'Trouver l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
                    If pGeometry.Dimension = esriGeometryDimension.esriGeometry0Dimension Then
                        pGeometryInt = pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry0Dimension)
                    ElseIf pGeometry.Dimension = esriGeometryDimension.esriGeometry1Dimension Then
                        pGeometryInt = pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry1Dimension)
                    ElseIf pGeometry.Dimension = esriGeometryDimension.esriGeometry2Dimension Then
                        pGeometryInt = pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry2Dimension)
                    End If

                    'Définir l'intersection intérieure
                    ExtraireIntersectionInterieure = pGeometryInt
                Catch ex As Exception
                    'On ne fait rien
                End Try
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pClone = Nothing
            pGeometryInt = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection de type surface respecte ou non 
    ''' l'expression spécifiée.
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
    '''<return>Les intersections entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    Private Function TraiterNombreIntersectionSurface(ByRef pTrackCancel As ITrackCancel,
                                                      Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(1) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(1) As Integer                           'Vecteur des OIds des éléments en relation.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionSurface = New GeometryBag
            TraiterNombreIntersectionSurface.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterNombreIntersectionSurface, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, gbLimite)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pSelectionSet

            'Afficher le message d'identification des points trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Initialiser les géométries des surfaces d'intersection (" & gpFeatureLayerSelection.Name & ") ..."
            'Initialiser le GeometryBag contenant les surfaces d'intersections résultantes.
            InitialiserSurfaceIntersection(pTrackCancel, pGeomSelColl, pGeomResColl)

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection And (gbLimite = gbLimiteRel) Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel, gbLimiteRel)
                End If

                'Afficher le message de recherche des éléments en relation
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                'Vérifier si on traite les relation de façon unique
                If gbUnique Then
                    'Afficher le message d'identification des surfaces d'intersection unique
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des surfaces d'intersection unique (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreSurfaceUnique(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                Else
                    'Afficher le message d'identification des surfaces d'intersection multiple
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des surfaces d'intersection multiple (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreSurfaceMultiple(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments trouvés
            SelectionnerNombreSurface(bEnleverSelection, oRegEx, iOidSel, pGeomSelColl, pTrackCancel, pGeomResColl, pNewSelectionSet)

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

            'définir le résultat des géométries sélectionnées
            TraiterNombreIntersectionSurface = CType(pGeomResColl, IGeometryBag)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'initialiser le GeometryBag contenant les surfaces d'intersections résultantes.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub InitialiserSurfaceIntersection(ByRef pTrackCancel As ITrackCancel, ByRef pGeomSelColl As IGeometryCollection, _
                                               ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pPolygon As IPolygon = Nothing              'Interface contenant le Polygone d'intersection.
        Dim pGeometryBag As IGeometryBag = Nothing      'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Vérifier si les points d'intersection sont unique
            If gbUnique Then
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir la valeur de retour par défaut
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            Else
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir le GeometryBag vide
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Définir le polygon vide
                    pPolygon = New Polygon
                    pPolygon.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Interface pour ajouter le polygon dans le GeometryBag
                    pGeomColl = CType(pGeometryBag, IGeometryCollection)
                    'Ajouter le polygon vide dans le GeometryBag
                    pGeomColl.AddGeometry(pPolygon)
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            End If

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPolygon = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections surfaces uniques entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreSurfaceUnique(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                              ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                              ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionSurface(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans la géométrie résultante
                    pGeomColl.AddGeometry(pGeometryInt)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections surfaces multiples entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreSurfaceMultiple(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                                ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                                ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionSurface(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir le GeometryBag contenu dans la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Définir le Polygone contenu dans le GeometryBag
                    pGeomColl = CType(pGeomColl.Geometry(0), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans le Polygone
                    pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte ou non le nombre d'intersections surface spécifiés.
    '''</summary>
    ''' 
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''<param name="oRegEx"> Objet utilisé pour vérifier si la valeur respecte l'expression régulière.</param>
    '''<param name="iOidSel"> Vecteur des OIds des éléments à traiter.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    '''<param name="pNewSelectionSet"> Interface pour sélectionner les éléments.</param>
    ''' 
    Private Sub SelectionnerNombreSurface(ByVal bEnleverSelection As Boolean, ByRef oRegEx As Regex, ByRef iOidSel() As Integer, _
                                          ByRef pGeomSelColl As IGeometryCollection, ByRef pTrackCancel As ITrackCancel, _
                                          ByRef pGeomResColl As IGeometryCollection, ByRef pNewSelectionSet As ISelectionSet)
        'Déclarer les variables de travail
        Dim oMatch As Match = Nothing                       'Objet qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter l'intersection des géométries.
        Dim pNewGeomColl As IGeometryCollection = Nothing   'Interface pour ajouter l'intersection des géométries.
        Dim pSurfaceColl As IGeometryCollection = Nothing   'Interface pour extraire le nombre d'intersection des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Définir la valeur de retour par défaut
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            'Définir le nouveau GeometryBag des géométries résultantes
            pNewGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter toutes les gométries d'intersection
            For i = 0 To pGeomResColl.GeometryCount - 1
                'Définir le GeometryBag contenu dans la géométrie résultante
                pGeomColl = CType(pGeomResColl.Geometry(i), IGeometryCollection)

                'Ajouter un polygon vide si la géométrie des intersections est vide
                If pGeomColl.GeometryCount = 0 Then pGeomColl.AddGeometry(New Polygon)

                'Traiter toutes les géométries du GeometryBag
                For j = 0 To pGeomColl.GeometryCount - 1
                    'Définir le Polygon contenu dans le GeometryBag
                    pSurfaceColl = CType(pGeomColl.Geometry(j), IGeometryCollection)

                    'Vérifier si on doit simplifier
                    If gbSimplifier Then
                        'Simplifier la géométrie
                        pTopoOp = CType(pSurfaceColl, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()
                    End If

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pSurfaceColl.GeometryCount.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pNewSelectionSet.Add(iOidSel(i))
                        'Vérifier si la géométrie d'intersection est vide
                        If pSurfaceColl.GeometryCount = 0 Then
                            'Ajouter l'enveloppe de la géométrie de l'élément
                            pNewGeomColl.AddGeometry(pGeomSelColl.Geometry(i))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbInt=" & pSurfaceColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                pGeomSelColl.Geometry(i), pSurfaceColl.GeometryCount)
                        Else
                            'Ajouter la géométrie d'intersection
                            pNewGeomColl.AddGeometry(CType(pSurfaceColl, IGeometry))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbInt=" & pSurfaceColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                CType(pSurfaceColl, IGeometry), pSurfaceColl.GeometryCount)
                        End If
                    End If
                Next

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Remplacer les géométries résultantes
            pGeomResColl = pNewGeomColl

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oMatch = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
            pNewGeomColl = Nothing
            pSurfaceColl = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire l'intersection de type surface entre la géométrie de l'élément à traiter et la géométrie d'un élément en relation.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface contenant la géométrie de l'élément à traiter.</param>
    '''<param name="pGeometryRel">Interface contenant la géométrie d'un élément en relation.</param>
    ''' 
    '''<returns>"IPolygon" contenant les intersections de type point entre la géométrie de l'élément traité et les géométries des éléments en relation.</returns>
    '''
    Private Function ExtraireIntersectionSurface(ByRef pGeometry As IGeometry, ByRef pGeometryRel As IGeometry) As IPolygon
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing                 'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour ajouter l'intersection des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Définir la valeur de retour par défaut
            ExtraireIntersectionSurface = New Polygon
            ExtraireIntersectionSurface.SpatialReference = pGeometry.SpatialReference
            ExtraireIntersectionSurface.SnapToSpatialReference()
            pGeomColl = CType(ExtraireIntersectionSurface, IGeometryCollection)

            'Interface pour extraire l'intersection entre les deux géométries
            pTopoOp = CType(pGeometry, ITopologicalOperator2)

            'Projeter la géométrie de l'élément en relation
            pGeometryRel.Project(pGeometry.SpatialReference)

            'Trouver l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
            pGeometryInt = pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry2Dimension)

            'Ajouter l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
            pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection de type ligne respecte ou non 
    ''' l'expression spécifiée.
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
    '''<return>Les intersections entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    Private Function TraiterNombreIntersectionLigne(ByRef pTrackCancel As ITrackCancel,
                                                    Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(1) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(1) As Integer                           'Vecteur des OIds des éléments en relation.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionLigne = New GeometryBag
            TraiterNombreIntersectionLigne.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterNombreIntersectionLigne, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, gbLimite)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pSelectionSet

            'Afficher le message d'identification des points trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Initialiser les géométries des lignes d'intersection (" & gpFeatureLayerSelection.Name & ") ..."
            'Initialiser le GeometryBag contenant les lignes d'intersections résultantes.
            InitialiserLigneIntersection(pTrackCancel, pGeomSelColl, pGeomResColl)

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection And (gbLimite = gbLimiteRel) Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel, gbLimiteRel)
                End If

                'Afficher le message de recherche des éléments en relation
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                'Vérifier si on traite les relation de façon unique
                If gbUnique Then
                    'Afficher le message d'identification des lignes d'intersection unique
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des lignes d'intersection unique (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreLigneUnique(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                Else
                    'Afficher le message d'identification des lignes d'intersection multiple
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des lignes d'intersection multiple (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombreLigneMultiple(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments trouvés
            SelectionnerNombreLigne(bEnleverSelection, oRegEx, iOidSel, pGeomSelColl, pTrackCancel, pGeomResColl, pNewSelectionSet)

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

            'définir le résultat des géométries sélectionnées
            TraiterNombreIntersectionLigne = CType(pGeomResColl, IGeometryBag)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'initialiser le GeometryBag contenant les lignes d'intersections résultantes.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub InitialiserLigneIntersection(ByRef pTrackCancel As ITrackCancel, ByRef pGeomSelColl As IGeometryCollection, _
                                             ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing            'Interface contenant la Polyline d'intersection.
        Dim pGeometryBag As IGeometryBag = Nothing      'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Vérifier si les points d'intersection sont unique
            If gbUnique Then
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir la valeur de retour par défaut
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            Else
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir le GeometryBag vide
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Définir le Polyline vide
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Interface pour ajouter le Polyline dans le GeometryBag
                    pGeomColl = CType(pGeometryBag, IGeometryCollection)
                    'Ajouter la polyline vide dans le GeometryBag
                    pGeomColl.AddGeometry(pPolyline)
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            End If

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections lignes uniques entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreLigneUnique(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                            ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                            ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionLigne(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans la géométrie résultante
                    pGeomColl.AddGeometry(pGeometryInt)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections lignes multiples entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombreLigneMultiple(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                              ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                              ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionLigne(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir le GeometryBag contenu dans la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Définir le polyline contenu dans le GeometryBag
                    pGeomColl = CType(pGeomColl.Geometry(0), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans le Polyline
                    pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte ou non le nombre d'intersections ligne spécifiés.
    '''</summary>
    ''' 
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''<param name="oRegEx"> Objet utilisé pour vérifier si la valeur respecte l'expression régulière.</param>
    '''<param name="iOidSel"> Vecteur des OIds des éléments à traiter.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    '''<param name="pNewSelectionSet"> Interface pour sélectionner les éléments.</param>
    ''' 
    Private Sub SelectionnerNombreLigne(ByVal bEnleverSelection As Boolean, ByRef oRegEx As Regex, ByRef iOidSel() As Integer, _
                                        ByRef pGeomSelColl As IGeometryCollection, ByRef pTrackCancel As ITrackCancel, _
                                        ByRef pGeomResColl As IGeometryCollection, ByRef pNewSelectionSet As ISelectionSet)
        'Déclarer les variables de travail
        Dim oMatch As Match = Nothing                       'Objet qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter l'intersection des géométries.
        Dim pNewGeomColl As IGeometryCollection = Nothing   'Interface pour ajouter l'intersection des géométries.
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface pour extraire le nombre d'intersection des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Définir la valeur de retour par défaut
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            'Définir le nouveau GeometryBag des géométries résultantes
            pNewGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter toutes les gométries d'intersection
            For i = 0 To pGeomResColl.GeometryCount - 1
                'Définir le GeometryBag contenu dans la géométrie résultante
                pGeomColl = CType(pGeomResColl.Geometry(i), IGeometryCollection)

                'Ajouter une polyline vide si la géométrie des intersections est vide
                If pGeomColl.GeometryCount = 0 Then pGeomColl.AddGeometry(New Polyline)

                'Traiter toutes les géométries du GeometryBag
                For j = 0 To pGeomColl.GeometryCount - 1
                    'Définir le Polyline contenu dans le GeometryBag
                    pLigneColl = CType(pGeomColl.Geometry(j), IGeometryCollection)

                    'Vérifier si on doit simplifier
                    If gbSimplifier Then
                        'Simplifier la géométrie
                        pTopoOp = CType(pLigneColl, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()
                    End If

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pLigneColl.GeometryCount.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pNewSelectionSet.Add(iOidSel(i))
                        'Vérifier si la géométrie d'intersection est vide
                        If pLigneColl.GeometryCount = 0 Then
                            'Ajouter l'enveloppe de la géométrie de l'élément
                            pNewGeomColl.AddGeometry(pGeomSelColl.Geometry(i))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbInt=" & pLigneColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                GeometrieToPolyline(pGeomSelColl.Geometry(i)), pLigneColl.GeometryCount)
                        Else
                            'Ajouter la géométrie d'intersection
                            pNewGeomColl.AddGeometry(CType(pLigneColl, IGeometry))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbInt=" & pLigneColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                GeometrieToPolyline(CType(pLigneColl, IGeometry)), pLigneColl.GeometryCount)
                        End If
                    End If
                Next

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Remplacer les géométries résultantes
            pGeomResColl = pNewGeomColl

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oMatch = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
            pNewGeomColl = Nothing
            pLigneColl = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire l'intersection de type point entre la géométrie de l'élément à traiter et la géométrie d'un élément en relation.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface contenant la géométrie de l'élément à traiter.</param>
    '''<param name="pGeometryRel">Interface contenant la géométrie d'un élément en relation.</param>
    ''' 
    '''<returns>"IPolyline" contenant les intersections de type point entre la géométrie de l'élément traité et les géométries des éléments en relation.</returns>
    '''
    Private Function ExtraireIntersectionLigne(ByRef pGeometry As IGeometry, ByRef pGeometryRel As IGeometry) As IPolyline
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing                 'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour ajouter l'intersection des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing           'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Définir la valeur de retour par défaut
            ExtraireIntersectionLigne = New Polyline
            ExtraireIntersectionLigne.SpatialReference = pGeometry.SpatialReference
            ExtraireIntersectionLigne.SnapToSpatialReference()
            pGeomColl = CType(ExtraireIntersectionLigne, IGeometryCollection)

            'Interface pour extraire l'intersection entre les deux géométries
            pTopoOp = CType(pGeometry, ITopologicalOperator2)

            'Projeter la géométrie de l'élément en relation
            pGeometryRel.Project(pGeometry.SpatialReference)

            'Trouver l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
            pGeometryInt = pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry1Dimension)

            'Ajouter l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
            pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection de type point respecte ou non 
    ''' l'expression spécifiée.
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
    '''<return>Les intersections entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    Private Function TraiterNombreIntersectionPoint(ByRef pTrackCancel As ITrackCancel,
                                                    Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(1) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(1) As Integer                           'Vecteur des OIds des éléments en relation.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionPoint = New GeometryBag
            TraiterNombreIntersectionPoint.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterNombreIntersectionPoint, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, gbLimite)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pSelectionSet

            'Afficher le message d'identification des points trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Initialiser les géométries des points d'intersection (" & gpFeatureLayerSelection.Name & ") ..."
            'Initialiser le GeometryBag contenant les points d'intersections résultantes.
            InitialiserPointIntersection(pTrackCancel, pGeomSelColl, pGeomResColl)

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection And (gbLimite = gbLimiteRel) Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel, gbLimiteRel)
                End If

                'Afficher le message de recherche des éléments en relation
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                'Vérifier si on traite les relation de façon unique
                If gbUnique Then
                    'Afficher le message d'identification des points trouvés
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des points d'intersection unique (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombrePointUnique(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                Else
                    'Afficher le message d'identification des points trouvés
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des points d'intersection multiple (" & pFeatureLayerRel.Name & ") ..."
                    'Identifier les éléments trouvés
                    IdentifierNombrePointMultiple(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, pGeomResColl)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments trouvés
            SelectionnerNombrePoint(bEnleverSelection, oRegEx, iOidSel, pGeomSelColl, pTrackCancel, pGeomResColl, pNewSelectionSet)

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

            'définir le résultat des géométries sélectionnées
            TraiterNombreIntersectionPoint = CType(pGeomResColl, IGeometryBag)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'initialiser le GeometryBag contenant les points d'intersections résultantes.
    '''</summary>
    ''' 
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub InitialiserPointIntersection(ByRef pTrackCancel As ITrackCancel, ByRef pGeomSelColl As IGeometryCollection, _
                                             ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pMultipoint As IMultipoint = Nothing        'Interface contenant le Multipoint d'intersection.
        Dim pGeometryBag As IGeometryBag = Nothing      'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Vérifier si les points d'intersection sont unique
            If gbUnique Then
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir la valeur de retour par défaut
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            Else
                'Initialiser tous les résultats d'intersection à vide
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Définir le GeometryBag vide
                    pGeometryBag = New GeometryBag
                    pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Définir le multipoint vide
                    pMultipoint = New Multipoint
                    pMultipoint.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
                    'Interface pour ajouter le multipoint dans le GeometryBag
                    pGeomColl = CType(pGeometryBag, IGeometryCollection)
                    'Ajouter le multipoint vide dans le GeometryBag
                    pGeomColl.AddGeometry(pMultipoint)
                    'Ajouter le Bag vide dans les géométries résultantes
                    pGeomResColl.AddGeometry(pGeometryBag)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            End If

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pMultipoint = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections points uniques entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombrePointUnique(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                            ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                            ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionPoint(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans la géométrie résultante
                    pGeomColl.AddGeometry(pGeometryInt)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier les intersections points multiples entre les éléments du FeatureLayer et les éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    ''' 
    Private Sub IdentifierNombrePointMultiple(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection, _
                                              ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                              ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing         'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter l'intersection des géométries.
        Dim iSel As Integer = -1                        'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                        'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments qui intersecte
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Extraire la géométrie d'intersection avec l'élément en relation
                    pGeometryInt = ExtraireIntersectionPoint(pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel))
                    'Définir le GeometryBag contenu dans la géométrie résultante
                    pGeomColl = CType(pGeomResColl.Geometry(iSel), IGeometryCollection)
                    'Définir le MultiPoint contenu dans le GeometryBag
                    pGeomColl = CType(pGeomColl.Geometry(0), IGeometryCollection)
                    'Ajouter la géométrie d'intersection dans le MultiPoint
                    pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte ou non le nombre d'intersections point spécifiés.
    '''</summary>
    ''' 
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''<param name="oRegEx"> Objet utilisé pour vérifier si la valeur respecte l'expression régulière.</param>
    '''<param name="iOidSel"> Vecteur des OIds des éléments à traiter.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    '''<param name="pNewSelectionSet"> Interface pour sélectionner les éléments.</param>
    ''' 
    Private Sub SelectionnerNombrePoint(ByVal bEnleverSelection As Boolean, ByRef oRegEx As Regex, ByRef iOidSel() As Integer, _
                                        ByRef pGeomSelColl As IGeometryCollection, ByRef pTrackCancel As ITrackCancel, _
                                        ByRef pGeomResColl As IGeometryCollection, ByRef pNewSelectionSet As ISelectionSet)
        'Déclarer les variables de travail
        Dim oMatch As Match = Nothing                       'Objet qui permet d'indiquer si la valeur respecte l'expression régulière.
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant l'intersection des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter l'intersection des géométries.
        Dim pNewGeomColl As IGeometryCollection = Nothing   'Interface pour ajouter l'intersection des géométries.
        Dim pPointColl As IGeometryCollection = Nothing     'Interface pour extraire le nombre d'intersection des géométries.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Définir la valeur de retour par défaut
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            'Définir le nouveau GeometryBag des géométries résultantes
            pNewGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter toutes les gométries d'intersection
            For i = 0 To pGeomResColl.GeometryCount - 1
                'Définir le GeometryBag contenu dans la géométrie résultante
                pGeomColl = CType(pGeomResColl.Geometry(i), IGeometryCollection)

                'Ajouter un multipoint vide si la géométrie des intersections est vide
                If pGeomColl.GeometryCount = 0 Then pGeomColl.AddGeometry(New Multipoint)

                'Traiter toutes les géométries du GeometryBag
                For j = 0 To pGeomColl.GeometryCount - 1
                    'Définir le MultiPoint contenu dans le GeometryBag
                    pPointColl = CType(pGeomColl.Geometry(j), IGeometryCollection)

                    'Vérifier si on doit simplifier
                    If gbSimplifier Then
                        'Simplifier la géométrie
                        pTopoOp = CType(pPointColl, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()
                    End If

                    'Valider la valeur d'attribut selon l'expression régulière
                    oMatch = oRegEx.Match(pPointColl.GeometryCount.ToString)

                    'Vérifier si on doit sélectionner l'élément
                    If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pNewSelectionSet.Add(iOidSel(i))
                        'Vérifier si la géométrie d'intersection est vide
                        If pPointColl.GeometryCount = 0 Then
                            'Ajouter l'enveloppe de la géométrie de l'élément
                            pNewGeomColl.AddGeometry(pGeomSelColl.Geometry(i))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbInt=" & pPointColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                GeometrieToMultiPoint(pGeomSelColl.Geometry(i)), pPointColl.GeometryCount)
                        Else
                            'Ajouter la géométrie d'intersection
                            pNewGeomColl.AddGeometry(CType(pPointColl, IGeometry))
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbInt=" & pPointColl.GeometryCount.ToString & " /ExpReg=" & gsExpression, _
                                                GeometrieToMultiPoint(CType(pPointColl, IGeometry)), pPointColl.GeometryCount)
                        End If
                    End If
                Next

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Remplacer les géométries résultantes
            pGeomResColl = pNewGeomColl

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oMatch = Nothing
            pGeometryBag = Nothing
            pGeomColl = Nothing
            pNewGeomColl = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire l'intersection de type point entre la géométrie de l'élément à traiter et la géométrie d'un élément en relation.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface contenant la géométrie de l'élément à traiter.</param>
    '''<param name="pGeometryRel">Interface contenant la géométrie d'un élément en relation.</param>
    ''' 
    '''<returns>"IMultipoint" contenant les intersections de type point entre la géométrie de l'élément traité et les géométries des éléments en relation.</returns>
    '''
    Private Function ExtraireIntersectionPoint(ByRef pGeometry As IGeometry, ByRef pGeometryRel As IGeometry) As IMultipoint
        'Déclarer les variables de travail
        Dim pGeometryInt As IGeometry = Nothing                 'Interface contenant la géométrie d'intersection.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour ajouter l'intersection des géométries.
        Dim pTopoOp As ITopologicalOperator = Nothing           'Interface qui permet d'extraire l'intersection entre deux géométries.

        Try
            'Définir la valeur de retour par défaut
            ExtraireIntersectionPoint = New Multipoint
            ExtraireIntersectionPoint.SpatialReference = pGeometry.SpatialReference
            ExtraireIntersectionPoint.SnapToSpatialReference()
            pGeomColl = CType(ExtraireIntersectionPoint, IGeometryCollection)

            'Interface pour extraire l'intersection entre les deux géométries
            pTopoOp = CType(pGeometry, ITopologicalOperator)

            'Projeter la géométrie de l'élément en relation
            pGeometryRel.Project(pGeometry.SpatialReference)

            'Trouver l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
            pGeometryInt = pTopoOp.Intersect(pGeometryRel, esriGeometryDimension.esriGeometry0Dimension)

            'Vérifier si la géométrie est un point
            If pGeometryInt.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Ajouter l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
                pGeomColl.AddGeometry(pGeometryInt)
            Else
                'Ajouter l'intersection entre la géométrie de l'élément traité et celle de l'élément en relation
                pGeomColl.AddGeometryCollection(CType(pGeometryInt, IGeometryCollection))
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeometryInt = Nothing
            pGeomColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont le nombre d'intersection avec d'autres éléments respecte ou non 
    ''' l'expression spécifiée.
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
    '''<return>Les intersections entre la géométrie des éléments sélectionnés et les géométries des éléments en relation.</return>
    ''' 
    Private Function TraiterNombreIntersectionElement(ByRef pTrackCancel As ITrackCancel,
                                                      Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim oRegEx = New Regex(gsExpression)                'Objet utilisé pour vérifier si la valeur respecte l'expression régulière.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pNewSelectionSet As ISelectionSet = Nothing     'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(1) As Integer                           'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(1) As Integer                           'Vecteur des OIds des éléments en relation.
        Dim iNbRes(1) As Integer                            'Vecteur des nombres d'éléments en relation.

        Try
            'Définir la géométrie par défaut
            TraiterNombreIntersectionElement = New GeometryBag
            TraiterNombreIntersectionElement.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterNombreIntersectionElement, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Définir le vecteur des nombres d'éléments en relation
            ReDim Preserve iNbRes(pSelectionSet.Count)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel, gbLimite)

            'Définir une nouvelle sélection Vide
            pNewSelectionSet = pFeatureSel.SelectionSet

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pSelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection And (gbLimite = gbLimiteRel) Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel, gbLimiteRel)
                End If

                'Afficher le message de recherche des éléments en relation
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Recherche des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message d'identification des éléments trouvés
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des éléments d'intersection (" & pFeatureLayerRel.Name & ") ..."
                'Identifier les éléments trouvés
                IdentifierNombreElement(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, pTrackCancel, iNbRes)
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments trouvés
            SelectionnerNombreElement(bEnleverSelection, oRegEx, iNbRes, iOidSel, pGeomSelColl, pTrackCancel, pGeomResColl, pNewSelectionSet)

            'Définir les éléments sélectionnés
            pFeatureSel.SelectionSet = pNewSelectionSet

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oRegEx = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pNewSelectionSet = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iNbRes = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'identifier les éléments du FeatureLayer qui possèdent des éléments en relation.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="iNbRes"> Vecteur des nombres d'éléments en relation.</param>
    ''' 
    Private Sub IdentifierNombreElement(ByRef pRelResult As IRelationResult, ByRef pGeomSelColl As IGeometryCollection, ByRef pGeomRelColl As IGeometryCollection,
                                        ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, ByRef pTrackCancel As ITrackCancel, _
                                        ByRef iNbRes() As Integer)
        'Déclarer les variables de travail
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Ajouter un élément trouvé
                    iNbRes(iSel) = iNbRes(iSel) + 1
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte ou non le nombre d'éléments spécifiés.
    '''</summary>
    ''' 
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''<param name="oRegEx"> Objet utilisé pour vérifier si la valeur respecte l'expression régulière.</param>
    '''<param name="iNbRes"> Vecteur des nombres d'éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIds des éléments à traiter.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pGeomResColl"> Interface contenant les géométries des éléments trouvés.</param>
    '''<param name="pNewSelectionSet"> Interface pour sélectionner les éléments.</param>
    ''' 
    Private Sub SelectionnerNombreElement(ByVal bEnleverSelection As Boolean, ByRef oRegEx As Regex, ByRef iNbRes() As Integer, ByRef iOidSel() As Integer, _
                                          ByRef pGeomSelColl As IGeometryCollection, ByRef pTrackCancel As ITrackCancel, _
                                          ByRef pGeomResColl As IGeometryCollection, ByRef pNewSelectionSet As ISelectionSet)
        'Déclarer les variables de travail
        Dim oMatch As Match = Nothing       'Objet qui permet d'indiquer si la valeur respecte l'expression régulière.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Traiter tous les éléments
            For i = 0 To pGeomSelColl.GeometryCount - 1
                'Valider la valeur d'attribut selon l'expression régulière
                oMatch = oRegEx.Match(iNbRes(i).ToString)

                'Vérifier si on doit sélectionner l'élément
                If (oMatch.Success And Not bEnleverSelection) Or (Not oMatch.Success And bEnleverSelection) Then
                    'Ajouter l'élément dans la sélection
                    pNewSelectionSet.Add(iOidSel(i))

                    'Ajouter la géométrie trouvée
                    pGeomResColl.AddGeometry(pGeomSelColl.Geometry(i))

                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #" & gsNomAttribut & " /NbInt=" & iNbRes(i).ToString & " /ExpReg=" & gsExpression, _
                                        pGeomSelColl.Geometry(i), iNbRes(i))
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            oMatch = Nothing
        End Try
    End Sub
#End Region
End Class
