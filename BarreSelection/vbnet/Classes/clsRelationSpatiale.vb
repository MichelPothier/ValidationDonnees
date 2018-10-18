Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.ArcMapUI

'**
'Nom de la composante : clsRelationSpatiale.vb
'
'''<summary>
''' Classe qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale avec des éléments en relation respecte ou non 
''' l’information spécifiée spécifiée.
''' 
''' Chaque géométrie des éléments traités est comparée avec ses éléments en relations.
''' 
''' La classe permet de traiter les attributs de relation spatiale EGALE, DISJOINT, CROISE, CHEVAUCHE, TOUCHE, CONTIENT, EST_INCLUS et RELATION.
''' 
''' EGALE : Il n'existe aucune différence entre la géométrie de l'élément à traiter et la géométrie des éléments en relation.
''' DISJOINT : Il n'existe aucune intersection entre la géométrie de l'élément à traiter et la géométrie des éléments en relation.
''' CROISE : La dimension de l'intersection entre deux géométries de même dimension est plus petite.
''' CHEVAUCHE : La dimension de l'intersection entre deux géométries de même dimension est égale.
''' TOUCHE : Il y a une intersection entre les limites de la géométrie de l'élément à traiter et la géométrie des éléments en relation.
''' CONTIENT : La géométrie de l'élément à traiter contient la géométrie des éléments en relation.
''' EST_INCLUS : La géométrie de l'élément à traiter est incluse dans la géométrie des éléments en relation.
''' RELATION : La géométrie de l'élément à traiter respecte la relation à 9 intersections (intérieur, extérieur et limite)
'''            ou la commande de comparaison (SCL: Shape Comparaison Langage) avec la géométrie des éléments en relation.
''' FILTRE : La géométrie de l'élément à traiter respecte la relation à 9 intersections (intérieur, extérieur et limite)
'''          ou la commande de comparaison (SCL: Shape Comparaison Langage) avec la géométrie des éléments en relation.
'''          La vérification est effectuée via un filtre spatial.
''' 
''' Note : La géométrie des éléments en relation est considérée comme une seule géométrie.
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 23 avril 2015
'''</remarks>
''' 
Public Class clsRelationSpatiale
    Inherits clsValeurAttribut

    'Déclarer les variables globales
    '''<summary>Commande de comparaison selon le langage SCL.</summary>
    Protected gsComparaison As String = ""

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "TOUCHE"
            Expression = "VRAI"
            Comparaison = ""
            gpFeatureLayersRelation = New Collection

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
    '''<param name="sComparaison"> Commande de comparaison de raffinement à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String, ByVal sExpression As String, Optional ByVal sComparaison As String = "")
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression
            Comparaison = sComparaison
            gpFeatureLayersRelation = New Collection

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
    ''' 
    '''<param name="pMap"> Interface ESRI contenant tous les FeatureLayers.</param>
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression régulière (1) à traiter.</param>
    ''' 
    '''</summary>
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
        gsComparaison = Nothing
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
            Nom = "RelationSpatiale"
        End Get
    End Property

    '''<summary>
    ''' Propriété qui permet de définir ou retourner la commande de comparaison selon le langage SCL ou le masque spatial.
    '''</summary>
    ''' 
    Public Property Comparaison() As String
        Get
            Comparaison = gsComparaison
        End Get
        Set(ByVal value As String)
            gsComparaison = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la ligne de paramètre à traiter.
    '''</summary>
    ''' 
    Public Overloads Overrides Property Parametres() As String
        Get
            'Retourner la valeur des paramètres
            Parametres = gsNomAttribut & " " & gsExpression
            'Vérifier la présence de la commande de comparaison
            If gsComparaison <> "" Then
                'Retourner aussi la commande de comparaison
                Parametres = Parametres & " " & gsComparaison
            End If
        End Get
        Set(ByVal value As String)
            'Déclarer les variables de travail
            Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression régulière

            'Mettre en majuscule les paramètres
            value = value.ToUpper
            'Extraire les paramètres
            params = value.Split(CChar(" "))
            'Vérifier si les deux paramètres sont présents
            If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION [SCL]")

            'Définir les valeurs par défaut
            gsNomAttribut = params(0)
            gsExpression = params(1)

            'Vérifier si aucune commande de comparaison
            If params.Length < 3 Then
                'Aucun commande de comparaison
                gsComparaison = ""

                'Si seulement la masque spatial est spécifié
            ElseIf params.Length = 3 Then
                'Définir le masque spatial sans les virgules
                gsComparaison = params(2).Replace(",", "")

                'Si une commande de comparaison est présente
            Else
                'Définir la première partie de la commande de comparaison
                gsComparaison = params(2)
                'Conserver toutes les parties de la commande de comparaison
                For i = 3 To params.Length - 1
                    'Définir la partie de la commande de comparaison
                    gsComparaison = gsComparaison & " " & params(i)
                Next
            End If
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
                'Définir le paramètre CHEVAUCHE
                ListeParametres.Add("CHEVAUCHE VRAI")
                ListeParametres.Add("CHEVAUCHE FAUX")

                'Définir le paramètre CONTIENT
                ListeParametres.Add("CONTIENT VRAI")
                ListeParametres.Add("CONTIENT FAUX")
                'Définir le paramètre CONTIENT-BORDE
                ListeParametres.Add("CONTIENT-BORDE-S/S VRAI T**,*1*,FF*")
                ListeParametres.Add("CONTIENT-BORDE-S/S FAUX T**,*1*,FF*")
                'Définir le paramètre CONTIENT-STRICT
                ListeParametres.Add("CONTIENT-STRICT VRAI T**,*F*,FF*")
                ListeParametres.Add("CONTIENT-STRICT FAUX T**,*F*,FF*")
                'Définir le paramètre CONTIENT-TANGENT
                ListeParametres.Add("CONTIENT-TANGENT-S/S VRAI T**,*0*,FF*")
                ListeParametres.Add("CONTIENT-TANGENT-S/S FAUX T**,*0*,FF*")
                ListeParametres.Add("CONTIENT-TANGENT-S/L VRAI T**,*0*,FF*")
                ListeParametres.Add("CONTIENT-TANGENT-S/L FAUX T**,*0*,FF*")
                ListeParametres.Add("CONTIENT-TANGENT-POINT-S/L VRAI T**,F0*,FF*")
                ListeParametres.Add("CONTIENT-TANGENT-POINT-S/L FAUX T**,F0*,FF*")

                'Définir le paramètre CROISE
                ListeParametres.Add("CROISE VRAI")
                ListeParametres.Add("CROISE FAUX")
                'Définir le paramètre CROISE-DEMI
                ListeParametres.Add("CROISE-DEMI-L/S VRAI T*T,T**,***")
                ListeParametres.Add("CROISE-DEMI-L/S FAUX T*T,T**,***")
                ListeParametres.Add("CROISE-DEMI-S/L VRAI TTT,***,***")
                ListeParametres.Add("CROISE-DEMI-S/L FAUX TTT,***,***")
                'Définir le paramètreCROISE-COMPLET
                ListeParametres.Add("CROISE-COMPLET-L/S VRAI T*T,F**,***")
                ListeParametres.Add("CROISE-COMPLET-L/S FAUX T*T,F**,***")
                ListeParametres.Add("CROISE-COMPLET-S/L VRAI TFT,***,***")
                ListeParametres.Add("CROISE-COMPLET-S/L FAUX TFT,***,***")

                'Définir le paramètre DISJOINT
                ListeParametres.Add("DISJOINT VRAI")
                ListeParametres.Add("DISJOINT FAUX")

                'Définir le paramètre EGALE
                ListeParametres.Add("EGALE VRAI")
                ListeParametres.Add("EGALE FAUX")

                'Définir le paramètre EST_INCLUS
                ListeParametres.Add("EST_INCLUS VRAI")
                ListeParametres.Add("EST_INCLUS FAUX")
                'Définir le paramètre EST_INCLUS-BORDE
                ListeParametres.Add("EST_INCLUS-BORDE-S/S VRAI T*F,*1F,***")
                ListeParametres.Add("EST_INCLUS-BORDE-S/S FAUX T*F,*1F,***")
                'Définir le paramètre EST_INCLUS-STRICT
                ListeParametres.Add("EST_INCLUS-STRICT VRAI T*F,*FF,***")
                ListeParametres.Add("EST_INCLUS-STRICT FAUX T*F,*FF,***")
                'Définir le paramètre EST_INCLUS-TANGENT
                ListeParametres.Add("EST_INCLUS-TANGENT-S/S VRAI T*F,*0F,***")
                ListeParametres.Add("EST_INCLUS-TANGENT-S/S FAUX T*F,*0F,***")
                ListeParametres.Add("EST_INCLUS-TANGENT-L/S VRAI TFF,*0F,***")
                ListeParametres.Add("EST_INCLUS-TANGENT-L/S FAUX TFF,*0F,***")
                ListeParametres.Add("EST_INCLUS-TANGENT-POINT-L/S VRAI TFF,*0F,***")
                ListeParametres.Add("EST_INCLUS-TANGENT-POINT-L/S FAUX TFF,*0F,***")

                'Définir le paramètre FILTRE
                ListeParametres.Add("FILTRE VRAI TFF,T0F,***")
                ListeParametres.Add("FILTRE FAUX TFF,T0F,***")

                'Définir le paramètre RELATION
                ListeParametres.Add("RELATION VRAI FF*,FT*,***")
                ListeParametres.Add("RELATION FAUX FF*,FT*,***")
                ListeParametres.Add("RELATION VRAI INTERSECT (G1.BOUNDARY, G2.EXTERIOR) = TRUE")
                ListeParametres.Add("RELATION FAUX INTERSECT (G1.BOUNDARY, G2.EXTERIOR) = TRUE")

                'Définir le paramètre TOUCHE
                ListeParametres.Add("TOUCHE VRAI")
                ListeParametres.Add("TOUCHE FAUX")
                'Définir le paramètre TOUCHE-BORDE
                ListeParametres.Add("TOUCHE-BORDE VRAI")
                ListeParametres.Add("TOUCHE-BORDE FAUX")
                ListeParametres.Add("TOUCHE-BORDE-S/S VRAI F**,*1*,***")
                ListeParametres.Add("TOUCHE-BORDE-S/S FAUX F**,*1*,***")
                'ListeParametres.Add("TOUCHE-BORDE-ANNEAU-A/B-S/S VRAI F**,*1F,***")
                'ListeParametres.Add("TOUCHE-BORDE-ANNEAU-A/B-S/S FAUX F**,*1F,***")
                'ListeParametres.Add("TOUCHE-BORDE-ANNEAU-B/A-S/S VRAI F**,*1*,*F*")
                'ListeParametres.Add("TOUCHE-BORDE-ANNEAU-B/A-S/S FAUX F**,*1*,*F*")
                'Définir le paramètre TOUCHE-TANGENT
                ListeParametres.Add("TOUCHE-TANGENT VRAI")
                ListeParametres.Add("TOUCHE-TANGENT FAUX")
                ListeParametres.Add("TOUCHE-TANGENT VRAI F**,*0*,***")
                ListeParametres.Add("TOUCHE-TANGENT FAUX F**,*0*,***")
                ListeParametres.Add("TOUCHE-TANGENT-POINT VRAI FF*,F0*,***")
                ListeParametres.Add("TOUCHE-TANGENT-POINT FAUX FF*,F0*,***")
                ListeParametres.Add("TOUCHE-TANGENT-COMPLET-L/S VRAI F*F,*0F,***")
                ListeParametres.Add("TOUCHE-TANGENT-COMPLET-L/S FAUX F*F,*0F,***")
                ListeParametres.Add("TOUCHE-TANGENT-COMPLET-S/L VRAI F**,*0*,FF*")
                ListeParametres.Add("TOUCHE-TANGENT-COMPLET-S/L FAUX F**,*0*,FF*")
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
                'Vérifier si la FeatureClass est de type Polyline et Polygon
                If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    'La contrainte est valide
                    FeatureClassValide = True
                    gsMessage = "La contrainte est valide"
                Else
                    gsMessage = "ERREUR : Le type de la FeatureClass n'est pas de type Point, MultiPoint, Polyline ou Polygon."
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
        'Déclarer les variables de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant un FeatureLayer en relation

        Try
            'La contrainte est invalide par défaut.
            AttributValide = False
            gsMessage = "ERREUR : La relation spatiale est invalide : " & gsNomAttribut

            'Vérifier si l'attribut est valide
            If gsNomAttribut.Contains("EGALE") Then
                'Traiter tous les FeatureLayer en relation
                For Each pFeatureLayer In gpFeatureLayersRelation
                    'Vérifier si les types de géométrie sont différents
                    If gpFeatureLayerSelection.FeatureClass.ShapeType <> pFeatureLayer.FeatureClass.ShapeType Then
                        'Si la dimension est différente
                        If ((gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                           And pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) _
                        Or (gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                           And pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint)) = False Then
                            'La contrainte est invalide.
                            AttributValide = False
                            gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                            Exit Function
                        End If
                    End If
                Next
                'La contrainte est valide.
                AttributValide = True
                gsMessage = "La contrainte est valide"

                'Vérifier si l'attribut est valide
            ElseIf gsNomAttribut.Contains("CHEVAUCHE") Then
                'Traiter tous les FeatureLayer en relation
                For Each pFeatureLayer In gpFeatureLayersRelation
                    'Vérifier si les types de géométrie sont différents
                    If gpFeatureLayerSelection.FeatureClass.ShapeType <> pFeatureLayer.FeatureClass.ShapeType Then
                        'Si la dimension est différente
                        If ((gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                           And pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) _
                        Or (gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                           And pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint)) = False Then
                            'La contrainte est invalide.
                            AttributValide = False
                            gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                            Exit Function
                        End If
                        'Si la géométrie est de type point
                    ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                        'La contrainte est invalide.
                        AttributValide = False
                        gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                        Exit Function
                    End If
                Next
                'La contrainte est valide.
                AttributValide = True
                gsMessage = "La contrainte est valide"

            ElseIf gsNomAttribut.Contains("CROISE") Then
                'Traiter tous les FeatureLayer en relation
                For Each pFeatureLayer In gpFeatureLayersRelation
                    'Vérifier si le type de géométrie est polyline
                    If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        'Si le type de géométrie n'est pas polyline ou polygon
                        If Not (pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                             Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon) Then
                            'La contrainte est invalide.
                            AttributValide = False
                            gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                            Exit Function
                        End If
                        'Si le type de géométrie est polygon
                    ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'Si le type de géométrie n'est pas polyline
                        If pFeatureLayer.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPolyline Then
                            'La contrainte est invalide.
                            AttributValide = False
                            gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                            Exit Function
                        End If
                    Else
                        'La contrainte est invalide.
                        AttributValide = False
                        gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                        Exit Function
                    End If
                Next
                'La contrainte est valide.
                AttributValide = True
                gsMessage = "La contrainte est valide"

            ElseIf gsNomAttribut.Contains("TOUCHE") Then
                'Traiter tous les FeatureLayer en relation
                For Each pFeatureLayer In gpFeatureLayersRelation
                    'Si la dimension est 0
                    If (gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                     Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) _
                    And (pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                     Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) Then
                        'La contrainte est invalide.
                        AttributValide = False
                        gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                        Exit Function
                    End If
                Next
                'La contrainte est valide.
                AttributValide = True
                gsMessage = "La contrainte est valide"

            ElseIf gsNomAttribut.Contains("CONTIENT") Then
                'Traiter tous les FeatureLayer en relation
                For Each pFeatureLayer In gpFeatureLayersRelation
                    'Vérifier si le type de géométrie est point ou un multipoint comparer à une polyline ou un polygon
                    If (gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                     Or gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) _
                    And (pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                     Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon) Then
                        'La contrainte est invalide.
                        AttributValide = False
                        gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                        Exit Function
                        'Vérifier si le type de géométrie est polyline
                    ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                       And pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        'La contrainte est invalide.
                        AttributValide = False
                        gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                        Exit Function
                    End If
                Next
                'La contrainte est valide.
                AttributValide = True
                gsMessage = "La contrainte est valide"

            ElseIf gsNomAttribut.Contains("EST_INCLUS") Then
                'Traiter tous les FeatureLayer en relation
                For Each pFeatureLayer In gpFeatureLayersRelation
                    'Vérifier si le type de géométrie est point ou un multipoint comparer à une polyline ou un polygon
                    If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon _
                    And (pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                     Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint _
                     Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline) Then
                        'La contrainte est invalide.
                        AttributValide = False
                        gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                        Exit Function
                        'Vérifier si le type de géométrie est polyline
                    ElseIf gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                    And (pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                     Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint) Then
                        'La contrainte est invalide.
                        AttributValide = False
                        gsMessage = "ERREUR : La relation spatiale est invalide pour ces types de géométries : " & gsNomAttribut
                        Exit Function
                    End If
                Next
                'La contrainte est valide.
                AttributValide = True
                gsMessage = "La contrainte est valide"

            ElseIf gsNomAttribut.Contains("DISJOINT") Or gsNomAttribut.Contains("RELATION") Or gsNomAttribut.Contains("FILTRE") Then
                'La contrainte est valide.
                AttributValide = True
                gsMessage = "La contrainte est valide"
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
    ''' Routine qui permet d'indiquer si l'expression régulière est valide.
    '''</summary>
    ''' 
    '''<return>Boolean qui indique si l'expression régulière est valide.</return>
    '''
    Public Overloads Overrides Function ExpressionValide() As Boolean
        Try
            'Par défaut l'expression est invalide
            ExpressionValide = False
            gsMessage = "ERREUR : L'expression est invalide : " & gsExpression

            'Vérifier si l'expression régulière est présente
            If Len(gsExpression) > 0 Then
                'Vérifier si l'expression est VRAI ou FAUX
                If gsExpression = "VRAI" Or gsExpression = "FAUX" Then
                    'Vérifier si le masque est valide
                    If gsNomAttribut = "RELATION" Or gsNomAttribut = "SPATIAL" Then
                        'Vérifier si la longueur est 9
                        If gsComparaison.Length >= 9 Then
                            'La contrainte est valide
                            ExpressionValide = True
                            gsMessage = "La contrainte est valide"
                        Else
                            gsMessage = "ERREUR : La longueur de la commande de comparaison n'est pas d'au moins 9 : " & gsComparaison
                        End If

                        'Vérifier si uen commande de comparaison est présente
                    ElseIf gsComparaison.Length > 0 Then
                        'Vérifier si la longueur est 9
                        If gsComparaison.Length >= 9 Then
                            'La contrainte est valide
                            ExpressionValide = True
                            gsMessage = "La contrainte est valide"
                        Else
                            gsMessage = "ERREUR : La longueur de la commande de comparaison n'est pas d'au moins 9 : " & gsComparaison
                        End If

                    Else
                        'La contrainte est valide
                        ExpressionValide = True
                        gsMessage = "La contrainte est valide"
                    End If
                Else
                    gsMessage = "ERREUR : L'expression ne correspond pas à VRAI ou FAUX : " & gsExpression
                End If
            Else
                gsMessage = "ERREUR : L'expression est absente."
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale entre deux géométries
    ''' respecte ou non celle spécifiée.
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

            'Créer la classe d'erreurs au besoin
            CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

            'Inverser la sélection si la relation n'est pas vrai
            If gsExpression <> "VRAI" Then bEnleverSelection = Not bEnleverSelection

            'Vérifier si l'attribut est EGALE
            If gsNomAttribut.Contains("EGALE") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurEgale(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est DISJOINT
            ElseIf gsNomAttribut.Contains("DISJOINT") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurDisjoint(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est CROISE
            ElseIf gsNomAttribut.Contains("CROISE") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurCroise(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est CHEVAUCHE
            ElseIf gsNomAttribut.Contains("CHEVAUCHE") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurChevauche(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est TOUCHE
            ElseIf gsNomAttribut.Contains("TOUCHE") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurTouche(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est CONTIENT
            ElseIf gsNomAttribut.Contains("CONTIENT") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurContient(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est EST_INCLUS
            ElseIf gsNomAttribut.Contains("EST_INCLUS") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurEstInclus(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est RELATION
            ElseIf gsNomAttribut.Contains("RELATION") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurRelation(pTrackCancel, bEnleverSelection)

                'Vérifier si l'attribut est FILTRE
            ElseIf gsNomAttribut.Contains("FILTRE") Then
                'Traiter le FeatureLayer selon l'opérateur spatial
                Selectionner = TraiterOperateurFiltre(pTrackCancel, bEnleverSelection)
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
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur EGALE VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurEgale(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurEgale = New GeometryBag
            TraiterOperateurEgale.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurEgale, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            pSelectionSet = pFeatureSel.SelectionSet

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale ÉGALE (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Within(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Extraire les Oids d'éléments trouvés
                ExtraireListeOidEgale(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd, pTrackCancel)
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur DISJOINT VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurDisjoint(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurDisjoint = New GeometryBag
            TraiterOperateurDisjoint.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurDisjoint, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface utilisé pour sélectionner les éléments trouvés
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale DISJOINT (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Extraire les Oids d'éléments trouvés
                ExtraireListeOid(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd)
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, Not bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur FILTRE VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurFiltre(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pSpatialIndex As ISpatialIndex = Nothing        'Interface utilisée pour définir un index spatial dans un GeometryBag.
        Dim pSpatialFilter As ISpatialFilter = Nothing      'Interface contenant le requête spatiale.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour lire les éléments.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour lire les éléments.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément lu.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie de l'élément lu.
        Dim iOidRel(1) As Integer                           'Vecteur des OIds des éléments en relation.
        Dim sRelDescription As String = ""

        Try
            'Définir la géométrie par défaut
            TraiterOperateurFiltre = New GeometryBag
            TraiterOperateurFiltre.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurFiltre, IGeometryCollection)

            'Inverser la relation
            sRelDescription = gsComparaison.Substring(0, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(3, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(6, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(1, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(4, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(7, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(2, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(5, 1)
            sRelDescription = sRelDescription & gsComparaison.Substring(8, 1)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Définir la requête spatiale
            pSpatialFilter = New SpatialFilter
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation
            pSpatialFilter.SpatialRelDescription = sRelDescription
            pSpatialFilter.OutputSpatialReference(gpFeatureLayerSelection.FeatureClass.ShapeFieldName) = TraiterOperateurFiltre.SpatialReference
            pSpatialFilter.GeometryField = gpFeatureLayerSelection.FeatureClass.ShapeFieldName

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Afficher le message de lecture des éléments
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                'Lire les éléments en relation
                LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)

                'Indexer le Bag des éléments en relation
                pSpatialIndex = CType(pGeomRelColl, ISpatialIndex)
                pSpatialIndex.AllowIndexing = True
                pSpatialIndex.Invalidate()

                'Définir la géométrie de la requete spatiale
                pSpatialFilter.Geometry = CType(pGeomRelColl, IGeometry)

                'Afficher le message de sélection des éléments trouvés
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments ..."
                'Vérifier si on doit enlever les éléments trouvés
                If (bEnleverSelection And gsExpression = "VRAI") Or (bEnleverSelection = False And gsExpression = "FAUX") Then
                    'Enlever les éléments trouvés de la sélection
                    pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultSubtract, False)
                Else
                    'Conserver les éléments trouvés de la sélection
                    pFeatureSel.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultAnd, False)
                End If
            Next

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Écriture des éléments en erreur (" & pFeatureLayerRel.Name & ") ..."
            'Interfaces pour extraire les éléments sélectionnés
            pFeatureSel.SelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier élément
            pFeature = pFeatureCursor.NextFeature()

            'Traiter tous les éléments du FeatureLayer
            Do Until pFeature Is Nothing
                'Définir la géométrie à traiter
                pGeometry = pFeature.ShapeCopy

                'Projeter la géométrie à traiter
                pGeometry.Project(gpFeatureLayerSelection.AreaOfInterest.SpatialReference)

                'Ajouter la géométrie en erreur
                pGeomResColl.AddGeometry(pGeometry)

                'Écrire une erreur
                EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #NbOidRel=1 /" & Parametres, pGeometry, 1)

                'Extraire le prochain élément à traiter
                pFeature = pFeatureCursor.NextFeature()
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pFeatureLayerRel = Nothing
            pSpatialIndex = Nothing
            pSpatialFilter = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            iOidRel = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur CROISE VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurCroise(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurCroise = New GeometryBag
            TraiterOperateurCroise.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurCroise, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale CROISE (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Crosses(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si une commande de comparaison de raffinement est présente
                If gsComparaison.Length > 0 Then
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOidRelation(gsComparaison, pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd, pTrackCancel)
                Else
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOid(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur CHEVAUCHE VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurChevauche(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurChevauche = New GeometryBag
            TraiterOperateurChevauche.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurChevauche, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale CHEVAUCHE (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Overlaps(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si une commande de comparaison de raffinement est présente
                If gsComparaison.Length > 0 Then
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOidRelation(gsComparaison, pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd, pTrackCancel)
                Else
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOid(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur TOUCHE VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurTouche(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurTouche = New GeometryBag
            TraiterOperateurTouche.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurTouche, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale " + gsNomAttribut + " (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Vérifier si la relation Touche et Borde
                If gsNomAttribut.Contains("BORDE") And gsComparaison.Length = 0 Then
                    'Exécuter la recherche et retourner le résultat de la relation spatiale
                    pRelResult = pRelOpNxM.RelationEx(CType(pGeomRelColl, IGeometryBag), esriSpatialRelationEnum.esriSpatialRelationLineTouch)
                    'Vérifier si la relation touche et Tangent
                ElseIf gsNomAttribut.Contains("TANGENT") And gsComparaison.Length = 0 Then
                    'Exécuter la recherche et retourner le résultat de la relation spatiale
                    pRelResult = pRelOpNxM.RelationEx(CType(pGeomRelColl, IGeometryBag), esriSpatialRelationEnum.esriSpatialRelationPointTouch)
                    'Vérifier si la relation touche seulement
                Else
                    'Exécuter la recherche et retourner le résultat de la relation spatiale
                    pRelResult = pRelOpNxM.Touches(CType(pGeomRelColl, IGeometryBag))
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si une commande de comparaison de raffinement est présente
                If gsComparaison.Length > 0 Then
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOidRelation(gsComparaison, pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd, pTrackCancel)
                Else
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOid(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur CONTIENT VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurContient(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurContient = New GeometryBag
            TraiterOperateurContient.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurContient, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale CONTIENT (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Contains(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si une commande de comparaison de raffinement est présente
                If gsComparaison.Length > 0 Then
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOidRelation(gsComparaison, pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd, pTrackCancel)
                Else
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOid(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale respecte ou non l'opérateur EST_INCLUS VRAI
    ''' avec ses éléments en relation.
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
    Private Function TraiterOperateurEstInclus(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurEstInclus = New GeometryBag
            TraiterOperateurEstInclus.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurEstInclus, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale EST_INCLUS (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Within(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Vérifier si une commande de comparaison de raffinement est présente
                If gsComparaison.Length > 0 Then
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOidRelation(gsComparaison, pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd, pTrackCancel)
                Else
                    'Extraire les Oids d'éléments trouvés
                    ExtraireListeOid(pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd)
                End If
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont la relation spatiale VRAI respecte ou non le masque à 9 intersections
    ''' ou une commande de comparaison avec ses éléments en relation.
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
    Private Function TraiterOperateurRelation(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing        'Interface pour sélectionner les éléments.
        Dim pGeomResColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries résultantes trouvées.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des l'élément à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pFeatureLayerRel As IFeatureLayer = Nothing     'Interface contenant un FeatureLayer en relation.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim bMemeClasse As Boolean = False                  'Indique si la classe à traiter est la même que celle en relation.
        Dim bMemeDefinition As Boolean = False              'Indique si la définition du Layer à traiter est la même que celle en relation.
        Dim bMemeSelection As Boolean = False               'Indique si la sélection du Layer à traiter est la même que celle en relation.
        Dim iOidSel(0) As Integer   'Vecteur des OIds des éléments à traiter.
        Dim iOidRel(0) As Integer   'Vecteur des OIds des éléments en relation.
        Dim iOidAdd(0) As Integer   'Vecteur du nombre de OIDs trouvés.

        Try
            'Définir la géométrie par défaut
            TraiterOperateurRelation = New GeometryBag
            TraiterOperateurRelation.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface contenant les géométries des éléments sélectionnés
            pGeomResColl = CType(TraiterOperateurRelation, IGeometryCollection)

            'Conserver la sélection de départ
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments (" & gpFeatureLayerSelection.Name & ") ..."
            'Lire les éléments à traiter 
            LireGeometrie(gpFeatureLayerSelection, pTrackCancel, pGeomSelColl, iOidSel)

            'Initialiser la liste des Oids trouvés
            ReDim Preserve iOidAdd(pGeomSelColl.GeometryCount)

            'Interface pour sélectionner les éléments
            pSelectionSet = pFeatureSel.SelectionSet

            'Traiter tous les featureLayers en relation
            For Each pFeatureLayerRel In gpFeatureLayersRelation
                'Vérifier si le Layer à traiter possède la même classe et la même définition que celui en relation
                Call MemeClasseMemeDefinition(gpFeatureLayerSelection, pFeatureLayerRel, bMemeClasse, bMemeDefinition, bMemeSelection)

                'Vérifier si la classe, la définition et la sélection est la même entre le Layer à traiter et celui en relation
                If bMemeClasse And bMemeDefinition And bMemeSelection Then
                    'Définir les géométries en relation
                    pGeomRelColl = pGeomSelColl
                    'Définir les Oids en relation
                    iOidRel = iOidSel
                Else
                    'Afficher le message de lecture des éléments
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Lecture des éléments en relation (" & pFeatureLayerRel.Name & ") ..."
                    'Lire les éléments en relation
                    LireGeometrie(pFeatureLayerRel, pTrackCancel, pGeomRelColl, iOidRel)
                End If

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traitement de la relation spatiale RELATION (" & gpFeatureLayerSelection.Name & "/" & pFeatureLayerRel.Name & ") ..."
                'Interface pour traiter la relation spatiale
                pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)
                'Exécuter la recherche et retourner le résultat de la relation spatiale
                pRelResult = pRelOpNxM.Intersects(CType(pGeomRelColl, IGeometryBag))

                'Afficher le message de traitement de la relation spatiale
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Extraction des Oids trouvés (" & gpFeatureLayerSelection.Name & ") ..."
                'Extraire les Oids d'éléments trouvés
                ExtraireListeOidRelation(gsComparaison, pRelResult, pGeomSelColl, pGeomRelColl, iOidSel, iOidRel, bMemeClasse, iOidAdd, pTrackCancel)
            Next

            'Afficher le message de sélection des éléments trouvés
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments trouvés (" & gpFeatureLayerSelection.Name & ") ..."
            'Sélectionner les éléments en erreur
            SelectionnerElementErreur(pGeomSelColl, iOidSel, iOidAdd, bEnleverSelection, pTrackCancel, pSelectionSet, pGeomResColl)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGeomResColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureLayerRel = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            iOidSel = Nothing
            iOidRel = Nothing
            iOidAdd = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet d'extraire la liste des Oids d'éléments du FeatureLayer traité qui respecte la relation du masque spatial à 9 intersections.
    '''</summary>
    ''' 
    '''<param name="sComparaison"> Contient la commande de comparaison de raffinement ou le masque spatial à 9 intersections.</param>
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    Private Sub ExtraireListeOidRelation(ByVal sComparaison As String, ByVal pRelResult As IRelationResult, ByVal pGeomSelColl As IGeometryCollection, _
                                         ByVal pGeomRelColl As IGeometryCollection, ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, _
                                         ByRef iOidAdd() As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pRelOp As IRelationalOperator2 = Nothing        'Interface pour vérifier la relation spatiale.
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.
        Dim sSCL As String = ""             'Contient la commande de comparaison selon le langage SCL.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pRelResult.RelationElementCount, pTrackCancel)

            'Vérifier si la commande de comparaison contient seulement le masque à 9 intersections
            If sComparaison.Length = 9 Then
                'Définir la commande de comparaison selon le langage SCL
                sSCL = "RELATE (G1, G2, '" & sComparaison & "')"

                'Si la commande de comparaison contient vraiment le SCL
            Else
                'Définir la commande de comparaison selon le langage SCL
                sSCL = sComparaison
            End If

            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Interface pour vérifier la relation spatiale
                    pRelOp = CType(pGeomSelColl.Geometry(iSel), IRelationalOperator2)

                    'Vérifier la relation spatiale
                    If pRelOp.Relation(pGeomRelColl.Geometry(iRel), sSCL) Then
                        'Indiquer que le OID respecte la relation
                        iOidAdd(iSel) = iOidAdd(iSel) + 1
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
            pRelOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire la liste des Oids d'éléments du FeatureLayer traité qui respecte la relation EGALE.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    Private Sub ExtraireListeOidEgale(ByVal pRelResult As IRelationResult, ByVal pGeomSelColl As IGeometryCollection, ByVal pGeomRelColl As IGeometryCollection, _
                                      ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, _
                                      ByRef iOidAdd() As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pRelOp As IRelationalOperator2 = Nothing        'Interface pour vérifier la relation spatiale.
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
                    'Interface pour vérifier la relation spatiale
                    pRelOp = CType(pGeomSelColl.Geometry(iSel), IRelationalOperator2)

                    'Vérifier la relation spatiale
                    If pRelOp.Equals(pGeomRelColl.Geometry(iRel)) Then
                        'Indiquer que le OID respecte la relation
                        iOidAdd(iSel) = iOidAdd(iSel) + 1
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
            pRelOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire la liste des Oids d'éléments du FeatureLayer traité qui respecte la relation du masque spatiale TOUCHE pour les surfaces.
    '''</summary>
    ''' 
    '''<param name="sMasqueSpatial"> Contient le masque spatial à 9 intersections.</param>
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''
    Private Sub ExtraireListeOidSurfaceTouche(ByVal sMasqueSpatial As String, ByVal pRelResult As IRelationResult, ByVal pGeomSelColl As IGeometryCollection, _
                                              ByVal pGeomRelColl As IGeometryCollection, ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, _
                                              ByRef iOidAdd() As Integer)
        'Déclarer les variables de travail
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Vérifier si le masque à 9 intersections est invalide
            If sMasqueSpatial.Length <> 9 Then
                'Retourner une erreur
                Err.Raise(1, , "Le masque spatial est invalide")
            End If

            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Vérifier si le résultat de l'intersection du masque correspondent (IL,LI,LL,EL,LE,IE)
                    If IntersectionLimiteInterieur(sMasqueSpatial.Substring(1, 1), pGeomRelColl.Geometry(iRel), pGeomSelColl.Geometry(iSel)) _
                    And IntersectionLimiteInterieur(sMasqueSpatial.Substring(3, 1), pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel)) _
                    And IntersectionLimiteLimite(sMasqueSpatial.Substring(4, 1), pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel)) _
                    And IntersectionLimiteExterieur(sMasqueSpatial.Substring(5, 1), pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel)) _
                    And IntersectionLimiteExterieur(sMasqueSpatial.Substring(7, 1), pGeomRelColl.Geometry(iRel), pGeomSelColl.Geometry(iSel)) _
                    And IntersectionInterieurExterieur(sMasqueSpatial.Substring(6, 1), pGeomRelColl.Geometry(iRel), pGeomSelColl.Geometry(iSel)) Then
                        'Indiquer que le OID respecte la relation
                        iOidAdd(iSel) = iOidAdd(iSel) + 1
                    End If
                End If

                'Vider la mémoire
                GC.Collect()
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire la liste des Oids d'éléments du FeatureLayer traité qui respecte la relation du masque spatiale CONTIENT pour les surfaces.
    '''</summary>
    ''' 
    '''<param name="sMasqueSpatial"> Contient le masque spatial à 9 intersections.</param>
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''
    Private Sub ExtraireListeOidSurfaceContient(ByVal sMasqueSpatial As String, ByVal pRelResult As IRelationResult, ByVal pGeomSelColl As IGeometryCollection, _
                                                ByVal pGeomRelColl As IGeometryCollection, ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, _
                                                ByRef iOidAdd() As Integer)
        'Déclarer les variables de travail
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Vérifier si le masque à 9 intersections est invalide
            If sMasqueSpatial.Length <> 9 Then
                'Retourner une erreur
                Err.Raise(1, , "Le masque spatial est invalide")
            End If

            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Vérifier si le résultat de l'intersection du masque correspondent (IL,LI,LL)
                    If IntersectionLimiteInterieur(sMasqueSpatial.Substring(1, 1), pGeomRelColl.Geometry(iRel), pGeomSelColl.Geometry(iSel)) _
                    And IntersectionLimiteInterieur(sMasqueSpatial.Substring(3, 1), pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel)) _
                    And IntersectionLimiteLimite(sMasqueSpatial.Substring(4, 1), pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel)) Then
                        'Indiquer que le OID respecte la relation
                        iOidAdd(iSel) = iOidAdd(iSel) + 1
                    End If
                End If

                'Vider la mémoire
                GC.Collect()
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire la liste des Oids d'éléments du FeatureLayer traité qui respecte la relation du masque spatiale EST_INCLUS pour les surfaces.
    '''</summary>
    ''' 
    '''<param name="sMasqueSpatial"> Contient le masque spatial à 9 intersections.</param>
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''
    Private Sub ExtraireListeOidSurfaceEstInclus(ByVal sMasqueSpatial As String, ByVal pRelResult As IRelationResult, ByVal pGeomSelColl As IGeometryCollection, _
                                                 ByVal pGeomRelColl As IGeometryCollection, ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, _
                                                 ByRef iOidAdd() As Integer)
        'Déclarer les variables de travail
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Vérifier si le masque à 9 intersections est invalide
            If sMasqueSpatial.Length <> 9 Then
                'Retourner une erreur
                Err.Raise(1, , "Le masque spatial est invalide")
            End If

            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Vérifier si le résultat de l'intersection du masque correspondent (IL,LI,LL)
                    If IntersectionLimiteInterieur(sMasqueSpatial.Substring(1, 1), pGeomRelColl.Geometry(iRel), pGeomSelColl.Geometry(iSel)) _
                    And IntersectionLimiteInterieur(sMasqueSpatial.Substring(3, 1), pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel)) _
                    And IntersectionLimiteLimite(sMasqueSpatial.Substring(4, 1), pGeomSelColl.Geometry(iSel), pGeomRelColl.Geometry(iRel)) Then
                        'Indiquer que le OID respecte la relation
                        iOidAdd(iSel) = iOidAdd(iSel) + 1
                    End If
                End If

                'Vider la mémoire
                GC.Collect()
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet  d'extraire la liste des Oids d'éléments du FeatureLayer traité qui respecte la relation spatiale.
    '''</summary>
    ''' 
    '''<param name="pRelResult"> Résultat du traitement de la relation spatiale obtenu.</param>
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="pGeomRelColl"> Interface contenant les géométries des éléments en relation.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidRel"> Vecteur des OIDs d'éléments en relation.</param>
    '''<param name="bMemeClasse"> Indique si on traite la même classe.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    Private Sub ExtraireListeOid(ByVal pRelResult As IRelationResult, ByVal pGeomSelColl As IGeometryCollection, ByVal pGeomRelColl As IGeometryCollection, _
                                 ByVal iOidSel() As Integer, ByVal iOidRel() As Integer, ByVal bMemeClasse As Boolean, _
                                 ByRef iOidAdd() As Integer)
        'Déclarer les variables de travail
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        Try
            'Traiter tous les éléments
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si on ne traite pas la même géométrie
                If Not (iOidSel(iSel) = iOidRel(iRel) And bMemeClasse) Then
                    'Indiquer que le OID respecte la relation
                    iOidAdd(iSel) = iOidAdd(iSel) + 1
                End If
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer traité qui respecte la relation spatiale VRAI du masque à 9 intersections ou du SCL.
    '''</summary>
    ''' 
    '''<param name="pGeomSelColl"> Interface contenant les géométries des éléments à traiter.</param>
    '''<param name="iOidSel"> Vecteur des OIDs d'éléments à traiter.</param>
    '''<param name="iOidAdd"> Vecteur du nombre de OIds trouvés.</param>
    '''<param name="bEnleverSelection">Indique si on veut enlever les éléments trouvés de la sélection.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="pSelectionSet">Interface contenant les éléments sélectionnés.</param>
    '''<param name="pGeomResColl">GéométryBag contenant les géométries en erreur.</param>
    '''
    Private Sub SelectionnerElementErreur(ByVal pGeomSelColl As IGeometryCollection, ByVal iOidSel() As Integer, ByVal iOidAdd() As Integer, _
                                          ByVal bEnleverSelection As Boolean, ByRef pTrackCancel As ITrackCancel, ByRef pSelectionSet As ISelectionSet, _
                                          ByRef pGeomResColl As IGeometryCollection)
        'Déclarer les variables de travail
        Dim pGDBridge As IGeoDatabaseBridge2 = Nothing      'Interface pour ajouter une liste de OID dans le SelectionSet.
        Dim iOid(0) As Integer      'Liste des OIDs trouvés.
        Dim iNbOid As Integer = 0   'Compteur de OIDs trouvés.

        Try
            'Afficher la barre de progression
            InitBarreProgression(0, pGeomSelColl.GeometryCount, pTrackCancel)

            'Vérifier si on veut conserver ceux qui ne respectent pas la relation
            If bEnleverSelection Then
                'Traiter tous les éléments
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Vérifier si on doit ajouter le OID dans la sélection
                    If iOidAdd(i) = 0 Then
                        'Définir la liste des OIDs à sélectionner
                        iNbOid = iNbOid + 1

                        'Redimensionner la liste des OIDs à sélectionner
                        ReDim Preserve iOid(iNbOid - 1)

                        'Définir le OIDs dans la liste
                        iOid(iNbOid - 1) = iOidSel(i)
                        'Ajouter la géométrie en erreur dans le Bag
                        pGeomResColl.AddGeometry(pGeomSelColl.Geometry(i))

                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #NbOidRel=" & iOidAdd(i).ToString & " /" & Parametres, pGeomSelColl.Geometry(i), iOidAdd(i))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next

                'Si on veut conserver ceux qui respectent la relation
            Else

                'Traiter tous les éléments
                For i = 0 To pGeomSelColl.GeometryCount - 1
                    'Vérifier si on doit ajouter le OID dans la sélection
                    If iOidAdd(i) > 0 Then
                        'Définir la liste des OIDs à sélectionner
                        iNbOid = iNbOid + 1

                        'Redimensionner la liste des OIDs à sélectionner
                        ReDim Preserve iOid(iNbOid - 1)

                        'Définir le OIDs dans la liste
                        iOid(iNbOid - 1) = iOidSel(i)

                        'Ajouter la géométrie en erreur dans le Bag
                        pGeomResColl.AddGeometry(pGeomSelColl.Geometry(i))

                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & iOidSel(i).ToString & " #NbOidRel=" & iOidAdd(i).ToString & " /" & Parametres, pGeomSelColl.Geometry(i), iOidAdd(i))
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Throw New CancelException("Traitement annulé !")
                Next
            End If

            'Interface pour enlever ou ajouter les OIDs trouvés dans le SelectionSet
            pGDBridge = New GeoDatabaseHelperClass()

            'Ajouter les OIDs trouvés dans le SelectionSet
            pGDBridge.AddList(pSelectionSet, iOid)

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pGDBridge = Nothing
            iOid = Nothing
        End Try
    End Sub
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Fonction qui permet d'indiquer si l'indice d'intersection entre les limites de deux géométrie A et B est respecté.
    '''</summary>
    ''' 
    '''<param name="sIndice"> Indice d'intersection du masque spatial (*,T,F,0,1).</param>
    '''<param name="pGeometryA"> Interface contenant la géométrie à traiter.</param>
    '''<param name="pGeometryB"> Interface contenant la géométrie en relation.</param>
    '''
    Private Function IntersectionLimiteLimite(ByVal sIndice As String, ByVal pGeometryA As IGeometry, ByVal pGeometryB As IGeometry) As Boolean
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface pour extraire la limite et traiter l'intersection.
        Dim pRelOp As IRelationalOperator2 = Nothing    'Interface pour vérifier la relation spatiale.
        Dim pLimiteA As IGeometry = Nothing             'Interface contenant la limite de la géométrie A.
        Dim pLimiteB As IGeometry = Nothing             'Interface contenant la limite de la géométrie B.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant le résultat de l'intersection des limites.

        'L'intersection n'est pas respecté par défaut
        IntersectionLimiteLimite = False

        Try
            'Si l'indice est * alors retourner l'indice d'intersection est respecté
            If sIndice = "*" Then Return True

            'Interface pour extraire la limite
            pTopoOp = CType(pGeometryA, ITopologicalOperator2)
            'Interface contenant le limite de A
            pLimiteA = pTopoOp.Boundary

            'Interface pour extraire la limite
            pTopoOp = CType(pGeometryB, ITopologicalOperator2)
            'Interface contenant le limite de B
            pLimiteB = pTopoOp.Boundary

            'Interface pour vérifier la relation spatiale
            pRelOp = CType(pLimiteA, IRelationalOperator2)

            'Vérifier si les limites sont disjoint
            If pRelOp.Disjoint(pLimiteB) Then
                'Si l'indice est F alors retourner l'indice d'intersection est respecté
                If sIndice = "F" Then Return True

                'Si les limites ne sont pas disjoint
            Else
                'Si l'indice est T alors retourner l'indice d'intersection est respecté
                If sIndice = "T" Then Return True

                'Vérifier si la dimension d'une des limites est un point
                If pLimiteA.Dimension = esriGeometryDimension.esriGeometry0Dimension _
                Or pLimiteB.Dimension = esriGeometryDimension.esriGeometry0Dimension Then
                    'Si l'indice est 0 alors retourner l'indice d'intersection est respecté
                    If sIndice = "0" Then Return True

                    'Si la dimension des limites sont des lignes
                Else
                    'Interface pour traiter l'intersection entre les limites
                    pTopoOp = CType(pLimiteA, ITopologicalOperator2)

                    'Extraire l'intersection de ligne
                    pGeometry = pTopoOp.Intersect(pLimiteB, esriGeometryDimension.esriGeometry1Dimension)

                    'Vérifier si l'intersection est de type point
                    If pGeometry.IsEmpty Then
                        'Si l'indice est 0 alors retourner l'indice d'intersection est respecté
                        If sIndice = "0" Then Return True

                        'Si l'intersection est de type ligne
                    Else
                        'Si l'indice est 1 alors retourner l'indice d'intersection est respecté
                        If sIndice = "1" Then Return True
                    End If
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pRelOp = Nothing
            pTopoOp = Nothing
            pLimiteA = Nothing
            pLimiteB = Nothing
            pGeometry = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'indiquer si l'indice d'intersection entre la limite de A et l'extérieur de B est respecté.
    '''</summary>
    ''' 
    '''<param name="sIndice"> Indice d'intersection du masque spatial (*,T,F,0,1).</param>
    '''<param name="pGeometryA"> Interface contenant la géométrie à traiter (A).</param>
    '''<param name="pGeometryB"> Interface contenant la géométrie en relation (B).</param>
    '''
    Private Function IntersectionLimiteExterieur(ByVal sIndice As String, ByVal pGeometryA As IGeometry, ByVal pGeometryB As IGeometry) As Boolean
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface pour extraire la limite et traiter l'intersection.
        Dim pLimiteA As IGeometry = Nothing             'Interface contenant la limite de la géométrie A.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant le résultat de l'intersection des limites.

        'L'intersection n'est pas respecté par défaut
        IntersectionLimiteExterieur = False

        Try
            'Si l'indice est * alors retourner l'indice d'intersection est respecté
            If sIndice = "*" Then Return True

            'Interface pour extraire la limite
            pTopoOp = CType(pGeometryA, ITopologicalOperator2)
            'Interface contenant le limite de A
            pLimiteA = pTopoOp.Boundary

            'Interface pour vérifier l'intersection
            pTopoOp = CType(pLimiteA, ITopologicalOperator2)

            'Extraire l'intersection entre la limite de A et l'extérieur de B
            pGeometry = pTopoOp.Difference(pGeometryB)

            'Vérifier si le résultat de l'intersection est vide
            If pGeometry.IsEmpty Then
                'Si l'indice est F alors retourner l'indice d'intersection est respecté
                If sIndice = "F" Then Return True

                'Si le résultat de l'intersection n'est pas vide
            Else
                'Si l'indice est T alors retourner l'indice d'intersection est respecté
                If sIndice = "T" Then Return True

                'Vérifier si la dimension de l'intersection est de type point
                If pGeometry.Dimension = esriGeometryDimension.esriGeometry0Dimension Then
                    'Si l'indice est 0 alors retourner l'indice d'intersection est respecté
                    If sIndice = "0" Then Return True

                    'Si le résultat de l'intersection est de type ligne
                Else
                    'Si l'indice est 1 alors retourner l'indice d'intersection est respecté
                    If sIndice = "1" Then Return True
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pLimiteA = Nothing
            pGeometry = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'indiquer si l'indice d'intersection entre la limite de A et l'intérieur de B est respecté.
    '''</summary>
    ''' 
    '''<param name="sIndice"> Indice d'intersection du masque spatial (*,T,F,0,1).</param>
    '''<param name="pGeometryA"> Interface contenant la géométrie à traiter (A).</param>
    '''<param name="pGeometryB"> Interface contenant la géométrie en relation (B).</param>
    '''
    Private Function IntersectionLimiteInterieur(ByVal sIndice As String, ByVal pGeometryA As IGeometry, ByVal pGeometryB As IGeometry) As Boolean
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface pour extraire la limite et traiter l'intersection.
        Dim pLimiteA As IGeometry = Nothing             'Interface contenant la limite de la géométrie A.
        Dim pLimiteB As IGeometry = Nothing             'Interface contenant la limite de la géométrie B.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant le résultat de l'intersection des limites.

        'L'intersection n'est pas respecté par défaut
        IntersectionLimiteInterieur = False

        Try
            'Si l'indice est * alors retourner l'indice d'intersection est respecté
            If sIndice = "*" Then Return True

            'Interface pour extraire la limite
            pTopoOp = CType(pGeometryA, ITopologicalOperator2)
            'Interface contenant le limite de A
            pLimiteA = pTopoOp.Boundary

            'Interface pour extraire la limite
            pTopoOp = CType(pGeometryB, ITopologicalOperator2)
            'Interface contenant le limite de B
            pLimiteB = pTopoOp.Boundary

            'Interface pour vérifier l'intersection
            pTopoOp = CType(pLimiteA, ITopologicalOperator2)

            'Vérifier si la dimension de la limite est un point
            If pLimiteA.Dimension = esriGeometryDimension.esriGeometry0Dimension Then
                'Extraire l'intersection entre la limite de A et la géométrie de B
                pGeometry = pTopoOp.Intersect(pGeometryB, esriGeometryDimension.esriGeometry0Dimension)

                'Si la dimension de la limite est une ligne
            Else
                'Extraire l'intersection entre la limite de A et la géométrie de B
                pGeometry = pTopoOp.Intersect(pGeometryB, esriGeometryDimension.esriGeometry1Dimension)
            End If

            'Vérifier si la limite de B peut être enlevé du résultat
            If pLimiteB.Dimension >= pGeometry.Dimension Then
                'Interface pour vérifier extraire la limite de B
                pTopoOp = CType(pGeometry, ITopologicalOperator2)

                'Extraire la limite de B
                pGeometry = pTopoOp.Difference(pLimiteB)
            End If

            'Vérifier si le résultat de l'intersection est vide
            If pGeometry.IsEmpty Then
                'Si l'indice est F alors retourner l'indice d'intersection est respecté
                If sIndice = "F" Then Return True

                'Si le résultat de l'intersection n'est pas vide
            Else
                'Si l'indice est T alors retourner l'indice d'intersection est respecté
                If sIndice = "T" Then Return True

                'Vérifier si la dimension de l'intersection est de type point
                If pGeometry.Dimension = esriGeometryDimension.esriGeometry0Dimension Then
                    'Si l'indice est 0 alors retourner l'indice d'intersection est respecté
                    If sIndice = "0" Then Return True

                    'Si le résultat de l'intersection est de type ligne
                Else
                    'Si l'indice est 1 alors retourner l'indice d'intersection est respecté
                    If sIndice = "1" Then Return True
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pLimiteA = Nothing
            pLimiteB = Nothing
            pGeometry = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'indiquer si l'indice d'intersection entre la l'intérieur de A et l'extérieur de B est respecté.
    '''</summary>
    ''' 
    '''<param name="sIndice"> Indice d'intersection du masque spatial (*,T,F,0,1,2).</param>
    '''<param name="pGeometryA"> Interface contenant la géométrie à traiter (A).</param>
    '''<param name="pGeometryB"> Interface contenant la géométrie en relation (B).</param>
    '''
    Private Function IntersectionInterieurExterieur(ByVal sIndice As String, ByVal pGeometryA As IGeometry, ByVal pGeometryB As IGeometry) As Boolean
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface pour extraire la limite et traiter l'intersection.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant le résultat de l'intersection des limites.

        'L'intersection n'est pas respecté par défaut
        IntersectionInterieurExterieur = False

        Try
            'Si l'indice est * alors retourner l'indice d'intersection est respecté
            If sIndice = "*" Then Return True

            'Interface pour extraire la limite
            pTopoOp = CType(pGeometryA, ITopologicalOperator2)

            'Extraire l'intersection entre la limite de A et l'extérieur de B
            pGeometry = pTopoOp.Difference(pGeometryB)

            'Vérifier si le résultat de l'intersection est vide
            If pGeometry.IsEmpty Then
                'Si l'indice est F alors retourner l'indice d'intersection est respecté
                If sIndice = "F" Then Return True

                'Si le résultat de l'intersection n'est pas vide
            Else
                'Si l'indice est T alors retourner l'indice d'intersection est respecté
                If sIndice = "T" Then Return True

                'Vérifier si la dimension de l'intersection est de type point
                If pGeometry.Dimension = esriGeometryDimension.esriGeometry0Dimension Then
                    'Si l'indice est 0 alors retourner l'indice d'intersection est respecté
                    If sIndice = "0" Then Return True

                    'Si le résultat de l'intersection est de type ligne
                ElseIf pGeometry.Dimension = esriGeometryDimension.esriGeometry1Dimension Then
                    'Si l'indice est 1 alors retourner l'indice d'intersection est respecté
                    If sIndice = "1" Then Return True

                    'Si le résultat de l'intersection est de type surface
                Else
                    'Si l'indice est 2 alors retourner l'indice d'intersection est respecté
                    If sIndice = "2" Then Return True
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pGeometry = Nothing
        End Try
    End Function
#End Region
End Class
