Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DataSourcesRaster

'**
'Nom de la composante : clsSensEcoulement.vb
'
'''<summary>
''' Classe qui permet de traiter le sens d'écoulement des cours d'eau simple.
''' 
''' La classe permet de traiter les deux attributs MNE et COURBE.
''' 
''' MNE : Le sens d'écoulement est déterminé à partir des valeurs d'élévation contenues dans un MNE de type maticielle.
''' COURBE : Le sens d'écoulement est déterminé à partir des valeurs d'élévation contenues dans une attribut d'une classe de courbes de niveau de type vectorielle.
''' CONTINUITE : Le sens d'écoulement est validé à partir de la continuité des cours d'eau simple.
''' NOEUD : Le sens d'écoulement est validé à partir du nombre de First (au moins 1) et de Last (au moins 1) pour chaque Noeud qui possède au moins 2 edges.
''' RESEAU : Le sens d'écoulement est validé à partir des réseaux créés.
''' CREER : Créer les occurences de réseaux de cours d'eau interconnectés.
''' 
''' Note : Le sens d'écoulement des cours va du haut en bas. 
''' 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 07 janvier 2016
'''</remarks>
'''
Public Class clsSensEcoulement
    Inherits clsValeurAttribut

    '''<summary>Contient l'information supplémentaire du traitement à effectuer.</summary>
    Protected gsInformation As String = ""
    '''<summary>Contient la tolérance d'égalité d'élévation.</summary>
    Protected gdTolEgaliteZ As Double = 0

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet avec les valeurs par défaut.
    ''' 
    '''</summary>
    '''
    Public Sub New()
        Try
            'Définir les valeurs par défaut
            NomAttribut = "NOEUD"
            Expression = "VRAI"
            gsInformation = ""

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
    '''<param name="sInformation"> Information supplémentaire de traitement.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer,
                   ByVal sNomAttribut As String, ByVal sExpression As String, Optional ByVal sInformation As String = "")
        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            NomAttribut = sNomAttribut
            Expression = sExpression
            gsInformation = sInformation
            gpFeatureLayersRelation = New Collection
            gpRasterLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'instancier la classe en objet.
    ''' 
    '''<param name="pFeatureLayerSelection"> Interface contenant le FeatureLayer de sélection à traiter.</param>
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression (1) à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pFeatureLayerSelection As IFeatureLayer, ByVal sParametres As String)
        'Déclarer les variables de travail

        Try
            'Définir les valeurs par défaut
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres
            gpFeatureLayersRelation = New Collection
            gpRasterLayersRelation = New Collection

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
    '''<param name="sParametres"> Paramètres contenant le nom de l'attribut (0) et l'expression (1) à traiter.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByRef pMap As IMap, ByRef pFeatureLayerSelection As IFeatureLayer, ByVal sParametres As String)
        'Déclarer les variables de travail

        Try
            'Définir les valeurs par défaut
            Map = pMap
            FeatureLayerSelection = pFeatureLayerSelection
            Parametres = sParametres
            gpFeatureLayersRelation = New Collection
            gpRasterLayersRelation = New Collection

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        'Vider la mémoire
        gsInformation = Nothing
        gdTolEgaliteZ = Nothing
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
            Nom = "SensEcoulement"
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
            'Vérifier la présence de l'information supplémentaire
            If gsInformation <> "" Then
                'Retourner aussi l'information supplémentaire
                Parametres = Parametres & " " & gsInformation
            End If
        End Get
        Set(ByVal value As String)
            'Déclarer les variables de travail
            Dim params() As String      'Liste des paramètres 0:NomAttribut, 1:Expression [2:Information]

            'Mettre en majuscule les paramètres
            value = value.ToUpper
            'Extraire les paramètres
            params = value.Split(CChar(" "))
            'Vérifier si les deux paramètres sont présents
            If params.Length < 2 Then Err.Raise(1, , "Deux paramètres sont obligatoires: ATTRIBUT EXPRESSION [AttributElevation]")

            'Définir les valeurs par défaut
            gsNomAttribut = params(0)
            gsExpression = params(1)

            'Vérifier si aucune commande de comparaison
            If params.Length < 3 Then
                'Aucun attribut d'élévation
                gsInformation = ""

                'Si l'information supplémentaire est spécifiée
            ElseIf params.Length = 3 Then
                'Définir l'information supplémentaire
                gsInformation = params(2)
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
                'Définir le paramètre pour trouver les cours d'eau qui ne respecte pas le sens d'écoulement basé sur le nombre de first>0 et de last>0 à un noeud.
                ListeParametres.Add("NOEUD VRAI")
                'Définir le paramètre pour trouver les cours d'eau qui respecte le sens d'écoulement basé sur le nombre de first>0 et de last>0 à un noeud.
                ListeParametres.Add("NOEUD FAUX")

                'Définir le paramètre pour trouver les lignes de cours d'eau qui ne respecte pas le sens d'écoulement basé sur les élévations d'un MNE et une tolérance d'égalité.
                ListeParametres.Add("LIGNE VRAI 1.0")
                'Définir le paramètre pour trouver les lignes de cours d'eau qui respecte le sens d'écoulement basé sur les élévations d'un MNE et une tolérance d'égalité.
                ListeParametres.Add("LIGNE FAUX 1.0")

                'Définir le paramètre pour trouver les polylignes de cours d'eau qui ne respecte pas le sens d'écoulement basé sur les élévations d'un MNE et une tolérance d'égalité.
                ListeParametres.Add("POLYLIGNE VRAI 1.0")
                'Définir le paramètre pour trouver les polylignes de cours d'eau qui respecte le sens d'écoulement basé sur les élévations d'un MNE et une tolérance d'égalité.
                ListeParametres.Add("POLYLIGNE FAUX 1.0")

                'Définir le paramètre pour trouver les réseaux de cours d'eau qui ne respecte pas le sens d'écoulement basé sur les fins de réseaux créés.
                ListeParametres.Add("RESEAU VRAI")
                'Définir le paramètre pour trouver les cours d'eau qui respecte le sens d'écoulement basé sur les fins de réseaux créés.
                ListeParametres.Add("RESEAU FAUX")

                'Définir le paramètre pour créer les réseaux de cours d'eau interconnectés simplifiés.
                ListeParametres.Add("CREER VRAI RESEAU")
                'Définir le paramètre pour créer les réseaux de cours d'eau interconnectés non simplifiés.
                ListeParametres.Add("CREER FAUX RESEAU")

                'Définir le paramètre pour créer les fins de réseaux de cours d'eau interconnectés basé sur les élévations d'un MNE.
                ListeParametres.Add("CREER VRAI FIN_RESEAU")
                'Définir le paramètre pour créer les débuts de réseaux de cours d'eau interconnectés basé sur les élévations d'un MNE.
                ListeParametres.Add("CREER FAUX FIN_RESEAU")
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
    ''' 
    '''<return>Boolean qui indique si l'attribut est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function AttributValide() As Boolean
        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            AttributValide = False
            gsMessage = "L'attribut est invalide : " & gsNomAttribut & " " & gsInformation

            'Vérifier si l'attribut est NOEUD ou RESEAU
            If gsNomAttribut = "NOEUD" Or gsNomAttribut = "RESEAU" Then
                'Vérifier si le type de traitement est valide
                If gsInformation = "" Then
                    'L'attribut est valide
                    AttributValide = True
                    gsMessage = "La contrainte est valide."

                    'Si le type de traitement est invalide
                Else
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : Aucun type de traitement n'est requis : " & gsNomAttribut & " " & gsInformation
                End If

            ElseIf gsNomAttribut = "LIGNE" Or gsNomAttribut = "POLYLIGNE" Then
                'Vérifier si l'information de tolérance est numérique
                If TestDBL(gsInformation) Then
                    'Définir la tolérance d'égalité d'élévation
                    gdTolEgaliteZ = ConvertDBL(gsInformation)

                    'Vérifier si la tolérance est valide
                    If gdTolEgaliteZ >= 0 Then
                        'L'attribut est valide
                        AttributValide = True
                        gsMessage = "La contrainte est valide."

                        'Si la tolérance d'égalité d'élévation est inférieure à zéro
                    Else
                        'Définir le message d'erreur
                        gsMessage = "ERREUR : La tolérance d'égalité d'élévation est inférieure à zéro : " & gsNomAttribut & " " & gsInformation
                    End If

                    'Si la tolérance d'égalité d'élévation n'est pas numérique
                Else
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : La tolérance d'égalité d'élévation n'est pas numérique : " & gsNomAttribut & " " & gsInformation
                End If

                'si l'attribut est CREER
            ElseIf gsNomAttribut = "CREER" Then
                'Vérifier si le type de traitement est valide
                If gsInformation = "FIN_RESEAU" Or gsInformation = "RESEAU" Then
                    'L'attribut est valide
                    AttributValide = True
                    gsMessage = "La contrainte est valide."

                    'Si le type de traitement est invalide
                Else
                    'Définir le message d'erreur
                    gsMessage = "ERREUR : Le type de traitement est invalide : " & gsNomAttribut & " " & gsInformation
                End If

                'Si l'attribut est invalide
            Else
                'Définir le message d'erreur
                gsMessage = "ERREUR : L'attribut est invalide : " & gsNomAttribut
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'indiquer si l'expression utilisée pour extraire l'élément de découpage est valide.
    ''' 
    '''<return>Boolean qui indique si l'expression est valide.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function ExpressionValide() As Boolean
        Try
            'Définir la valeur par défaut
            ExpressionValide = True
            gsMessage = "La contrainte est valide."

            'Vérifier si l'expression est valide
            If Expression <> "VRAI" And Expression <> "FAUX" Then
                'Définir l'erreur
                ExpressionValide = False
                gsMessage = "ERREUR : L'expression est invalide (VRAI:HAUT-BAS ou FAUX:BAS-HAUT)"
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
    Public Overloads Overrides Function FeatureLayersRelationValide() As Boolean
        'Déclarer les variables de travail
        Dim pFeatureLayer As IFeatureLayer = Nothing     'Interface contenant le FeatureLayer des courbes

        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            FeatureLayersRelationValide = False
            gsMessage = "ERREUR : Le FeatureLayer des fins de réseaux est invalide."

            'Vérifier si le nom de l'attribut est RESEAU
            If gsNomAttribut = "RESEAU" Then
                'Vérifier si les FeatureLayers en relation sont absent
                If gpFeatureLayersRelation Is Nothing Then
                    'Message d'erreur
                    gsMessage = "ERREUR : Le FeatureLayer des fins de réseaux est absent et ne peut être ajouté."

                    'Si des FeatureLayers en relation sont présents
                Else
                    'Retirer le FeatureLayer en relation correspondant au FeatureLayer de sélection
                    If gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then gpFeatureLayersRelation.Remove(gpFeatureLayerSelection.Name)

                    'Vérifier si aucun élément en relation 
                    If gpFeatureLayersRelation.Count = 0 Then
                        'Message d'erreur
                        gsMessage = "ERREUR : Le FeatureLayer des fins de réseaux en relation est absent."

                        'Vérifier la présence des FeatureLayers en relation
                    ElseIf gpFeatureLayersRelation.Count = 1 Then
                        'Traiter tous les FeatureLayers en relation
                        pFeatureLayer = CType(gpFeatureLayersRelation.Item(1), IFeatureLayer)

                        'Vérifier si le FeatureLayer est valide
                        If pFeatureLayer IsNot Nothing Then
                            'Vérifier si la FeatureClass est invalide
                            If pFeatureLayer.FeatureClass IsNot Nothing Then
                                'Vérifier si la FeatureClass est invalide
                                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                                    'La contrainte est valide
                                    FeatureLayersRelationValide = True
                                    gsMessage = "La contrainte est valide."

                                    'Si le type de géométrie est invalide
                                Else
                                    'Message d'erreur
                                    gsMessage = "ERREUR : La FeatureClass des fins de réseaux doit être de type <Point>."
                                End If

                                'Si la classe est invalide
                            Else
                                'Message d'erreur
                                gsMessage = "ERREUR : La FeatureClass des fins de réseaux est invalide."
                            End If

                            'Si la Layer est invalide
                        Else
                            'Message d'erreur
                            gsMessage = "ERREUR : Le FeatureLayer des fins de réseaux est invalide."
                        End If

                        'Si plusieurs Layer en relation sont sélectionnés
                    Else
                        'Message d'erreur
                        gsMessage = "ERREUR : Un seul FeatureLayer des fins de réseaux doit être sélectionné."
                    End If
                End If

                'Si aucun FeatureLayer en relation n'est nécessaire pour le traitement
            Else
                'La contrainte est valide
                FeatureLayersRelationValide = True
                gsMessage = "La contrainte est valide."
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
    Public Overloads Overrides Function RasterLayersRelationValide() As Boolean
        'Déclarer les variables de travail
        Dim pRasterLayer As IRasterLayer = Nothing     'Interface contenant le RasterLayer du MNE

        Try
            'Définir la valeur par défaut, la contrainte est invalide.
            RasterLayersRelationValide = False
            gsMessage = "ERREUR : Le FeatureLayer de MNE est invalide."

            'Vérifier si le nom de l'attribut est LIGNE, POLYLIGNE ou CREER-FIN_RESEAU
            If gsNomAttribut = "LIGNE" Or gsNomAttribut = "POLYLIGNE" _
            Or (gsNomAttribut = "CREER" And gsInformation = "FIN_RESEAU") Then
                'Vérifier si les FeatureLayers en relation sont absent
                If gpRasterLayersRelation Is Nothing Then
                    'Message d'erreur
                    gsMessage = "ERREUR : Le RasterLayer du MNE en relation doit être spécifié."

                    'Si des RasterLayers en relation sont présents
                Else
                    'Vérifier si aucun élément en relation 
                    If gpRasterLayersRelation.Count = 0 Then
                        'Message d'erreur
                        gsMessage = "ERREUR : Le RasterLayer du MNE en relation est absent."

                        'Vérifier la présence des FeatureLayers en relation
                    ElseIf gpRasterLayersRelation.Count = 1 Then
                        'Traiter tous les RasterLayers en relation
                        pRasterLayer = CType(gpRasterLayersRelation.Item(1), IRasterLayer)

                        'Vérifier si le RasterLayer est valide
                        If pRasterLayer IsNot Nothing Then
                            'Vérifier si le Raster est invalide
                            If pRasterLayer.Raster IsNot Nothing Then
                                'Vérifier si le Raster est invalide
                                If pRasterLayer.BandCount = 1 Then
                                    'La contrainte est valide
                                    RasterLayersRelationValide = True
                                    gsMessage = "La contrainte est valide."

                                    'Si le type de géométrie est invalide
                                Else
                                    'Message d'erreur
                                    gsMessage = "ERREUR : Le RasterLayer de MNE doit posséder une seule bande."
                                End If

                            Else
                                'Message d'erreur
                                gsMessage = "ERREUR : Le RasterLayer de MNE est invalide."
                            End If

                        Else
                            'Message d'erreur
                            gsMessage = "ERREUR : Le RasterLayer du MNE est invalide."
                        End If

                    Else
                        'Message d'erreur
                        gsMessage = "ERREUR : Un seul RasterLayer de MNE doit être sélectionné."
                    End If
                End If

                'Si aucun RasterLayer de MNE n'est nécessaire pour le traitement 
            Else
                'La contrainte est valide
                RasterLayersRelationValide = True
                gsMessage = "La contrainte est valide."
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
    ''' Routine qui permet de sélectionner les éléments du FeatureLayer dont les géométries des éléments respectent ou non le sens d'écoulement.
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
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Public Overloads Overrides Function Selectionner(ByRef pTrackCancel As ITrackCancel, Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing              'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pSelectionSet As ISelectionSet = Nothing                'Interface utilisé pour extraire les éléments sélectionnés du FeatureLayer.
        Dim pCursor As ICursor = Nothing                            'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing              'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeatureLayerReseau As IFeatureLayer = Nothing          'Interface contenant les éléments de réseaux de cours d'eau interconnectés.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing             'Interface contenant la topologie.
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

            'Enlever la sélection
            pFeatureSel.Clear()

            'Si le nom de l'attribut est CREER
            If gsNomAttribut = "CREER" Then
                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                'Initialiser la liste des Layers en relation pour créer la topologie
                gpFeatureLayersRelation = New Collection
                'Vérifier si la Layer de sélection est absent dans les Layers relations
                If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                    'Ajouter le layer de sélection dans les layers en relation
                    gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                End If
                'Interface pour extraire la tolérance de précision de la référence spatiale
                pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                'Création de la topologie
                pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)

                'Si le type de traitement est la création des fin de réseaux
                If gsInformation = "FIN_RESEAU" Then
                    'Interfaces pour extraire les éléments sélectionnés
                    pSelectionSet.Search(Nothing, False, pCursor)
                    pFeatureCursor = CType(pCursor, IFeatureCursor)

                    'Vérifier si on veut créer les fins de réseaux
                    If gsExpression = "VRAI" Then
                        'Afficher le message du traitement à exécuter
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter la création des fins de réseaux (" & gpFeatureLayerSelection.Name & ") ..."
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur("FIN_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPoint)

                        'Si on veut créer les débuts de réseaux
                    Else
                        'Afficher le message du traitement à exécuter
                        If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter la création des débuts de réseaux (" & gpFeatureLayerSelection.Name & ") ..."
                        'Créer la classe d'erreurs au besoin
                        CreerFeatureClassErreur("DEBUT_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryMultipoint)
                    End If

                    'Traiter le FeatureLayer
                    Selectionner = TraiterSensEcoulementCreerFinReseau(pTopologyGraph, pFeatureCursor, pTrackCancel, bEnleverSelection)

                    'Si le type de traitement est la création des réseaux
                ElseIf gsInformation = "RESEAU" Then
                    'Créer la classe d'erreurs au besoin
                    CreerFeatureClassErreur("RESEAU_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                    'Afficher le message du traitement à exécuter
                    If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter la création des réseaux (" & gpFeatureLayerSelection.Name & ") ..."
                    'Traiter le FeatureLayer
                    Selectionner = TraiterSensEcoulementCreerReseau(pTopologyGraph, pTrackCancel, bEnleverSelection)
                End If

                'Si le nom de l'attribut est NOEUD
            ElseIf gsNomAttribut = "NOEUD" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection, esriGeometryType.esriGeometryPoint)

                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                'Initialiser la liste des Layers en relation pour créer la topologie
                gpFeatureLayersRelation = New Collection
                'Vérifier si la Layer de sélection est absent dans les Layers relations
                If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                    'Ajouter le layer de sélection dans les layers en relation
                    gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                End If
                'Interface pour extraire la tolérance de précision de la référence spatiale
                pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                'Création de la topologie
                pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)

                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter le sens d'écoulement des noeuds selon le nombre de First>0 et de Last>0 (" & gpFeatureLayerSelection.Name & ") ..."
                'Traiter le FeatureLayer
                Selectionner = TraiterSensEcoulementNoeud(pTopologyGraph, pTrackCancel, ((gsExpression = "VRAI") = bEnleverSelection))

                'Si le nom de l'attribut est LIGNE
            ElseIf gsNomAttribut = "LIGNE" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                'Interfaces pour extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter le sens d'écoulement des lignes selon un MNE (" & gpFeatureLayerSelection.Name & ") ..."
                'Traiter le FeatureLayer
                Selectionner = TraiterSensEcoulementLigne(pFeatureCursor, pTrackCancel, bEnleverSelection)

                'Si le nom de l'attribut est POLYLIGNE
            ElseIf gsNomAttribut = "POLYLIGNE" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                'Initialiser la liste des Layers en relation pour créer la topologie
                gpFeatureLayersRelation = New Collection
                'Vérifier si la Layer de sélection est absent dans les Layers relations
                If Not gpFeatureLayersRelation.Contains(gpFeatureLayerSelection.Name) Then
                    'Ajouter le layer de sélection dans les layers en relation
                    gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)
                End If
                'Interface pour extraire la tolérance de précision de la référence spatiale
                pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                'Création de la topologie
                pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)

                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter le sens d'écoulement des polylignes selon un MNE (" & gpFeatureLayerSelection.Name & ") ..."
                'Traiter le FeatureLayer
                Selectionner = TraiterSensEcoulementPolyligne(pTopologyGraph, pTrackCancel, ((gsExpression = "VRAI") = bEnleverSelection))

                'Si le nom de l'attribut est RESEAU
            ElseIf gsNomAttribut = "RESEAU" Then
                'Créer la classe d'erreurs au besoin
                CreerFeatureClassErreur(Me.Nom & "_" & gpFeatureLayerSelection.FeatureClass.AliasName, gpFeatureLayerSelection)

                'Conserver le FeatureLayer de fins de réseaux
                pFeatureLayerReseau = CType(gpFeatureLayersRelation.Item(1), IFeatureLayer)

                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Créer la topologie (" & gpFeatureLayerSelection.Name & ") ..."
                'Initialiser la liste des Layers en relation pour créer la topologie
                gpFeatureLayersRelation = New Collection
                'Ajouter le layer de sélection dans les layers en relation
                gpFeatureLayersRelation.Add(gpFeatureLayerSelection, gpFeatureLayerSelection.Name)

                'Interface pour extraire la tolérance de précision de la référence spatiale
                pSRTolerance = CType(Selectionner.SpatialReference, ISpatialReferenceTolerance)
                'Création de la topologie
                pTopologyGraph = CreerTopologyGraph(EnveloppeSelectionSet(pSelectionSet, Selectionner.SpatialReference), gpFeatureLayersRelation, pSRTolerance.XYTolerance)

                'Initialiser la liste des Layers en relation
                gpFeatureLayersRelation = New Collection
                'Ajouter le layer de sélection dans les layers en relation
                gpFeatureLayersRelation.Add(pFeatureLayerReseau, pFeatureLayerReseau.Name)

                'Afficher le message du traitement à exécuter
                If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Traiter le sens d'écoulement des réseaux selon les fins de réseaux (" & gpFeatureLayerSelection.Name & ") ..."
                'Traiter le FeatureLayer
                Selectionner = TraiterSensEcoulementReseau(pTopologyGraph, pTrackCancel, bEnleverSelection)
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
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeatureLayerReseau = Nothing
            pTopologyGraph = Nothing
            pSRTolerance = Nothing
            'Variables globales
            gpFeatureCursorErreur = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Fonction qui permet de créer les débuts ou les fins des réseaux de cours d'eau interconnectés.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie des cours d'eau simple.</param>
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les débuts ou fins de réseaux de cours d'eau interconnectés.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterSensEcoulementCreerFinReseau(ByRef pTopologyGraph As ITopologyGraph4, ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                         Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pRasterLayer As IRasterLayer = Nothing          'Interface contenant le Layer du MNE.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant les points de début et de fin de réseau.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour ajouter les points de fin de réseau.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour définir les points de début de réseau.
        Dim pFinReseau As IPoint = Nothing                  'Interface contenat le sommet de fin de réseau.

        Try
            'Définir la géométrie par défaut
            TraiterSensEcoulementCreerFinReseau = New GeometryBag
            TraiterSensEcoulementCreerFinReseau.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSensEcoulementCreerFinReseau, IGeometryCollection)

            'Définir le RasterLayer du MNE
            pRasterLayer = CType(gpRasterLayersRelation.Item(1), IRasterLayer)

            'Si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline And pRasterLayer IsNot Nothing Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Créer un nouveau multipoint vide
                    pMultipoint = New Multipoint
                    pMultipoint.SpatialReference = TraiterSensEcoulementCreerFinReseau.SpatialReference

                    'Trouver le point de fin du réseau et de début de réseau
                    TrouverFinReseau(pTopologyGraph, pRasterLayer, pFeature, pMultipoint, pFinReseau)

                    'Vérifier si un point de fin a été trouvé
                    If pFinReseau IsNot Nothing Then
                        'Vérifier si on doit créer les fins de réseau
                        If gsExpression = "VRAI" Then
                            'Ajouter le point de fin de réseau
                            pGeomSelColl.AddGeometry(pFinReseau)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Zmin=" & pFinReseau.Z.ToString _
                                                & " /" & gsNomAttribut & " " & gsExpression & " " & gsInformation, pFinReseau, CSng(pFinReseau.Z))

                            'Si on doit créer les débuts de réseau
                        Else
                            'Interface pour définir les points de début de réseau
                            pTopoOp = CType(pMultipoint, ITopologicalOperator2)
                            'Définir les points de début de réseau
                            pMultipoint = CType(pTopoOp.Difference(pFinReseau), IMultipoint)
                            'Ajouter les points de débuts de réseau
                            pGeomSelColl.AddGeometry(pMultipoint)
                            'Interface pour ajouter des sommets dans le multipoint
                            pPointColl = CType(pMultipoint, IPointCollection)
                            'Écrire une erreur
                            EcrireFeatureErreur("OID=" & pFeature.OID.ToString _
                                                & " #Zmin=" & pPointColl.Point(0).Z.ToString & ", NbPoint=" & pPointColl.PointCount.ToString _
                                                & " /" & gsNomAttribut & " " & gsExpression & " " & gsInformation, pMultipoint, CSng(pPointColl.Point(0).Z))
                        End If
                    End If

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
            pRasterLayer = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pMultipoint = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
            pFinReseau = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de trouver les fins des réseaux de cours d'eau interconnectés.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie des cours d'eau simple.</param>
    '''<param name="pRasterLayer"> Interface contenant le Layer du MNE.</param>
    '''<param name="pFeature"> Interface contenant l'élément à traiter.</param>
    '''<param name="pMultipoint"> Interface contenant les points de début et de fin de réseau.</param>
    '''<param name="pFinReseau"> Interface contenant le sommet de fin de réseau.</param>
    ''' 
    '''</summary>
    '''
    Private Sub TrouverFinReseau(ByRef pTopologyGraph As ITopologyGraph4, ByRef pRasterLayer As IRasterLayer, ByRef pFeature As IFeature, _
                                 ByRef pMultipoint As IMultipoint, ByRef pFinReseau As IPoint)
        'Déclarer les variables de travail
        Dim pRaster2 As IRaster2 = Nothing                  'Interface contenant le MNE pour extraire l'élévation.
        Dim pGeodataset As IGeoDataset = Nothing            'Interface contenant la référence spatiale du MNE.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la géométrie d'un Edge de début ou de fin.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour ajouter les points de fin de réseau.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface pour extraire les Nodes.
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface pour extraire les Nodes du Edge traité.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node à traiter.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le Edge traité.
        Dim pPointFrom As IPoint = Nothing                  'Interface contenant un sommet de début de réseau.
        Dim pPointTo As IPoint = Nothing                    'Interface contenant un sommet de fin de ligne.
        Dim iCol As Integer = Nothing               'Contient la rangée correspondant à la valeur X dans le MNE.
        Dim iRow As Integer = Nothing               'Contient la colonne correspondant à la valeur Y dans le MNE.

        Try
            'Interface contenant le MNE pour extraire l'élévation.
            pRaster2 = CType(pRasterLayer.Raster, IRaster2)
            'Interface pour extraire la référence spatiale du MNE
            pGeodataset = CType(pRaster2, IGeoDataset)

            'Interface pour ajouter des sommets dans le multipoint
            pPointColl = CType(pMultipoint, IPointCollection)

            'Extraire les noeuds de l'élément
            pEnumTopoNode = pTopologyGraph.GetParentNodes(gpFeatureLayerSelection.FeatureClass, pFeature.OID)

            'Extraire le premier Node à traiter
            pEnumTopoNode.Reset()
            pTopoNode = pEnumTopoNode.Next()

            'Traiter tous les nodes
            Do Until pTopoNode Is Nothing
                'Interface pour extraire les edges du Node
                pEnumNodeEdge = pTopoNode.Edges(True)

                'Vérifier si le nombre de Edges est 1
                If pEnumNodeEdge.Count = 1 Then
                    'Interface contenant le point du noeud de début
                    pPointFrom = CType(pTopoNode.Geometry, IPoint)
                    pPointFrom.Project(pGeodataset.SpatialReference)

                    'Extraire la position de la rangée à partir de la valeur X
                    iCol = pRaster2.ToPixelColumn(pPointFrom.X)
                    'Extraire la position de la colonne à partir de la valeur Y
                    iRow = pRaster2.ToPixelRow(pPointFrom.Y)
                    'Extraire la valeur d'élévation du MNE
                    pPointFrom.Z = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))

                    'Vérifier si aucun sommet ajouté
                    If pPointColl.PointCount = 0 Then
                        'Définir le point de fin de réseau
                        pFinReseau = pPointFrom
                    Else
                        'Si le points est plus bas que le premier
                        If pPointFrom.Z < pFinReseau.Z Then
                            'Définir le point de fin de réseau
                            pFinReseau = pPointFrom
                        End If
                    End If

                    'Ajouter le point de début ou de fin de réseau
                    pPointColl.AddPoint(pPointFrom)

                    'Extraire le edge du Node
                    pEnumNodeEdge.Reset()
                    pEnumNodeEdge.Next(pTopoEdge, True)

                    'Vérifier si le edge est trouvé
                    If pTopoEdge IsNot Nothing Then
                        'Extraire la géométrie du edge
                        pPolyline = CType(pTopoEdge.Geometry, IPolyline)

                        'Vérifier si le point de début est le premier du Edge
                        If pPointFrom.X = pPolyline.FromPoint.X And pPointFrom.Y = pPolyline.FromPoint.Y Then
                            'Définir le point de fin
                            pPointTo = pPolyline.ToPoint
                            'Si le point de début est le dernier du Edge
                        Else
                            'Définir le point de fin
                            pPointTo = pPolyline.FromPoint
                        End If
                        'Projeter le point
                        pPointTo.Project(pGeodataset.SpatialReference)
                        'Extraire la position de la rangée à partir de la valeur X
                        iCol = pRaster2.ToPixelColumn(pPointTo.X)
                        'Extraire la position de la colonne à partir de la valeur Y
                        iRow = pRaster2.ToPixelRow(pPointTo.Y)
                        'Extraire la valeur d'élévation du MNE
                        pPointTo.Z = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))
                    End If
                End If

                'Extraire le prochain Node à traiter
                pTopoNode = pEnumTopoNode.Next()
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pRaster2 = Nothing
            pGeodataset = Nothing
            pPolyline = Nothing
            pPointColl = Nothing
            pEnumTopoNode = Nothing
            pEnumNodeEdge = Nothing
            pTopoNode = Nothing
            pTopoEdge = Nothing
            pPointFrom = Nothing
            pPointTo = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de créer des réseaux de cours d'eau interconnectés simplifiés ou non.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie des cours d'eau simple.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les réseaux de cours d'eau interconnectés.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterSensEcoulementCreerReseau(ByRef pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel,
                                                      Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface pour extraire les Nodes du Edge traité.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface pour extraire les Nodes.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries sélectionnées.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire le nombre de composantes.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la le réseau traité.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour simplifier la géométrie du réseau.

        Try
            'Définir la géométrie par défaut
            TraiterSensEcoulementCreerReseau = New GeometryBag
            TraiterSensEcoulementCreerReseau.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSensEcoulementCreerReseau, IGeometryCollection)

            'Si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire tous les Nodes de la topologie
                pEnumTopoNode = pTopologyGraph.Nodes

                'Extraire le premier Node à traiter
                pEnumTopoNode.Reset()
                pTopoNode = pEnumTopoNode.Next()

                'Afficher la barre de progression
                InitBarreProgression(0, pEnumTopoNode.Count, pTrackCancel)

                'Vider la sélection des Edges
                pTopologyGraph.SetSelectionEmpty(esriTopologyElementType.esriTopologyEdge)

                'Traiter tous les nodes
                Do Until pTopoNode Is Nothing
                    'Vérifier si le Node a été traité
                    If Not pTopoNode.IsSelected Then
                        'Interface pour vérifier si le node est une extrémité
                        pEnumNodeEdge = pTopoNode.Edges(True)

                        'Vérifier si le node est une extrémité
                        If pEnumNodeEdge.Count = 1 Then
                            'Créer une nouvelle polyline vide contrenant le réseau à trouver
                            pPolyline = New Polyline
                            pPolyline.SpatialReference = TraiterSensEcoulementCreerReseau.SpatialReference

                            'Créer le réseau à partir d'un node de début ou de fin
                            CreerReseau(pPolyline, pTopologyGraph, pTopoNode, bEnleverSelection)

                            'Ajouter la géométrie trouvée
                            If pPolyline.IsEmpty = False Then pGeomSelColl.AddGeometry(pPolyline)

                            'Vérifier si on doit simplifier le réseau de cours d'eau
                            If gsExpression = "VRAI" Then
                                'Interface utilisé pour simplifier la géométrie du réseau
                                pTopoOp = CType(pPolyline, ITopologicalOperator2)

                                'Indiquer que la géométrie n'est pas simple
                                pTopoOp.IsKnownSimple_2 = False

                                'Simplifier la géométrie
                                pTopoOp.Simplify()
                            End If

                            'Interface pour extraire le nombre de composantes du réseau
                            pGeomColl = CType(pPolyline, IGeometryCollection)

                            'Écrire une erreur de façon à creer la géométrie du réseau
                            EcrireFeatureErreur("Réseau:" & pGeomSelColl.GeometryCount.ToString _
                                                & " /NbNoeuds=" & pEnumTopoNode.Count.ToString & ", NbComposantes=" & pGeomColl.GeometryCount.ToString _
                                                & " /" & gsNomAttribut & " " & gsExpression & " " & gsInformation, pPolyline, pGeomColl.GeometryCount)
                        End If
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain Node à traiter
                    pTopoNode = pEnumTopoNode.Next()
                Loop
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pEnumNodeEdge = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pFeatureSel = Nothing
            pGeomColl = Nothing
            pGeomSelColl = Nothing
            pPolyline = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer tout le réseau rattaché à un Node de départ sélectionné.
    ''' 
    ''' Un réseau est un ensemble de lignes interconnecté entre eux.
    '''</summary>
    '''
    '''<param name="pPolyline"> Interface ESRI contenant le réseau trouvé.</param>
    '''<param name="pTopologyGraph"> Interface ESRI contenant la topologie des éléments visibles.</param>
    '''<param name="pTopoNode"> Interface ESRI contenant le Node de la topologie utilisé pour rechercher le réseau.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    ''' 
    Public Sub CreerReseau(ByRef pPolyline As IPolyline, ByRef pTopologyGraph As ITopologyGraph4, ByRef pTopoNode As ITopologyNode, _
                           Optional ByVal bEnleverSelection As Boolean = True)
        'Déclarer les variables de travail
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface pour extraire les Nodes du Edge traité.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le Edge traité.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter la géométrie du Edge dans la Poilyline.

        Try
            'Vérifier si le Node a été traité
            If Not pTopoNode.IsSelected Then
                'Interface pour ajouter les lignes trouvées
                pGeomColl = CType(pPolyline, IGeometryCollection)

                'Indiquer que le Node est traité
                pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoNode)

                'Interface pour extraire tous les edges du Node
                pEnumNodeEdge = pTopoNode.Edges(True)

                'Extraire le premier edge du Node
                pEnumNodeEdge.Reset()
                pEnumNodeEdge.Next(pTopoEdge, True)

                'Traiter tous les Edges du Node
                Do Until pTopoEdge Is Nothing
                    'Vérifier si le Edge a été traité
                    If Not pTopoEdge.IsSelected Then
                        'Ajouter le Edge dans la Polyline
                        pGeomColl.AddGeometryCollection(CType(pTopoEdge.Geometry, IGeometryCollection))

                        'Indiquer que le Node est traité
                        pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoEdge)

                        'Vérifier si le node traité est le même que celui de début du edge
                        If pTopoNode.Equals(pTopoEdge.FromNode) Then
                            'Sélectionner le réseau par le ToNode
                            CreerReseau(pPolyline, pTopologyGraph, pTopoEdge.ToNode)

                            'Si le node traité est le même que celui de fin du edge
                        Else
                            'Sélectionner le réseau par le FromNode
                            CreerReseau(pPolyline, pTopologyGraph, pTopoEdge.FromNode)
                        End If
                    End If

                    'Extraire le prochain edge du Node
                    pEnumNodeEdge.Next(pTopoEdge, True)
                Loop
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pEnumNodeEdge = Nothing
            pTopoEdge = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de sélectionner les éléments du FeatureLayer dont la géométrie respecte ou non le sens d'écoulement spécifié 
    ''' selon le nombre de First (au moins 1) et de Last (au moins 1) pour les noeuds qui possèdent au moins 2 edges.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie des cours d'eau simple.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterSensEcoulementNoeud(ByRef pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel,
                                                     Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface pour extraire les Nodes du Edge traité.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface pour extraire les Nodes.
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node à traiter.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le Edge traité.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries sélectionnées.
        Dim nbFirst As Integer = 0              'Contient le nombre de First à un Noeud.
        Dim nbLast As Integer = 0               'Contient le nombre de Last à un Noeud.
        Dim sMessage As String = ""             'Contient le message d'erreur.
        Dim sOID As String = ""                 'Contient le OIDs des éléments en erreur.
        Dim bSucces As Boolean = False          'Indique si le sens recherché est un succès.

        Try
            'Définir la géométrie par défaut
            TraiterSensEcoulementNoeud = New GeometryBag
            TraiterSensEcoulementNoeud.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSensEcoulementNoeud, IGeometryCollection)

            'Si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Extraire tous les Nodes de la topologie
                pEnumTopoNode = pTopologyGraph.Nodes

                'Extraire le premier Node à traiter
                pEnumTopoNode.Reset()
                pTopoNode = pEnumTopoNode.Next()

                'Afficher la barre de progression
                InitBarreProgression(0, pEnumTopoNode.Count, pTrackCancel)

                'Vider la sélection des Edges
                pTopologyGraph.SetSelectionEmpty(esriTopologyElementType.esriTopologyEdge)

                'Traiter tous les nodes
                Do Until pTopoNode Is Nothing
                    'Par défaut Indiquer que le sens d'écoulement est valide
                    bSucces = True
                    sMessage = "#Le sens d'écoulement est valide"

                    'Interface pour vérifier si le node est une extrémité
                    pEnumNodeEdge = pTopoNode.Edges(True)

                    'Vérifier si le node n'est pas une extrémité
                    If pEnumNodeEdge.Count > 1 Then
                        'Initialiser les compteurs
                        nbFirst = 0
                        nbLast = 0
                        sOID = ""

                        'Extraire le premier edge du Node
                        pEnumNodeEdge.Reset()
                        pEnumNodeEdge.Next(pTopoEdge, True)

                        'Traiter tous les Edges du Node
                        Do Until pTopoEdge Is Nothing
                            'Vérifier si c'est un First
                            If pTopoNode.Equals(pTopoEdge.FromNode) Then
                                'Compter le nombre de First
                                nbFirst = nbFirst + 1

                                'Si c'est un Last
                            Else
                                'Compter le nombre de Last
                                nbLast = nbLast + 1
                            End If

                            'Extraire le prochain edge du Node
                            pEnumNodeEdge.Next(pTopoEdge, True)
                        Loop

                        'Vérifier si le sens d'écoulement est valide pour le Noeud traité
                        If nbFirst = 0 Or nbLast = 0 Then
                            'Indiquer que le sens d'écoulement est invalide
                            bSucces = False
                            sMessage = "#Le sens d'écoulement est invalide"
                        End If

                        'Vérifier si on doit sélectionner l'élément
                        If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                            'Ajouter le Node dans la Géométrie en erreur
                            pGeomSelColl.AddGeometry(pTopoNode.Geometry)

                            'Contient les parents sélectionnés
                            pEnumTopoParent = pTopoNode.Parents
                            'Extraire le premier élément parent
                            pEnumTopoParent.Reset()
                            pEsriTopoParent = pEnumTopoParent.Next()
                            'Lire tous les éléments
                            Do Until pEsriTopoParent.m_pFC Is Nothing
                                'Ajouter le OID dans la sélection
                                pFeatureSel.SelectionSet.Add(pEsriTopoParent.m_FID)
                                'Conserver le OID
                                sOID = sOID & pEsriTopoParent.m_FID.ToString & ","
                                'Extraire le prochain élément
                                pEsriTopoParent = pEnumTopoParent.Next()
                            Loop
                            'Enlever la dernière virgule
                            If sOID.Length > 0 Then sOID = sOID.Substring(0, sOID.Length - 1)

                            'Écrire une erreur de façon à creer la géométrie du réseau
                            EcrireFeatureErreur("OID=" & sOID & sMessage & " #NbFirst=" & nbFirst.ToString _
                                                & ", NbLast=" & nbLast.ToString & ", NbEdges=" & pEnumNodeEdge.Count.ToString _
                                                & " /" & gsNomAttribut & " " & gsExpression & " " & gsInformation, pTopoNode.Geometry, pEnumNodeEdge.Count)
                        End If
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain Node à traiter
                    pTopoNode = pEnumTopoNode.Next()
                Loop
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pEnumNodeEdge = Nothing
            pEnumTopoNode = Nothing
            pEnumTopoParent = Nothing
            pEsriTopoParent = Nothing
            pTopoNode = Nothing
            pTopoEdge = Nothing
            pFeatureSel = Nothing
            pGeomSelColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de sélectionner les éléments du FeatureLayer dont les lignes de cours d'eau spécifiées respectent ou non le sens d'écoulement  
    ''' selon les valeurs d'élévation d'un MNE en relation.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pFeatureCursor"> Interface utilisé pour extraire les éléments sélectionnés.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterSensEcoulementLigne(ByRef pFeatureCursor As IFeatureCursor, ByRef pTrackCancel As ITrackCancel,
                                                Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface ESRI contenant les géométries sélectionnées.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la ligne des cours d'Eau simple.
        Dim pRasterLayer As IRasterLayer = Nothing          'Interface contenant le Layer du MNE.
        Dim pRaster2 As IRaster2 = Nothing                  'Interface contenant le MNE pour extraire l'élévation.
        Dim pGeodataset As IGeoDataset = Nothing            'Interface contenant la référence spatiale du MNE.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets du cours d'eau.
        Dim iCol As Integer = Nothing       'Contient la rangée correspondant à la valeur X dans le MNE.
        Dim iRow As Integer = Nothing       'Contient la colonne correspondant à la valeur Y dans le MNE.
        Dim dFrom As Double = Nothing       'Contient la valeur d'élévation du point de début d'un cours d'eau.
        Dim dTo As Double = Nothing         'Contient la valeur d'élévation du point de fin d'un cours d'eau.
        Dim iDesc As Integer = 0            'Contient le nombre de droite dont l'élévation est descendente.
        Dim iAsc As Integer = 0             'Contient le nombre de droite dont l'élévation est ascendente.
        Dim iEgal As Integer = 0            'Contient le nombre de droite dont l'élévation est égale.
        Dim bSucces As Boolean = False      'Indique si le sens recherché est un succès.

        Try
            'Définir la géométrie par défaut
            TraiterSensEcoulementLigne = New GeometryBag
            TraiterSensEcoulementLigne.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference
            'Interface contenant les géométries des éléments sélectionnés
            pGeomSelColl = CType(TraiterSensEcoulementLigne, IGeometryCollection)

            'Définir le RasterLayer du MNE
            pRasterLayer = CType(gpRasterLayersRelation.Item(1), IRasterLayer)
            'Interface contenant le MNE pour extraire l'élévation.
            pRaster2 = CType(pRasterLayer.Raster, IRaster2)
            'Interface pour extraire la référence spatiale du MNE
            pGeodataset = CType(pRaster2, IGeoDataset)

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline And pRasterLayer IsNot Nothing Then
                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature()

                'Traiter tous les éléments sélectionnés du FeatureLayer
                Do While Not pFeature Is Nothing
                    'Interface pour projeter
                    pPolyline = CType(pFeature.ShapeCopy, IPolyline)
                    pPolyline.Project(pGeodataset.SpatialReference)

                    'Extraire la position de la rangée à partir de la valeur X
                    iCol = pRaster2.ToPixelColumn(pPolyline.FromPoint.X)
                    'Extraire la position de la colonne à partir de la valeur Y
                    iRow = pRaster2.ToPixelRow(pPolyline.FromPoint.Y)
                    'Extraire la valeur d'élévation du MNE
                    dFrom = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))

                    'Initialiser les compteurs
                    iDesc = 0
                    iAsc = 0
                    iEgal = 0

                    'Interface pour extraire les point du cours d'eau
                    pPointColl = CType(pPolyline, IPointCollection)

                    'Traiter tous les points du cours d'eau
                    For i = 1 To pPointColl.PointCount - 1
                        'Extraire la position de la rangée à partir de la valeur X
                        iCol = pRaster2.ToPixelColumn(pPointColl.Point(i).X)
                        'Extraire la position de la colonne à partir de la valeur Y
                        iRow = pRaster2.ToPixelRow(pPointColl.Point(i).Y)
                        'Extraire la valeur d'élévation du MNE
                        dTo = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))

                        'Compter le nombre de descendant
                        If dFrom > dTo Then
                            iDesc = iDesc + 1

                            'Compter le nombre d'ascendant
                        ElseIf dTo > dFrom Then
                            iAsc = iAsc + 1

                            'Compter le nombre d'égal
                        Else
                            iEgal = iEgal + 1
                        End If

                        'Le dernier devient le premier
                        dFrom = dTo
                    Next

                    'Extraire la position de la rangée à partir de la valeur X
                    iCol = pRaster2.ToPixelColumn(pPolyline.FromPoint.X)
                    'Extraire la position de la colonne à partir de la valeur Y
                    iRow = pRaster2.ToPixelRow(pPolyline.FromPoint.Y)
                    'Extraire la valeur d'élévation du MNE
                    dFrom = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))

                    'Vérifier si l'état recherché est un succès, si non ascendant
                    bSucces = Not (dTo - dFrom > gdTolEgaliteZ) = (gsExpression = "VRAI")

                    'Vérifier si on doit sélectionner l'élément
                    If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                        'Ajouter l'élément dans la sélection
                        pFeatureSel.Add(pFeature)
                        'Ajouter l'enveloppe de l'élément sélectionné
                        pGeomSelColl.AddGeometry(pPolyline)
                        'Écrire une erreur
                        EcrireFeatureErreur("OID=" & pFeature.OID.ToString & " #Sens descendant:" & bSucces.ToString _
                                            & " /ZDébut:" & dFrom.ToString & "+" & gsInformation & " >= ZFin:" & dTo.ToString & " = " & bSucces.ToString _
                                            & " /NbDesc:" & iDesc.ToString & " + NbAsc:" & iAsc.ToString _
                                            & " + NbEgale:" & iEgal.ToString & " = NbDroites:" & (pPointColl.PointCount - 1).ToString _
                                            & " /" & gsNomAttribut & " " & gsExpression & " " & gsInformation, pPolyline, CSng(dTo - dFrom))
                    End If

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
            pFeatureSel = Nothing
            pFeature = Nothing
            pGeomSelColl = Nothing
            pPolyline = Nothing
            pRasterLayer = Nothing
            pRaster2 = Nothing
            pGeodataset = Nothing
            pPointColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de sélectionner les éléments du FeatureLayer dont les polylignes de cours d'eau spécifiées respectent ou non le sens d'écoulement
    ''' selon les valeurs d'élévation d'un MNE en relation.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie des cours d'eau simple.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterSensEcoulementPolyligne(ByRef pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel,
                                                    Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface pour extraire les Nodes du Edge traité.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface pour extraire les Nodes.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node à traiter.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pRasterLayer As IRasterLayer = Nothing          'Interface contenant le Layer du MNE.
        Dim pRaster2 As IRaster2 = Nothing                  'Interface contenant le MNE.
        Dim pGeodataset As IGeoDataset = Nothing            'Interface contenant la référence spatiale du MNE.
        Dim pPointFrom As IPoint = Nothing                  'Interface contenant le point de début.
        Dim pSetEdge As ISet = Nothing                      'Interface contenant l'ensemble des Edges non inversés d'une polyligne continue.
        Dim pSetEdgeInv As ISet = Nothing                   'Interface contenant l'ensemble des Edges inversés d'une polyligne continue.
        Dim iCol As Integer = Nothing       'Contient la rangée correspondant à la valeur X dans le MNE.
        Dim iRow As Integer = Nothing       'Contient la colonne correspondant à la valeur Y dans le MNE.
        Dim dFromDepart As Double = Nothing 'Contient la valeur d'élévation du point de départ d'une polyligne continue.
        Dim dFrom As Double = Nothing       'Contient la valeur d'élévation du point de début d'un cours d'eau.
        Dim iNbDesc As Integer = 0          'Contient le nombre de ligne dont l'élévation est descendente.
        Dim iNbAsc As Integer = 0           'Contient le nombre de ligne dont l'élévation est ascendente.
        Dim iNbEgal As Integer = 0          'Contient le nombre de ligne dont l'élévation est égale.
        Dim sElevation As String = ""       'Contient les élévations trouvées.
        Dim bDescendant As Boolean = True   'Indique si le sens est descendant.
        Try
            'Définir la géométrie par défaut
            TraiterSensEcoulementPolyligne = New GeometryBag
            TraiterSensEcoulementPolyligne.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir le RasterLayer du MNE
            pRasterLayer = CType(gpRasterLayersRelation.Item(1), IRasterLayer)
            'Interface contenant le MNE pour extraire l'élévation.
            pRaster2 = CType(pRasterLayer.Raster, IRaster2)
            'Interface pour extraire la référence spatiale du MNE
            pGeodataset = CType(pRaster2, IGeoDataset)

            'Interface pour extraire et sélectionner les éléments
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)

            'Si la géométrie est de type Polyline
            If gpFeatureLayerSelection.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                'Vider la sélection des Edges
                pTopologyGraph.SetSelectionEmpty(esriTopologyElementType.esriTopologyEdge)

                'Extraire tous les Nodes de la topologie
                pEnumTopoNode = pTopologyGraph.Nodes

                'Afficher la barre de progression
                InitBarreProgression(0, pEnumTopoNode.Count, pTrackCancel)

                'Extraire le premier Node à traiter
                pEnumTopoNode.Reset()
                pTopoNode = pEnumTopoNode.Next()

                'Traiter tous les nodes
                Do Until pTopoNode Is Nothing
                    'Vérifier si le Node a été traité
                    If Not pTopoNode.IsSelected Then
                        'Interface pour vérifier si le node est une extrémité
                        pEnumNodeEdge = pTopoNode.Edges(True)

                        'Vérifier si le node n'est pas une continuité
                        If pEnumNodeEdge.Count <> 2 Then
                            'Initialiser les compteurs
                            iNbDesc = 0
                            iNbAsc = 0
                            iNbEgal = 0

                            'Initialisation des ensembles de travail contenant les edges de la polygne continue
                            pSetEdge = New ESRI.ArcGIS.esriSystem.Set
                            pSetEdgeInv = New ESRI.ArcGIS.esriSystem.Set

                            'Définir le sommet de début
                            pPointFrom = CType(pTopoNode.Geometry, IPoint)
                            'Projet le point
                            pPointFrom.Project(pGeodataset.SpatialReference)
                            'Extraire la position de la rangée à partir de la valeur X
                            iCol = pRaster2.ToPixelColumn(pPointFrom.X)
                            'Extraire la position de la colonne à partir de la valeur Y
                            iRow = pRaster2.ToPixelRow(pPointFrom.Y)
                            'Extraire la valeur d'élévation du MNE
                            dFromDepart = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))
                            dFrom = dFromDepart

                            'Sélectionner dans les polylignes continues les lignes inversées
                            PolyligneContinue(pRasterLayer, dFromDepart, dFrom, pFeatureSel, TraiterSensEcoulementPolyligne, pSetEdge, pSetEdgeInv, _
                                              iNbDesc, iNbAsc, iNbEgal, sElevation, pTopologyGraph, pTopoNode, bDescendant, bEnleverSelection)

                            'Indiquer que le Node est traité
                            pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoNode)
                        End If
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain Node à traiter
                    pTopoNode = pEnumTopoNode.Next()
                Loop
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pEnumNodeEdge = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pFeatureSel = Nothing
            pRasterLayer = Nothing
            pRaster2 = Nothing
            pGeodataset = Nothing
            pPointFrom = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de valider la continuité des lignes d'un réseau à partir de la topologie et des élévations d'un MNE. 
    ''' Les éléments qui respectent ou non la continuité des polylignes spécifiées seront sélectionnés.
    ''' 
    ''' Un réseau est un ensemble de lignes interconnectées entre eux.
    '''</summary>
    '''
    '''<param name="pRasterLayer"> Interface contenant le Layer du MNE.</param>
    '''<param name="dFromDepart"> Contient la valeur d'élévation du point de début d'un cours d'eau continu.</param>
    '''<param name="dFrom"> Contient la valeur d'élévation du point de début d'un segment de cours d'eau.</param>
    '''<param name="pFeatureSel"> Interface ESRI utilisé pour sélectionner les éléments trouvés.</param>
    '''<param name="pGeometryBag"> Interface contenant les géométries en erreur.</param>
    '''<param name="pSetEdge"> Interface contenant les lignes non inversées d'un cours d'eau continu.</param>
    '''<param name="pSetEdgeInv"> Interface contenant les lignes inversées d'un cours d'eau continu.</param>
    '''<param name="iNbDesc"> Contient le nombre de ligne dont l'élévation est descendente.</param>
    '''<param name="iNbAsc"> Contient le nombre de ligne dont l'élévation est ascendente.</param>
    '''<param name="iNbEgal"> Contient le nombre de ligne dont l'élévation est égale.</param>
    '''<param name="sElevation"> Contient les élévations trouvées.</param>
    '''<param name="pTopologyGraph"> Interface ESRI contenant la topologie des éléments visibles.</param>
    '''<param name="pTopoNodeFrom"> Interface ESRI contenant le Node de début d'une ligne.</param>
    '''<param name="bDescendant"> Permettre d'indiquer si le sens d'écoulement est Descendant=True ou Ascendant=False.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    ''' 
    Public Sub PolyligneContinue(ByRef pRasterLayer As IRasterLayer, ByRef dFromDepart As Double, ByRef dFrom As Double, ByRef pFeatureSel As IFeatureSelection, _
                                 ByRef pGeometryBag As IGeometry, ByRef pSetEdge As ISet, ByRef pSetEdgeInv As ISet, _
                                 ByRef iNbDesc As Integer, ByRef iNbAsc As Integer, ByRef iNbEgal As Integer, _
                                 ByRef sElevation As String, ByRef pTopologyGraph As ITopologyGraph4, ByRef pTopoNodeFrom As ITopologyNode, _
                                 ByVal bDescendant As Boolean, Optional ByVal bEnleverSelection As Boolean = True)
        'Déclarer les variables de travail
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing    'Interface pour extraire les Nodes du Edge traité.
        Dim pTopoNodeTo As ITopologyNode = Nothing      'Interface contenant le noeud de fin.
        Dim pTopoEdge As ITopologyEdge = Nothing        'Interface contenant le Edge traité.
        Dim dTo As Double = Nothing                     'Contient la valeur d'élévation du point de fin d'un cours d'eau.

        Try
            'Vérifier si le Node a été traité
            If Not pTopoNodeFrom.IsSelected Then
                'Interface pour extraire tous les edges du Node
                pEnumNodeEdge = pTopoNodeFrom.Edges(True)

                'Si ce n'est pas la fin d'une polyligne continue
                If (pSetEdge.Count = 0 And pSetEdgeInv.Count = 0) Or pEnumNodeEdge.Count = 2 Then
                    'Extraire le premier edge du Node
                    pEnumNodeEdge.Reset()
                    pEnumNodeEdge.Next(pTopoEdge, True)

                    'Traiter tous les Edges du Node
                    Do Until pTopoEdge Is Nothing
                        'Vérifier si le Edge a été traité
                        If Not pTopoEdge.IsSelected Then
                            'Indiquer que le Edge est traité
                            pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoEdge)

                            'Vérifier si on traite une nouvelle polyligne à partir du noeud de départ
                            If pSetEdge.Count = 0 And pSetEdgeInv.Count = 0 Then
                                'On réinitialise la valeur Z de départ
                                dFrom = dFromDepart
                            End If

                            'Vérifier si le node traité est le même que celui de début du edge (Sens descendant)
                            If pTopoNodeFrom.Equals(pTopoEdge.FromNode) Then
                                'Traiter le sens Descendant de la polyligne continue
                                PolyligneContinueDescendant(pRasterLayer, pEnumNodeEdge, pTopoEdge, pTopoNodeTo, dFrom, dTo, pSetEdge, pSetEdgeInv, sElevation, bDescendant)

                                'Si le node traité est le même que celui de fin du edge (sens ascendant)
                            Else
                                'Traiter le sens Ascendant de la polyligne continue
                                PolyligneContinueAscendant(pRasterLayer, pEnumNodeEdge, pTopoEdge, pTopoNodeTo, dFrom, dTo, pSetEdge, pSetEdgeInv, sElevation, bDescendant)
                            End If

                            'Compter le nombre de descendant
                            If dFrom > dTo Then
                                iNbDesc = iNbDesc + 1

                                'Compter le nombre d'ascendant
                            ElseIf dTo > dFrom Then
                                iNbAsc = iNbAsc + 1

                                'Compter le nombre d'égal
                            Else
                                iNbEgal = iNbEgal + 1
                            End If

                            'Définir l'élévation de début à partir de celui de fin
                            dFrom = dTo

                            'Vérifier la continuité du réseau par le ToNode
                            PolyligneContinue(pRasterLayer, dFromDepart, dFrom, pFeatureSel, pGeometryBag, pSetEdge, pSetEdgeInv, _
                                              iNbDesc, iNbAsc, iNbEgal, sElevation, pTopologyGraph, pTopoNodeTo, bDescendant, bEnleverSelection)
                        End If

                        'Extraire le prochain edge du Node
                        pEnumNodeEdge.Next(pTopoEdge, True)
                    Loop

                    'Indiquer que le Node est traité
                    pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoNodeFrom)

                    'Si c'est une fin de polyligne continue
                Else
                    'Définir la valeur d'élévation de fin
                    dTo = dFrom

                    'Traiter la fin d'une polyligne continue
                    PolyligneContinueFin(dFromDepart, dFrom, pFeatureSel, pGeometryBag, pSetEdge, pSetEdgeInv, _
                                         iNbDesc, iNbAsc, iNbEgal, sElevation, bDescendant, bEnleverSelection)
                End If

                'Si le noeud est déjà traité
            Else
                'Afficher le message de noeud déjà traité
                Debug.Print("Erreur : Noeud déja traité!")
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pEnumNodeEdge = Nothing
            pTopoNodeTo = Nothing
            pTopoEdge = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de valider la continuité des lignes descendantes d'un réseau à partir de la topologie et des élévations d'un MNE. 
    ''' 
    ''' Un réseau est un ensemble de lignes interconnectées entre eux.
    '''</summary>
    '''
    '''<param name="pRasterLayer"> Interface contenant le Layer du MNE.</param>
    '''<param name="pEnumNodeEdge"> Interface pour extraire les Nodes du Edge traité.</param>
    '''<param name="pTopoEdge"> Interface contenant le Edge traité.</param>
    '''<param name="pTopoNodeTo"> Interface contenant le noeud de fin.</param>
    '''<param name="dFrom"> Contient la valeur d'élévation du point de début d'un segment de cours d'eau.</param>
    '''<param name="dTo"> Contient la valeur d'élévation du point de fin d'un segment de cours d'eau.</param>
    '''<param name="pSetEdge"> Interface contenant les lignes non inversées d'un cours d'eau continu.</param>
    '''<param name="pSetEdgeInv"> Interface contenant les lignes inversées d'un cours d'eau continu.</param>
    '''<param name="sElevation"> Contient les élévations trouvées.</param>
    '''<param name="bDescendant"> Permettre d'indiquer si le sens d'écoulement est Descendant=True ou Ascendant=False.</param>
    ''' 
    Public Sub PolyligneContinueDescendant(ByRef pRasterLayer As IRasterLayer, ByRef pEnumNodeEdge As IEnumNodeEdge, ByRef pTopoEdge As ITopologyEdge, _
                                           ByRef pTopoNodeTo As ITopologyNode, ByRef dFrom As Double, ByRef dTo As Double, ByRef pSetEdge As ISet, ByRef pSetEdgeInv As ISet, _
                                           ByRef sElevation As String, ByRef bDescendant As Boolean)
        'Déclarer les variables de travail
        Dim pRaster2 As IRaster2 = Nothing          'Interface contenant le MNE.
        Dim pGeodataset As IGeoDataset = Nothing    'Interface contenant la référence spatiale du MNE.
        Dim pPointTo As IPoint = Nothing            'Interface contenant le point de fin.
        Dim iCol As Integer = Nothing               'Contient la rangée correspondant à la valeur X dans le MNE.
        Dim iRow As Integer = Nothing               'Contient la colonne correspondant à la valeur Y dans le MNE.

        Try
            'Interface contenant le MNE pour extraire l'élévation.
            pRaster2 = CType(pRasterLayer.Raster, IRaster2)
            'Interface pour extraire la référence spatiale du MNE
            pGeodataset = CType(pRaster2, IGeoDataset)

            'Définir le node de fin
            pTopoNodeTo = pTopoEdge.ToNode
            'Définir le sommet de fin
            pPointTo = CType(pTopoNodeTo.Geometry, IPoint)
            'Projet le point
            pPointTo.Project(pGeodataset.SpatialReference)
            'Extraire la position de la rangée à partir de la valeur X
            iCol = pRaster2.ToPixelColumn(pPointTo.X)
            'Extraire la position de la colonne à partir de la valeur Y
            iRow = pRaster2.ToPixelRow(pPointTo.Y)
            'Extraire la valeur d'élévation du MNE
            dTo = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))

            'Si la polyligne est vide, on débute la recherche de la polyligne continue
            If (pSetEdge.Count = 0 And pSetEdgeInv.Count = 0) Then
                'Conserver les élévations de début
                sElevation = dFrom.ToString("F2") & " >= " & dTo.ToString("F2")

                'Indiquer que le sens de la polyligne est descendant
                bDescendant = True

                'Si la polyligne est encore continue
            ElseIf pEnumNodeEdge.Count = 2 Then
                'Vérifier si le sens est Descendant
                If bDescendant Then
                    'Conserver l'élévation de fin
                    sElevation = sElevation & " >= " & dTo.ToString("F2")
                Else
                    'Conserver l'élévation de fin
                    sElevation = sElevation & " <= " & dTo.ToString("F2")
                End If
            End If

            'Vérifier si le sens n'est pas respecté
            If bDescendant = False Then
                'Ajouter le Edge dont le sens est inversé
                pSetEdgeInv.Add(pTopoEdge)
            Else
                'Ajouter le Edge dont le sens n'est pas inversé
                pSetEdge.Add(pTopoEdge)
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pRaster2 = Nothing
            pGeodataset = Nothing
            pPointTo = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de valider la continuité des lignes ascendantes d'un réseau à partir de la topologie et des élévations d'un MNE. 
    ''' 
    ''' Un réseau est un ensemble de lignes interconnectées entre eux.
    '''</summary>
    '''
    '''<param name="pRasterLayer"> Interface contenant le Layer du MNE.</param>
    '''<param name="pEnumNodeEdge"> Interface pour extraire les Nodes du Edge traité.</param>
    '''<param name="pTopoEdge"> Interface contenant le Edge traité.</param>
    '''<param name="pTopoNodeTo"> Interface contenant le noeud de fin.</param>
    '''<param name="dFrom"> Contient la valeur d'élévation du point de début d'un segment de cours d'eau.</param>
    '''<param name="dTo"> Contient la valeur d'élévation du point de fin d'un segment de cours d'eau.</param>
    '''<param name="pSetEdge"> Interface contenant les lignes non inversées d'un cours d'eau continu.</param>
    '''<param name="pSetEdgeInv"> Interface contenant les lignes inversées d'un cours d'eau continu.</param>
    '''<param name="sElevation"> Contient les élévations trouvées.</param>
    '''<param name="bDescendant"> Permettre d'indiquer si le sens d'écoulement est Descendant=True ou Ascendant=False.</param>
    ''' 
    Public Sub PolyligneContinueAscendant(ByRef pRasterLayer As IRasterLayer, ByRef pEnumNodeEdge As IEnumNodeEdge, ByRef pTopoEdge As ITopologyEdge, _
                                          ByRef pTopoNodeTo As ITopologyNode, ByRef dFrom As Double, ByRef dTo As Double, ByRef pSetEdge As ISet, ByRef pSetEdgeInv As ISet, _
                                          ByRef sElevation As String, ByRef bDescendant As Boolean)
        'Déclarer les variables de travail
        Dim pRaster2 As IRaster2 = Nothing          'Interface contenant le MNE.
        Dim pGeodataset As IGeoDataset = Nothing    'Interface contenant la référence spatiale du MNE.
        Dim pPointTo As IPoint = Nothing            'Interface contenant le point de fin.
        Dim iCol As Integer = Nothing               'Contient la rangée correspondant à la valeur X dans le MNE.
        Dim iRow As Integer = Nothing               'Contient la colonne correspondant à la valeur Y dans le MNE.

        Try
            'Interface contenant le MNE pour extraire l'élévation.
            pRaster2 = CType(pRasterLayer.Raster, IRaster2)
            'Interface pour extraire la référence spatiale du MNE
            pGeodataset = CType(pRaster2, IGeoDataset)

            'Définir le node de fin
            pTopoNodeTo = pTopoEdge.FromNode
            'Définir le sommet de fin
            pPointTo = CType(pTopoNodeTo.Geometry, IPoint)
            'Projet le point
            pPointTo.Project(pGeodataset.SpatialReference)
            'Extraire la position de la rangée à partir de la valeur X
            iCol = pRaster2.ToPixelColumn(pPointTo.X)
            'Extraire la position de la colonne à partir de la valeur Y
            iRow = pRaster2.ToPixelRow(pPointTo.Y)
            'Extraire la valeur d'élévation du MNE
            dTo = CDbl(pRaster2.GetPixelValue(0, iCol, iRow))

            'Si la polyligne est vide, on débute la recherche de la polyligne continue
            If (pSetEdge.Count = 0 And pSetEdgeInv.Count = 0) Then
                'Conserver les élévations de début
                sElevation = dFrom.ToString("F2") & " <= " & dTo.ToString("F2")

                'Indiquer que le sens de la polyligne est ascendant
                bDescendant = False

                'Si la polyligne est encore continue
            ElseIf pEnumNodeEdge.Count = 2 Then
                'Vérifier si le sens est Descendant
                If bDescendant Then
                    'Conserver l'élévation de fin
                    sElevation = sElevation & " >= " & dTo.ToString("F2")
                Else
                    'Conserver l'élévation de fin
                    sElevation = sElevation & " <= " & dTo.ToString("F2")
                End If
            End If

            'Vérifier si le sens n'est pas respecté
            If bDescendant Then
                'Ajouter le Edge dont le sens est inversé
                pSetEdgeInv.Add(pTopoEdge)
            Else
                'Ajouter le Edge dont le sens n'est pas inversé
                pSetEdge.Add(pTopoEdge)
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pRaster2 = Nothing
            pGeodataset = Nothing
            pPointTo = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de valider la fin d'une polyligne continue d'un réseau à partir de la topologie et des élévations d'un MNE. 
    ''' Les éléments qui respectent ou non la continuité d'une polyligne spécifiée seront sélectionnés.
    ''' 
    ''' Un réseau est un ensemble de lignes interconnectées entre eux.
    '''</summary>
    '''
    '''<param name="dFromDepart"> Contient la valeur d'élévation du point de début d'un cours d'eau continu.</param>
    '''<param name="dTo"> Contient la valeur d'élévation du point de fin d'un segment de cours d'eau.</param>
    '''<param name="pFeatureSel"> Interface ESRI utilisé pour sélectionner les éléments trouvés.</param>
    '''<param name="pGeometryBag"> Interface contenant les géométries en erreur.</param>
    '''<param name="pSetEdge"> Interface contenant les lignes non inversées d'un cours d'eau continu.</param>
    '''<param name="pSetEdgeInv"> Interface contenant les lignes inversées d'un cours d'eau continu.</param>
    '''<param name="iNbDesc"> Contient le nombre de ligne dont l'élévation est descendente.</param>
    '''<param name="iNbAsc"> Contient le nombre de ligne dont l'élévation est ascendente.</param>
    '''<param name="iNbEgal"> Contient le nombre de ligne dont l'élévation est égale.</param>
    '''<param name="sElevation"> Contient les élévations trouvées.</param>
    '''<param name="bDescendant"> Permettre d'indiquer si le sens d'écoulement est Descendant=True ou Ascendant=False.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    ''' 
    Public Sub PolyligneContinueFin(ByRef dFromDepart As Double, ByRef dTo As Double, ByRef pFeatureSel As IFeatureSelection, _
                                    ByRef pGeometryBag As IGeometry, ByRef pSetEdge As ISet, ByRef pSetEdgeInv As ISet, _
                                    ByRef iNbDesc As Integer, ByRef iNbAsc As Integer, ByRef iNbEgal As Integer, ByRef sElevation As String, _
                                    ByVal bDescendant As Boolean, Optional ByVal bEnleverSelection As Boolean = True)
        'Déclarer les variables de travail
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le Edge traité.
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'utilisé pour ajouter les géométries en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter la géométrie du Edge dans la Poilyline.
        Dim pSetEdgeTmp As ISet = Nothing                   'Interface contenant l'ensemble des Edges temporaire.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne inversée de la polyligne continue.
        Dim pPolylineInv As IPolyline = Nothing             'Interface contenant toutes les lignes inversées de la polyligne continue.
        Dim sOid As String = ""             'Contient la liste des OIDs.
        Dim sMessage As String = ""         'Contient le message d'erreur.
        Dim bSucces As Boolean = False      'Indique si le sens recherché est un succès.

        Try
            'Vérifier si le sens est descendant
            If bDescendant Then
                'Par défaut le sens trouvé est un succès
                bSucces = True
                sMessage = " #Descendant : Sens d'écoulement valide"

                'Vérifier le sens de la polyligne
                If dTo - dFromDepart > gdTolEgaliteZ Then
                    'Indiquer que le sens d'écoulenent est inversé complètement
                    bSucces = False
                    sMessage = " #Descendant complet : Erreur de sens d'écoulement inversé"

                    'Inverser les Edges inversés et non inversés
                    pSetEdgeTmp = pSetEdgeInv
                    pSetEdgeInv = pSetEdge
                    pSetEdge = pSetEdgeTmp
                End If

                'Vérifier la présence de lignes inversées
                If pSetEdge.Count > 0 And pSetEdgeInv.Count > 0 Then
                    'Indiquer que le sens d'écoulenent est inversé partiellement
                    bSucces = False
                    sMessage = " #Descendant partiel : Erreur de sens d'écoulement inversé"
                End If

                'Si le sens est ascendant
            Else
                'Par défaut le sens trouvé est un succès
                bSucces = True
                sMessage = " #Ascendant : Sens d'écoulement valide"

                'Vérifier le sens de la polyligne est inversé
                If dFromDepart - dTo > gdTolEgaliteZ Then
                    'Indiquer que le sens d'écoulenent est inversé complètement
                    bSucces = False
                    sMessage = " #Ascendant complet : Erreur de sens d'écoulement inversé"

                    'Inverser les Edges inversés et non inversés
                    pSetEdgeTmp = pSetEdgeInv
                    pSetEdgeInv = pSetEdge
                    pSetEdge = pSetEdgeTmp
                End If

                'Vérifier la présence de lignes inversées
                If pSetEdge.Count > 0 And pSetEdgeInv.Count > 0 Then
                    'Indiquer que le sens d'écoulenent est inversé partiellement
                    bSucces = False
                    sMessage = " #Ascendant partiel : Erreur de sens d'écoulement inversé"
                End If
            End If

            'Vérifier si on doit sélectionner l'élément
            If (bSucces And Not bEnleverSelection) Or (Not bSucces And bEnleverSelection) Then
                'Si les erreurs sont les edges non-inversés 
                If bSucces Then
                    'On définit ceux non inversé en erreur
                    pSetEdgeInv = pSetEdge
                End If

                'Créer une nouvelle polyline inversée vide
                pPolylineInv = New Polyline
                pPolylineInv.SpatialReference = pGeometryBag.SpatialReference
                'Interface pour ajouter les lignes inversées de la polyligne continue
                pGeomColl = CType(pPolylineInv, IGeometryCollection)

                'Initialiser l'extraction
                pSetEdgeInv.Reset()
                'Extraire le premier Edge inversé
                pTopoEdge = CType(pSetEdgeInv.Next(), ITopologyEdge)
                'Traiter tous les Edges inversés
                Do Until pTopoEdge Is Nothing
                    'Contient les parents sélectionnés
                    pEnumTopoParent = pTopoEdge.Parents
                    'Extraire le premier élément parent
                    pEnumTopoParent.Reset()
                    pEsriTopoParent = pEnumTopoParent.Next()
                    'Lire tous les éléments
                    Do Until pEsriTopoParent.m_pFC Is Nothing
                        'Ajouter le OID dans la sélection
                        pFeatureSel.SelectionSet.Add(pEsriTopoParent.m_FID)
                        'Écrire une erreur
                        sOid = sOid & pEsriTopoParent.m_FID.ToString & ","
                        'Extraire le prochain élément
                        pEsriTopoParent = pEnumTopoParent.Next()
                    Loop

                    'Interface contenant la ligne inversée
                    pPolyline = CType(pTopoEdge.Geometry, IPolyline)
                    'Projeter la polyligne
                    pPolyline.Project(pGeometryBag.SpatialReference)
                    'Ajouter la ligne du Edge inversé
                    pGeomColl.AddGeometryCollection(CType(pPolyline, IGeometryCollection))

                    'Extraire le prochain Edge inversé
                    pTopoEdge = CType(pSetEdgeInv.Next(), ITopologyEdge)
                Loop

                'Interfac utilisé pour ajouter les lignes en erreur
                pGeomSelColl = CType(pGeometryBag, IGeometryCollection)

                'Ajouter les lignes trouvées
                pGeomSelColl.AddGeometry(pPolylineInv)

                'Si le sens est descendant
                If bDescendant Then
                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & sOid.Substring(0, sOid.Length - 1) & sMessage _
                                        & " /ZDébut:" & dFromDepart.ToString & "+" & gsInformation & " >= ZFin:" & dTo.ToString & " = " & bSucces.ToString _
                                        & " /Élévations=" & sElevation & " /nbDesc:" & iNbDesc.ToString & " + nbAsc:" & iNbAsc.ToString _
                                        & " + nbEgale:" & iNbEgal.ToString & " = NbComposantes:" & (pSetEdge.Count + pSetEdgeInv.Count).ToString _
                                        & " /" & gsNomAttribut & " " & gsExpression & " " & gsInformation, pPolylineInv, CSng(dTo - dFromDepart))

                    'Si le sens est ascendant
                Else
                    'Écrire une erreur
                    EcrireFeatureErreur("OID=" & sOid.Substring(0, sOid.Length - 1) & sMessage _
                                        & " /ZFin:" & dFromDepart.ToString & " <= ZDébut:" & dTo.ToString & "+" & gsInformation & " = " & bSucces.ToString _
                                        & " /Élévations=" & sElevation & " /nbDesc:" & iNbDesc.ToString & " + nbAsc:" & iNbAsc.ToString _
                                        & " + nbEgale:" & iNbEgal.ToString & " = NbComposantes:" & (pSetEdge.Count + pSetEdgeInv.Count).ToString _
                                        & " /" & gsNomAttribut & " " & gsExpression & " " & gsInformation, pPolylineInv, CSng(dTo - dFromDepart))
                End If
            End If

            'Initialiser les variables de travail
            iNbDesc = 0
            iNbAsc = 0
            iNbEgal = 0
            sElevation = ""

            'Initialisation des ensembles de travail contenant les edges de la polygne continue
            pSetEdge = New ESRI.ArcGIS.esriSystem.Set
            pSetEdgeInv = New ESRI.ArcGIS.esriSystem.Set

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pEnumTopoParent = Nothing
            pTopoEdge = Nothing
            pEsriTopoParent = Nothing
            pGeomSelColl = Nothing
            pGeomColl = Nothing
            pSetEdgeTmp = Nothing
            pPolyline = Nothing
            pPolylineInv = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de sélectionner les éléments du FeatureLayer dont les réseaux de cours d'eau spécifiées respectent ou non le sens d'écoulement ascendant
    ''' selon les fin de réseaux.
    ''' 
    ''' Seuls les éléments sélectionnés sont traités.
    ''' Si aucun élément n'est sélectionné, tous les éléments sont considérés sélectionnés.
    ''' 
    ''' bEnleverSelection=True : Les éléments qui respectent la contrainte sont enlevés dans la sélection.
    ''' bEnleverSelection=False : Les éléments qui respectent la contrainte sont conservés dans la sélection.
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie des cours d'eau simple.</param>
    '''<param name="pTrackCancel"> Permet d'annuler la sélection avec la touche ESC du clavier.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    '''
    '''<return>Les géométries des éléments sélectionnés.</return>
    ''' 
    '''</summary>
    '''
    Private Function TraiterSensEcoulementReseau(ByRef pTopologyGraph As ITopologyGraph4, ByRef pTrackCancel As ITrackCancel,
                                                 Optional ByVal bEnleverSelection As Boolean = True) As IGeometry
        'Déclarer les variables de travail
        Dim pFeatureLayerFin As IFeatureLayer = Nothing     'Interface contenant les fins de réseaux à traiter.
        Dim pFeatureSel As IFeatureSelection = Nothing      'Interface utilisé pour extraire et sélectionner les éléments du FeatureLayer.
        Dim pCursor As ICursor = Nothing                    'Interface utilisé pour extraire les donnéées à traiter.
        Dim pFeatureCursor As IFeatureCursor = Nothing      'Interface utilisé pour extraire les éléments à traiter.
        Dim pFeature As IFeature = Nothing                  'Interface contenant l'élément à traiter.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface pour extraire les Nodes.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node à traiter.

        Try
            'Définir la géométrie par défaut
            TraiterSensEcoulementReseau = New GeometryBag
            TraiterSensEcoulementReseau.SpatialReference = gpFeatureLayerSelection.AreaOfInterest.SpatialReference

            'Définir le FeatureLayer des fins de réseaux
            pFeatureLayerFin = CType(gpFeatureLayersRelation.Item(1), IFeatureLayer)

            'Interfaces pour extraire les éléments sélectionnés des fins de réseaux
            pFeatureSel = CType(pFeatureLayerFin, IFeatureSelection)

            'Vérifier si aucun fin de réseau n'est sélectionné
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionner tous les éléments
                pFeatureSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If

            'Afficher la barre de progression
            InitBarreProgression(0, pFeatureSel.SelectionSet.Count, pTrackCancel)

            'Extraire les éléments sélectionnés
            pFeatureSel.SelectionSet.Search(Nothing, True, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)

            'Extraire le premier fin de réseau
            pFeature = pFeatureCursor.NextFeature

            'Interface contenant les éléments sélectionnés du FeatureLayer de sélection
            pFeatureSel = CType(gpFeatureLayerSelection, IFeatureSelection)
            'Vider la sélection
            pFeatureSel.Clear()

            'Traiter tous les fins de réseaux
            Do Until pFeature Is Nothing
                'Vider la sélection des Nodes
                'Sélectionner le noeud correspondant au fin de réseau traité
                pTopologyGraph.SelectByGeometry(esriTopologyElementType.esriTopologyNode, esriTopologySelectionResultEnum.esriTopologySelectionResultNew, pFeature.Shape)

                'Interface pour extraire les Noeuds sélectionnés
                pEnumTopoNode = pTopologyGraph.NodeSelection

                'Vérifier si un noeud a été sélectionné
                If pEnumTopoNode.Count > 0 Then
                    'Extraire le premier Noeud
                    pEnumTopoNode.Reset()
                    pTopoNode = pEnumTopoNode.Next

                    'Vider la sélection des Nodes
                    pTopologyGraph.SetSelectionEmpty(esriTopologyElementType.esriTopologyNode)

                    'Valider le sens d'élement des lignes du réseau de cours d'eau
                    ValiderReseau(pFeatureSel, CType(TraiterSensEcoulementReseau, IGeometryBag), pTopologyGraph, pTopoNode, bEnleverSelection)
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le prochain fin de réseau
                pFeature = pFeatureCursor.NextFeature
            Loop

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFeatureLayerFin = Nothing
            pFeatureSel = Nothing
            pCursor = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de valider le sens d'écoulement de toutes les lignes qui composent un réseau rattaché à un Node de départ sélectionné.
    ''' 
    ''' Un réseau est un ensemble de lignes interconnecté entre eux.
    '''</summary>
    '''
    '''<param name="pFeatureSel"> Interface ESRI contenant la sélection des éléments.</param>
    '''<param name="pGeometryBag"> Interface ESRI contenant les lignes qui ne respectent pas le sens d'écoulement.</param>
    '''<param name="pTopologyGraph"> Interface ESRI contenant la topologie des éléments visibles.</param>
    '''<param name="pTopoNode"> Interface ESRI contenant le Node de la topologie utilisé pour rechercher le réseau.</param>
    '''<param name="bEnleverSelection"> Indique si on doit enlever les éléments de la sélection, sinon ils sont conservés. Défaut=True.</param>
    ''' 
    Public Sub ValiderReseau(ByRef pFeatureSel As IFeatureSelection, ByRef pGeometryBag As IGeometryBag, _
                             ByRef pTopologyGraph As ITopologyGraph4, ByRef pTopoNode As ITopologyNode, _
                             Optional ByVal bEnleverSelection As Boolean = True)
        'Déclarer les variables de travail
        Dim pEnumNodeEdge As IEnumNodeEdge = Nothing        'Interface pour extraire les Nodes du Edge traité.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le Edge traité.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter la géométrie du Edge dans la Poilyline.
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.

        Try
            'Vérifier si le Node a été traité
            If Not pTopoNode.IsSelected Then
                'Interface pour ajouter les lignes inversées trouvées
                pGeomColl = CType(pGeometryBag, IGeometryCollection)

                'Indiquer que le Node est traité
                pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoNode)

                'Interface pour extraire tous les edges du Node
                pEnumNodeEdge = pTopoNode.Edges(True)

                'Extraire le premier edge du Node
                pEnumNodeEdge.Reset()
                pEnumNodeEdge.Next(pTopoEdge, True)

                'Traiter tous les Edges du Node
                Do Until pTopoEdge Is Nothing
                    'Vérifier si le Edge a été traité
                    If Not pTopoEdge.IsSelected Then
                        'Indiquer que le Node est traité
                        pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoEdge)

                        'Vérifier si le node traité est le même que celui de début du edge
                        If pTopoNode.Equals(pTopoEdge.FromNode) Then
                            'Vérifier si la ligne du Edge doit être sélectionnée
                            If bEnleverSelection Then
                                'Contient les parents sélectionnés
                                pEnumTopoParent = pTopoEdge.Parents

                                'Extraire le premier élément parent
                                pEnumTopoParent.Reset()
                                pEsriTopoParent = pEnumTopoParent.Next()

                                'Ajouter le OID dans la sélection
                                pFeatureSel.SelectionSet.Add(pEsriTopoParent.m_FID)

                                'Ajouter le Edge dans les géométries trouvées
                                pGeomColl.AddGeometry(pTopoEdge.Geometry)

                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pEsriTopoParent.m_FID.ToString & " #Erreur de sens d'écoulement invalide" _
                                                    & " /" & gsNomAttribut & " " & gsExpression, pTopoEdge.Geometry)
                            End If

                            'Valider le réseau par le ToNode
                            ValiderReseau(pFeatureSel, pGeometryBag, pTopologyGraph, pTopoEdge.ToNode)

                            'Si le node traité est le même que celui de fin du edge
                        Else
                            'Vérifier si la ligne du Edge doit être sélectionnée
                            If bEnleverSelection = False Then
                                'Contient les parents sélectionnés
                                pEnumTopoParent = pTopoEdge.Parents

                                'Extraire le premier élément parent
                                pEnumTopoParent.Reset()
                                pEsriTopoParent = pEnumTopoParent.Next()

                                'Ajouter le OID dans la sélection
                                pFeatureSel.SelectionSet.Add(pEsriTopoParent.m_FID)

                                'Ajouter le Edge dans les géométries trouvées
                                pGeomColl.AddGeometry(pTopoEdge.Geometry)

                                'Écrire une erreur
                                EcrireFeatureErreur("OID=" & pEsriTopoParent.m_FID.ToString & " #Sens d'écoulement valide" _
                                                    & " /" & gsNomAttribut & " " & gsExpression, pTopoEdge.Geometry)
                            End If

                            'Valider le réseau par le FromNode
                            ValiderReseau(pFeatureSel, pGeometryBag, pTopologyGraph, pTopoEdge.FromNode)
                        End If
                    End If

                    'Extraire le prochain edge du Node
                    pEnumNodeEdge.Next(pTopoEdge, True)
                Loop
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pEnumNodeEdge = Nothing
            pTopoEdge = Nothing
            pGeomColl = Nothing
            pEnumTopoParent = Nothing
            pEsriTopoParent = Nothing
        End Try
    End Sub
#End Region
End Class
