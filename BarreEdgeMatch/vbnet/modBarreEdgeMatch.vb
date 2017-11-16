Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Editor

'**
'Nom de la composante : modEdgeMatch.vb 
'
'''<summary>
'''Librairies de routines utilisée pour effectuer le EdgeMatch.
'''</summary>
'''
'''<remarks>
'''Auteur : Michel Pothier
'''Date : 4 juillet 2011
'''</remarks>
'''
Module modBarreEdgeMatch
    'Liste des variables publiques utilisées
    ''' <summary>Interface ESRI contenant l'application ArcMap.</summary>
    Public m_Application As IApplication = Nothing
    ''' <summary>Interface ESRI contenant le document ArcMap.</summary>
    Public m_MxDocument As IMxDocument = Nothing
    ''' <summary>Objet qui permet de gérer la Map et ses layers.</summary>
    Public m_MapLayer As clsGererMapLayer = Nothing
    ''' <summary>Interface ESRI contenant la classe de découpage.</summary>
    Public m_FeatureLayerDecoupage As IFeatureLayer = Nothing
    ''' <summary>Interface ESRI contenant le nom de l'attribut contenant l'identifiant de découpage.</summary>
    Public m_IdentifiantDecoupage As String = Nothing
    ''' <summary>Interface ESRI contenant les limites communes de découpage.</summary>
    Public m_LimiteDecoupage As IPolyline = Nothing
    ''' <summary>Interface ESRI contenant les limites communes de découpage avec les points d'adjacence.</summary>
    Public m_LimiteDecoupageAvecPoint As IPolyline = Nothing
    ''' <summary>Objet contenant la liste des attributs d'adjacence à valider.</summary>
    Public m_AttributAdjacence As Collection = Nothing
    ''' <summary>Interface ESRI contenant la liste des points d'adjacence.</summary>
    Public m_ListePointAdjacence As IGeometryCollection = Nothing
    ''' <summary>Objet contenant la liste des éléments à traiter.</summary>
    Public m_ListeElementTraiter As Collection = Nothing
    ''' <summary>Objet contenant la liste des éléments aux points d'adjacence.</summary>
    Public m_ListeElementPointAdjacent As Collection = Nothing
    ''' <summary>Objet contenant les points d'adjacence qui possède des erreurs.</summary>
    Public m_ListeErreurPointAdjacent As Collection = Nothing
    ''' <summary>Objet contenant les éléments d'erreurs de précision, d'adjacence et d'attributs.</summary>
    Public m_ErreurFeature As Collection = Nothing
    ''' <summary>Objet contenant les éléments d'erreurs de précision.</summary>
    Public m_ErreurFeaturePrecision As Collection = Nothing
    ''' <summary>Objet contenant les éléments d'erreurs d'adjacence.</summary>
    Public m_ErreurFeatureAdjacence As Collection = Nothing
    ''' <summary>Objet contenant les éléments d'erreurs d'attributs.</summary>
    Public m_ErreurFeatureAttribut As Collection = Nothing
    '''<summary>Interface contenant la référence spatiale projeter par défaut.</summary>
    Public m_SpatialReferenceProj As ISpatialReference = Nothing

    ''' <summary>Objet contenant le formulaire du EdgeMatch.</summary>
    Public m_MenuEdgeMatch As dckMenuEdgeMatch
    '''<summary> Contient la tolérance d'adjacence utilisée pour corriger les éléments adjacents.</summary>
    Public m_TolAdjacence As Double = 3.0
    '''<summary> Contient la tolérance d'adjacence originale utilisée pour corriger les éléments adjacents.</summary>
    Public m_TolAdjacenceOri As Double = 3.0
    '''<summary> Contient la tolérance de recherche des éléments à traiter.</summary>
    Public m_TolRecherche As Double = 1.0
    '''<summary> Contient la tolérance de recherche originale des éléments à traiter.</summary>
    Public m_TolRechercheOri As Double = 1.0
    '''<summary> Contient la précision des données utilisé pour traiter les éléments adjacents.</summary>
    Public m_Precision As Double = 0.001
    '''<summary> Indiquer si on permet d'avoir plusieurs éléments adjacents.</summary>
    Public m_AdjacenceUnique As Boolean = False
    '''<summary> Indiquer si on permet d'avoir des classes différentes entre les éléments adjacents.</summary>
    Public m_ClasseDifferente As Boolean = False
    '''<summary> Indiquer si on permet d'avoir des identifiants pareils entre les éléments adjacents.</summary>
    Public m_SansIdentifiant As Boolean = False

    '''<summary> Interface ESRI contenant le symbol pour la texte.</summary>
    Public m_SymboleTexte As ISymbol
    '''<summary> Interface ESRI contenant le symbol pour la géométrie de type point.</summary>
    Public m_SymbolePoint As ISymbol
    '''<summary> Interface ESRI contenant le symbol pour la géométrie de type ligne.</summary>
    Public m_SymboleLigne As ISymbol
    '''<summary> Interface ESRI contenant le symbol pour la géométrie de type surface.</summary>
    Public m_SymboleSurface As ISymbol

    ' '''<summary>Interface ESRI qui permet de gérer l'événement lorsque qu'un nouveau document est ouvert.</summary>
    'Public m_DocumentEventsOpenDocument As IDocumentEvents_OpenDocumentEventHandler
    ' '''<summary>Interface ESRI qui permet de gérer l'événement lorsqu'un nouveau document est activé.</summary>
    'Public m_ActiveViewEventsContentsChanged As IActiveViewEvents_ContentsChangedEventHandler
    ' '''<summary>Interface ESRI qui permet de gérer l'événement lorsqu'un nouvel item est ajouté à la Map active.</summary>
    'Public m_ActiveViewEventsItemAdded As IActiveViewEvents_ItemAddedEventHandler
    ' '''<summary>Interface ESRI qui permet de gérer l'événement lors d'un changement de référence spatiale à la Map active.</summary>
    'Public m_ActiveViewEventsSpatialReferenceChanged As IActiveViewEvents_SpatialReferenceChangedEventHandler
    ' '''<summary>Interface ESRI qui permet de gérer l'événement lorsqu'un nouvel item est retiré de la Map active</summary>
    'Public m_ActiveViewEventsItemDeleted As IActiveViewEvents_ItemDeletedEventHandler

    '''<summary>Valeur initiale de la dimension en hauteur du menu.</summary>
    Public m_Height As Integer = 300
    '''<summary>Valeur initiale de la dimension en largeur du menu.</summary>
    Public m_Width As Integer = 300
    '''<summary>Valeur initiale de la différence de dimension en hauteur entre le menu et le listbox de classes.</summary>
    Public m_ClasseHeight As Integer = 179
    '''<summary>Valeur initiale de la différence de dimension en hauteur entre le menu et le listbox d'attributs.</summary>
    Public m_AttributHeight As Integer = 231

    '''<summary>
    '''Définir la liste des groupes permis d'une incohérence
    '''</summary>
    ''' 
    Public Enum citsTypeErreur
        citsErreurInconnu = -1
        citsErreurPrecision = 1
        citsErreurAdjacenceIsole = 2
        citsIGStructure = 3
        citsIGContenu = 4
        citsIGPrecision = 5
    End Enum

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
        Dim ElementA As Integer
        Dim IdentifiantA As String
        Dim ValeurA As String
        Dim PosAttA As Integer

        'Définition de l'information pour le point B
        Dim PointB As IPoint
        Dim FeatureB As IFeature
        Dim ElementB As Integer
        Dim IdentifiantB As String
        Dim ValeurB As String
        Dim PosAttB As Integer
    End Structure

    '''<summary>
    ''' Initialiser les couleurs par défaut pour le texte qui sera affiché avec les
    ''' géométries déssiné et initialiser les 3 couleurs pour les géométries de type
    ''' point, ligne et surface.
    '''</summary>
    '''
    Public Sub InitSymbole()
        'Déclarer les variables de travail
        Dim pRgbColor As IRgbColor = Nothing                'Interface ESRI contenant la couleur RGB.
        Dim pTextSymbol As ITextSymbol = Nothing            'Interface ESRI contenant un symbole de texte.
        Dim pMarkerSymbol As ISimpleMarkerSymbol = Nothing  'Interface ESRI contenant un symbole de point.
        Dim pLineSymbol As ISimpleLineSymbol = Nothing      'Interface ESRi contenant un symbole de ligne.
        Dim pFillSymbol As ISimpleFillSymbol = Nothing      'Interface ESRI contenant un symbole de surface.

        'Permet d'initialiser la symbologie
        Try
            'Vérifier si le symbole est invalide
            If m_SymboleTexte Is Nothing Then
                'Définir la couleur pour le texte
                pRgbColor = New RgbColor
                pRgbColor.RGB = 0
                'Définir la symbologie pour le texte
                pTextSymbol = New ESRI.ArcGIS.Display.TextSymbol
                pTextSymbol.Color = pRgbColor
                pTextSymbol.Font.Bold = True
                pTextSymbol.HorizontalAlignment = esriTextHorizontalAlignment.esriTHACenter
                'Conserver le symbole
                m_SymboleTexte = CType(pTextSymbol, ISymbol)
            End If

            'Vérifier si le symbole est invalide
            If m_SymbolePoint Is Nothing Then
                'Définir la couleur rouge pour le polygon
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                'Définir la symbologie pour la limite d'un polygone
                pMarkerSymbol = New SimpleMarkerSymbol
                pMarkerSymbol.Color = pRgbColor
                pMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCircle
                pMarkerSymbol.Size = 3
                'Conserver le symbole 
                m_SymbolePoint = CType(pMarkerSymbol, ISymbol)
            End If

            'Vérifier si le symbole est invalide
            If m_SymboleLigne Is Nothing Then
                'Définir la couleur rouge pour le polygon
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                'Définir la symbologie pour la limite d'un polygone
                pLineSymbol = New SimpleLineSymbol
                pLineSymbol.Color = pRgbColor
                pLineSymbol.Style = esriSimpleLineStyle.esriSLSDot
                pLineSymbol.Width = 3
                'Conserver le symbole en mémoire
                m_SymboleLigne = CType(pLineSymbol, ISymbol)
            End If

            'Vérifier si le symbole est invalide
            If m_SymboleSurface Is Nothing Then
                'Définir la couleur rouge pour le polygon
                pRgbColor = New RgbColor
                pRgbColor.Red = 255
                'Définir la symbologie pour la limite d'un polygone
                pLineSymbol = New SimpleLineSymbol
                pLineSymbol.Color = pRgbColor
                'Définir la symbologie pour l'intérieur d'un polygone
                pFillSymbol = New SimpleFillSymbol
                pFillSymbol.Color = pRgbColor
                pFillSymbol.Outline = pLineSymbol
                pFillSymbol.Style = esriSimpleFillStyle.esriSFSBackwardDiagonal
                'Conserver le symbole 
                m_SymboleSurface = CType(pFillSymbol, ISymbol)
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            pRgbColor = Nothing
            pTextSymbol = Nothing
            pMarkerSymbol = Nothing
            pLineSymbol = Nothing
            pFillSymbol = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de dessiner dans la vue active les géométries Point, MultiPoint, 
    ''' Polyline et/ou Polygon. Ces géométries peuvent être contenu dans un GeometryBag.
    ''' 
    ''' Un point est représenté par un carré, une ligne est représentée par une ligne pleine 
    ''' et la surface est représentée par une ligne pleine pour la limite et des lignes à
    ''' 45 dégrés pour l'intérieur. 
    ''' 
    ''' On peut également afficher le numéro de la géométrie pour un GeometryBag.
    '''</summary>
    '''
    '''<param name="pGeometry"> Interface ESRI contenant la géométrie èa dessiner.</param>
    '''<param name="bRafraichir"> Indique si on doit rafraîchir la vue active.</param>
    ''' 
    '''<return>Un booleen est retourner pour indiquer la fonction s'est bien exécutée.</return>
    '''
    Public Function bDessinerGeometrie(ByVal pGeometry As IGeometry, Optional ByVal bRafraichir As Boolean = False) As Boolean
        'Déclarer les variables de travail
        Dim pScreenDisplay As IScreenDisplay = Nothing  'Interface ESRI contenant le document de ArcMap.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface ESRI contenant la fenêtre d'affichage.
        Dim pArea As IArea = Nothing                    'Interface ESRI contenant un symbole de texte.
        Dim pGeomTexte As IGeometry = Nothing           'Interface ESRI pour la position du texte d'une surface.
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Vérifier si la géométrie est absente
            If pGeometry Is Nothing Then Exit Function

            'Vérifier si la géométrie est vide
            If pGeometry.IsEmpty Then Exit Function

            'Initialiser les variables de travail
            pScreenDisplay = m_MxDocument.ActiveView.ScreenDisplay

            'Vérifier si on doit rafraichir l'écran
            If bRafraichir Then
                'Rafraîchier l'affichage
                m_MxDocument.ActiveView.Refresh()
                System.Windows.Forms.Application.DoEvents()
            End If

            'Transformation du système de coordonnées selon la vue active
            pGeometry.Project(m_MxDocument.FocusMap.SpatialReference)

            'Vérifier si la géométrie est un GeometryBag
            If pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                'Interface pour traiter toutes les géométries présentes dans le GeometryBag
                pGeomColl = CType(pGeometry, IGeometryCollection)

                'Dessiner toutes les géométrie présentes dans une collection de géométrie
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Vérifier si on veut afficher le texte du numéro d'élément
                    If Not pGeomColl.Geometry(i).IsEmpty Then
                        'Dessiner la géométrie traitée
                        Call bDessinerGeometrie(pGeomColl.Geometry(i), False)
                        'Trouver le premier point de la géométrie
                        pArea = CType(pGeomColl.Geometry(i).Envelope, IArea)
                        pGeomTexte = CType(pArea.Centroid, IGeometry)

                        'Afficher le texte avec sa symbologie dans la vue active
                        With pScreenDisplay
                            .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                            .SetSymbol(m_SymboleTexte)
                            .DrawText(pGeomTexte, CStr(i + 1))
                            .FinishDrawing()
                        End With
                    End If
                Next i

                'Vérifier si la géométrie est un point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(m_SymbolePoint)
                    .DrawPoint(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est un multi-point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(m_SymbolePoint)
                    .DrawMultipoint(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est une ligne
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(m_SymboleLigne)
                    .DrawPolyline(pGeometry)
                    .FinishDrawing()
                End With

                'Vérifier si la géométrie est une surface
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon _
            Or pGeometry.GeometryType = esriGeometryType.esriGeometryEnvelope Then
                'Afficher la géométrie avec sa symbologie dans la vue active
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                    .SetSymbol(m_SymboleSurface)
                    .DrawPolygon(pGeometry)
                    .FinishDrawing()
                End With
            End If

            'Retourner le résultat
            bDessinerGeometrie = True

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            pScreenDisplay = Nothing
            pGeomColl = Nothing
            pArea = Nothing
            pGeomTexte = Nothing
            i = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de dessiner dans la vue active les géométries Point, MultiPoint, 
    ''' Polyline et/ou Polygon. Ces géométries peuvent être contenu dans un GeometryBag.
    ''' 
    ''' Un point est représenté par un carré, une ligne est représentée par une ligne pleine 
    ''' et la surface est représentée par une ligne pleine pour la limite et des lignes à
    ''' 45 dégrés pour l'intérieur. 
    ''' 
    ''' On peut également afficher le numéro de la géométrie pour un GeometryBag.
    '''</summary>
    '''
    '''<param name="pPoint"> Interface ESRI contenant le point nécessaire (position) pour dessiner le texte.</param>
    '''<param name="sTexte"> Contient le texte à dessiner.</param>
    '''<param name="bRafraichir"> Indique si on doit rafraîchir la vue active.</param>
    ''' 
    '''<return>Un booleen est retourner pour indiquer la fonction s'est bien exécutée.</return>
    '''
    Public Function bDessinerTexte(ByVal pPoint As IPoint, ByVal sTexte As String, Optional ByVal bRafraichir As Boolean = False) As Boolean
        'Déclarer les variables de travail
        Dim pScreenDisplay As IScreenDisplay = Nothing  'Interface ESRI contenant le document de ArcMap

        Try
            'Vérifier si la géométrie est absente
            If pPoint Is Nothing Then Exit Function

            'Vérifier si la géométrie est vide
            If pPoint.IsEmpty Then Exit Function

            'Initialiser les variables de travail
            pScreenDisplay = m_MxDocument.ActiveView.ScreenDisplay

            'Vérifier si on doit rafraichir l'écran
            If bRafraichir Then
                'Rafraîchier l'affichage
                m_MxDocument.ActiveView.Refresh()
                System.Windows.Forms.Application.DoEvents()
            End If

            'Transformation du système de coordonnées selon la vue active
            pPoint.Project(m_MxDocument.FocusMap.SpatialReference)

            'Afficher le texte avec sa symbologie dans la vue active
            With pScreenDisplay
                .StartDrawing(pScreenDisplay.hDC, CType(ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache, Short))
                .SetSymbol(m_SymboleTexte)
                .DrawText(pPoint, sTexte)
                .FinishDrawing()
            End With

            'Retourner le résultat
            bDessinerTexte = True

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            pScreenDisplay = Nothing
        End Try
    End Function


    '''<summary>
    ''' Routine qui permet de dessiner la géométrie et le texte des erreurs de précision, d'adjacence et d'attribut.
    '''</summary>
    '''
    Public Sub DessinerErreurs()
        'Déclarer les variables de travail
        Dim pMouseCursor As IMouseCursor = Nothing  'Interface qui permet de changer l'image du curseur
        Dim oErreur As Structure_Erreur = Nothing   'Structure contenant une erreur d'attribut
        Dim pPoint As IPoint = Nothing              'Interface ESRI contenant le point d'adjacence sélectionné
        Dim i As Integer = Nothing                  'Index du point d'adjacence

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)

            'Vérifier si on doit dessiner les erreurs de précision
            If m_MenuEdgeMatch.chkPrecision.Checked Then
                'Traiter toutes les erreurs
                For i = 1 To m_ErreurFeaturePrecision.Count
                    'Définir l'erreur
                    oErreur = CType(m_ErreurFeaturePrecision.Item(i), Structure_Erreur)

                    'Vérifier si le point B est présent
                    If oErreur.PointB Is Nothing Then
                        'Définir le point d'erreur sélectionné
                        pPoint = oErreur.PointA
                        pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
                    Else
                        'Définir le point d'erreur sélectionné
                        pPoint = oErreur.PointB
                        pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
                    End If

                    'Dessiner la géométrie de l'erreur
                    Call bDessinerGeometrie(pPoint, False)
                    'Dessiner le texte de l'erreur
                    Call bDessinerTexte(pPoint, i.ToString, False)
                Next
            End If

            'Vérifier si on doit dessiner les erreurs d'adjacence
            If m_MenuEdgeMatch.chkAdjacence.Checked Then
                'Traiter toutes les erreurs
                For i = 1 To m_ErreurFeatureAdjacence.Count
                    'Définir l'erreur
                    oErreur = CType(m_ErreurFeatureAdjacence.Item(i), Structure_Erreur)

                    'Vérifier si le point B est présent
                    If oErreur.PointB Is Nothing Then
                        'Définir le point d'erreur sélectionné
                        pPoint = oErreur.PointA
                        pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
                    Else
                        'Définir le point d'erreur sélectionné
                        pPoint = oErreur.PointB
                        pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
                    End If

                    'Dessiner la géométrie de l'erreur
                    Call bDessinerGeometrie(pPoint, False)
                    'Dessiner le texte de l'erreur
                    Call bDessinerTexte(pPoint, (m_ErreurFeaturePrecision.Count + i).ToString, False)
                Next
            End If

            'Vérifier si on doit dessiner les erreurs d'attribut
            If m_MenuEdgeMatch.chkAttribut.Checked Then
                'Traiter toutes les erreurs
                For i = 1 To m_ErreurFeatureAttribut.Count
                    'Définir l'erreur
                    oErreur = CType(m_ErreurFeatureAttribut.Item(i), Structure_Erreur)

                    'Vérifier si le point B est présent
                    If oErreur.PointB Is Nothing Then
                        'Définir le point d'erreur sélectionné
                        pPoint = oErreur.PointA
                        pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
                    Else
                        'Définir le point d'erreur sélectionné
                        pPoint = oErreur.PointB
                        pPoint.Project(m_MxDocument.FocusMap.SpatialReference)
                    End If

                    'Dessiner la géométrie de l'erreur
                    Call bDessinerGeometrie(pPoint, False)
                    'Dessiner le texte de l'erreur
                    Call bDessinerTexte(pPoint, (m_ErreurFeaturePrecision.Count + m_ErreurFeatureAdjacence.Count + i).ToString, False)
                Next
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pMouseCursor = Nothing
            oErreur = Nothing
            pPoint = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'afficher la fenêtre graphique selon l'enveloppe de la limite commune de découpage plus 10%.
    '''</summary>
    '''
    Public Sub ZoomLimiteDecoupage()
        'Déclarer les variables de travail
        Dim pEnvelope As IEnvelope = Nothing    'Interface ESRI contenant l'enveloppe de l'unité de travail.

        Try
            'Transformation du système de coordonnées selon la vue active
            m_LimiteDecoupage.Project(m_MxDocument.FocusMap.SpatialReference)

            'Définir l'enveloppe de la limite commune
            pEnvelope = m_LimiteDecoupage.Envelope

            'Agrandir l'enveloppe de 10% de l'élément en erreur
            pEnvelope.Expand(pEnvelope.Width / 10, pEnvelope.Height / 10, False)

            'Vérifier si la hauteur est invalide
            If pEnvelope.Height <= m_TolRecherche Then
                'Définir la nouvelle fenêtre de travail en Y
                pEnvelope.YMax = pEnvelope.YMax + m_TolRecherche
                pEnvelope.YMin = pEnvelope.YMin - m_TolRecherche
                'si la hauteur est invalide
            ElseIf pEnvelope.Width <= m_TolRecherche Then
                'Définir la nouvelle fenêtre de travail en X
                pEnvelope.XMax = pEnvelope.XMax + m_TolRecherche
                pEnvelope.XMin = pEnvelope.XMin - m_TolRecherche
            End If

            'Définir la nouvelle fenêtre de travail
            m_MxDocument.ActiveView.Extent = pEnvelope

            'Rafraîchier l'affichage
            m_MxDocument.ActiveView.Refresh()

            'Permet de vider la mémoire sur les évènements
            System.Windows.Forms.Application.DoEvents()

        Finally
            'Vider la mémoire
            pEnvelope = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de définir et afficher les limites communes entre les éléments de la classe de découpage 
    ''' qui sont présents dans la fenêtre graphique. 
    '''</summary>
    '''
    Public Sub DefinirLimiteDecoupage()
        'Déclarer les variables de travail
        Dim pMap As IMap = Nothing                      'Interface utilisé pour sélectionner les éléments.
        Dim pFeatureSel As IFeatureSelection = Nothing  'Interface utilisé pour vérifier la présence des éléments sélectionnés.
        Dim pSelectionSet As ISelectionSet = Nothing    'Interface contenant les éléments sélectionnés.
        Dim pFeatureCursor As IFeatureCursor = Nothing  'Interface utilisé pour extraire les éléments sélectionnés.
        Dim pCursor As ICursor = Nothing                'Interface utilisé pour extraire les éléments sélectionnés.
        Dim pFeature As IFeature = Nothing              'Interface contenant un élément sélectionné.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter des géométries.
        Dim pGeomBag As IGeometryCollection = Nothing   'Interface contenant toutes les géométries des éléments.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface utilisé pour pour calculer l'intersection des limites commune.
        Dim pGeoDataset As IGeoDataset = Nothing        'Interface contenant la référence spatiale.
        Dim i As Integer = Nothing                      'Compteur
        Dim j As Integer = Nothing                      'Compteur

        Try
            'Rendre sélectionnable la classe de découpage
            m_FeatureLayerDecoupage.Selectable = True

            'Interface pour extraire les éléments sélectionnés
            pFeatureSel = CType(m_FeatureLayerDecoupage, IFeatureSelection)

            'Définir la Map active
            pMap = m_MxDocument.FocusMap

            'Vérifier si moins de deux éléments sélectionnés
            If pFeatureSel.SelectionSet.Count = 0 Then
                'Sélectionner tous les éléments sélectionnable
                pMap.SelectByShape(m_MxDocument.ActiveView.Extent, Nothing, False)
            End If

            'Interface contenant les éléments sélectionnés
            pSelectionSet = pFeatureSel.SelectionSet

            'Vérifier le nombre d'éléments sélectionné
            If pSelectionSet.Count > 1 Then
                'Définir la référence spatiale
                pGeoDataset = CType(m_FeatureLayerDecoupage, IGeoDataset)

                'Définir un nouveau Bag
                pGeometry = New GeometryBag
                pGeometry.SpatialReference = pGeoDataset.SpatialReference
                pGeomBag = CType(pGeometry, IGeometryCollection)

                'Extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, True, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature

                'Traiter tous les éléments
                Do Until pFeature Is Nothing
                    'Ajouter la géométrie dans la Bag
                    pGeomBag.AddGeometry(pFeature.ShapeCopy)

                    'Extraire le premier élément
                    pFeature = pFeatureCursor.NextFeature
                Loop

                'Désélectionner tous les éléments
                pMap.ClearSelection()

                'Transformer les tolérances en géographique au besoin
                pGeometry = CType(pGeomBag, IGeometry)
                Call TransformerTolerances(pGeometry.Envelope)

                'Définir la géométrie contenant les limites de découpage
                m_LimiteDecoupage = CType(New Polyline, IPolyline)
                m_LimiteDecoupage.SpatialReference = pGeoDataset.SpatialReference
                'Interface pour ajouter les limites de découpage
                pGeomColl = CType(m_LimiteDecoupage, IGeometryCollection)

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
                If m_LimiteDecoupage.IsEmpty Then
                    'Afficher un message
                    MsgBox("ATTENTION : Aucune limite commune de découpage trouvée!")
                Else
                    'Interface pour simplifier les limites communes
                    pTopoOp = CType(m_LimiteDecoupage, ITopologicalOperator2)
                    'Simplifier les limites communes
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()
                End If

                'Si seulement un éléments de découpage
            ElseIf pSelectionSet.Count = 1 Then
                'Extraire les éléments sélectionnés
                pSelectionSet.Search(Nothing, True, pCursor)
                pFeatureCursor = CType(pCursor, IFeatureCursor)

                'Extraire le premier élément
                pFeature = pFeatureCursor.NextFeature

                'Transformer les tolérances en géographique au besoin
                pGeometry = CType(pFeature.Shape, IGeometry)
                Call TransformerTolerances(pGeometry.Envelope)

                'Interface pour simplifier les limites communes
                pTopoOp = CType(pFeature.ShapeCopy, ITopologicalOperator2)

                'Définir la limite de découpage
                m_LimiteDecoupage = CType(pTopoOp.Boundary, IPolyline)

                'Désélectionner tous les éléments
                pMap.ClearSelection()

                'Si aucun élément de découpage n'est trouvé
            Else
                'Afficher un message
                MsgBox("ATTENTION : Aucun élément de découpage n'est trouvé!")
            End If

        Catch e As Exception
            'Message d'erreur
            MessageBox.Show(e.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pMap = Nothing
            pFeatureSel = Nothing
            pSelectionSet = Nothing
            pFeatureCursor = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
            pGeomBag = Nothing
            pTopoOp = Nothing
            pGeoDataset = Nothing
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
        Dim bSelection As Boolean = Nothing             'Indice de sélection des éléments de la classe de découpage

        Try
            'Définir la nouvelle liste des éléments pour chaque point d'adjacence
            m_ListeElementPointAdjacent = New Collection
            'Définir la nouvelle liste des points d'adjacence en erreur
            m_ListeErreurPointAdjacent = New Collection
            'Définir une nouvelle Collection pour les éléments en erreur
            m_ErreurFeature = New Collection
            m_ErreurFeaturePrecision = New Collection
            m_ErreurFeatureAdjacence = New Collection
            m_ErreurFeatureAttribut = New Collection

            'Définir un nouveau Bag pour les points d'adjacence
            pGeometry = New GeometryBag
            pGeometry.SpatialReference = m_MxDocument.ActiveView.FocusMap.SpatialReference
            'Définir un nouveau Bag pour les points d'adjacence
            m_ListePointAdjacence = CType(pGeometry, IGeometryCollection)

            'Désactiver la sélection du layer de découpage
            bSelection = m_FeatureLayerDecoupage.Selectable
            m_FeatureLayerDecoupage.Selectable = False

            'Définition des points d'adjacence virtuels
            Call DefinirPointAdjacenceVirtuel(pTrackCancel)

            'Trouver les points d'adjacence et les éléments qui y sont liés
            Call TrouverListePointAdjacence(pTrackCancel)

            'Valider la distance entre les points d'adjacence
            Call ValiderDistanceListePointAdjacence(pTrackCancel)

            'Valider l'adjacence de la liste des points d'adjacence
            Call ValiderAdjacenceListePointAdjacence(pTrackCancel)

            'Remettre la sélection du layer de découpage
            m_FeatureLayerDecoupage.Selectable = bSelection

            'Désélectionner tous les éléments
            m_MxDocument.ActiveView.FocusMap.ClearSelection()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
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
            m_LimiteDecoupage.Project(m_MxDocument.ActiveView.FocusMap.SpatialReference)

            'Cloner la limite de découpage
            pCloneLimite = CType(m_LimiteDecoupage, IClone)
            m_LimiteDecoupageAvecPoint = CType(pCloneLimite.Clone, IPolyline)

            'Interface pour traiter les sommets
            pAdjColl = CType(m_LimiteDecoupageAvecPoint, IPointCollection)
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
        Dim pEnumFeature As IEnumFeature = Nothing      'Interface ESRI utilisé pour extraire les éléments sélectionnés.
        Dim pEnumFeatureSet As IEnumFeatureSetup = Nothing  'Interface ESRI utilisé pour extraire les éléments sélectionnés.
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
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Sélection des éléments à la limite du découpage ..."

            'Créer la nouvelle collection des éléments à traiter
            m_ListeElementTraiter = New Collection

            'Interface utilisé pour extraire le point correspondant exactement aux limites communes
            pHitTest = CType(m_LimiteDecoupageAvecPoint, IHitTest)

            'Interface qui permet d'extraire la composante trouvée
            pGeomColl = CType(m_LimiteDecoupageAvecPoint, IGeometryCollection)

            'Interface pour créer le buffer
            pTopoOp = CType(m_LimiteDecoupageAvecPoint, ITopologicalOperator2)
            'Créer le buffer de la limite
            pBufferLimite = pTopoOp.Buffer(m_TolRecherche)
            'Simplifier le buffer de la limite
            pTopoOp = CType(pBufferLimite, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Sélectionner tous les éléments sélectionnable
            m_MxDocument.ActiveView.FocusMap.SelectByShape(pBufferLimite, Nothing, False)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Identification des points d'adjacence ..."
            'Afficher la barre de progression
            InitBarreProgression(0, m_MxDocument.ActiveView.FocusMap.SelectionCount, pTrackCancel)

            'Interface pour extraire les éléments sélectionnés
            pEnumFeature = CType(m_MxDocument.ActiveView.FocusMap.FeatureSelection, IEnumFeature)
            'Interface pour extraire toutes les valeurs d'attributs
            pEnumFeatureSet = CType(pEnumFeature, IEnumFeatureSetup)
            pEnumFeatureSet.AllFields = True
            pEnumFeatureSet.Recycling = True

            'Extraire le premier élément
            pEnumFeature.Reset()
            pFeature = pEnumFeature.Next

            'Traiter tous les éléments
            Do Until pFeature Is Nothing
                'Créer un lien vers un élément à traiter
                qElementTraiter = New Structure_Element_Traiter
                qElementTraiter.OID = pFeature.OID
                qElementTraiter.FeatureClass = CType(pFeature.Class, IFeatureClass)
                'Ajouter l'élément dans la liste à traiter avec un numéro d'index
                m_ListeElementTraiter.Add(qElementTraiter, (m_ListeElementTraiter.Count + 1).ToString)

                'Définir la valeur de l'identifiant
                sIdentifiant = fsIdentifiantDecoupage(pFeature, m_IdentifiantDecoupage)

                'Interface contenant la géométrie de l'élément
                pGeometry = pFeature.ShapeCopy
                'Projeter la géométrie
                pGeometry.Project(pBufferLimite.SpatialReference)

                'Vérifier si la géométrie est une surface
                If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Convertir la géométrie en multipoint
                    pAdjacence = fpConvertirEnMultiPoint(pGeometry)
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
                    pAdjacence = fpConvertirEnMultiPoint(pGeometry)
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
                If pTrackCancel.Continue = False Then Exit Do

                'Extraire le premier élément
                pFeature = pEnumFeature.Next
            Loop

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pEnumFeature = Nothing
            pEnumFeatureSet = Nothing
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
            If m_AdjacenceUnique Then
                'Définir la tolérance selon celle de recherche
                dTolerance = m_TolRecherche
                'Si l'adjacence multiple est permise
            Else
                'Définir la tolérance selon celle d'adjacence
                dTolerance = m_TolAdjacence
            End If

            'Vérifier la présence d'un sommet existant selon une tolérance
            If pHitTest.HitTest(pPoint, dTolerance, esriGeometryHitPartType.esriGeometryPartVertex, _
                                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                'Définir la référence spatiale du point trouvé
                pNewPoint.SpatialReference = pPoint.SpatialReference
                'Vérifier si c'est un point non identifié
                If pNewPoint.ID = 0 Then
                    'Ajouter la géométrie d'adjacence
                    m_ListePointAdjacence.AddGeometry(pNewPoint)
                    'Définir le ID du sommet
                    pNewPoint.ID = m_ListePointAdjacence.GeometryCount * -1
                    'Mettre à jour le ID du point
                    pAdjColl = CType(pGeomColl.Geometry(nNumeroPartie), IPointCollection)
                    pAdjColl.UpdatePoint(nNumeroSommet, pNewPoint)
                    'Ajouter l'élément adjacent
                    qElementColl = New Collection
                    qElementColl.Add(m_ListeElementTraiter.Count)
                    'Ajouter le point d'adjacence
                    m_ListeElementPointAdjacent.Add(qElementColl, pNewPoint.ID.ToString)
                Else
                    'Définir la liste des éléments adjacents
                    qElementColl = CType(m_ListeElementPointAdjacent.Item(pNewPoint.ID.ToString), Collection)
                    qElementColl.Add(m_ListeElementTraiter.Count)
                End If

                'Vérifier si la distance est plus grande que la précision
                If dDistance > m_Precision Then
                    'Ajouter une erreur de précision
                    pPoint.ID = pNewPoint.ID

                    'Projeter les points
                    pPoint.Project(m_SpatialReferenceProj)
                    pNewPoint.Project(m_SpatialReferenceProj)
                    'Calculer la distance
                    pProxOp = CType(pPoint, IProximityOperator)
                    dDistance = pProxOp.ReturnDistance(pNewPoint)

                    'Vérifier si la distance est plus grande que la tolérance de recherche
                    If dDistance > m_TolRechercheOri Then
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + m_TolAdjacenceOri.ToString("0.0##") + ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.ElementA = m_ListeElementTraiter.Count
                        oErreur.FeatureA = pFeature
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        m_ErreurFeaturePrecision.Add(oErreur)

                    Else
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Ne touche pas à un point existant (" & dDistance.ToString("0.0##") & "<=" & m_TolRechercheOri.ToString("0.0##") & ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.FeatureA = pFeature
                        oErreur.ElementA = m_ListeElementTraiter.Count
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        m_ErreurFeaturePrecision.Add(oErreur)
                    End If

                    'Conserver l'dentifiant du point d'adjacence en erreur
                    If Not m_ListeErreurPointAdjacent.Contains(pNewPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pNewPoint.ID, pNewPoint.ID.ToString)
                End If

                'Vérifier la présence d'un sommet sur une droite selon une tolérance
            ElseIf pHitTest.HitTest(pPoint, m_TolRecherche, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                   pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                'Définir la référence spatiale du point trouvé
                pNewPoint.SpatialReference = pPoint.SpatialReference
                'Ajouter la géométrie d'adjacence
                m_ListePointAdjacence.AddGeometry(pNewPoint)
                'Interface pour ajouter un sommet
                pAdjColl = CType(pGeomColl.Geometry(nNumeroPartie), IPointCollection)
                'Définir le ID du sommet
                pNewPoint.ID = m_ListePointAdjacence.GeometryCount
                'Insérer un nouveau sommet
                pAdjColl.InsertPoints(nNumeroSommet + 1, 1, pNewPoint)
                'Ajouter l'élément adjacent
                qElementColl = New Collection
                qElementColl.Add(m_ListeElementTraiter.Count)
                'Ajouter le point d'adjacence
                m_ListeElementPointAdjacent.Add(qElementColl, pNewPoint.ID.ToString)

                'Vérifier si la distance est plus grande que la précision
                If dDistance > m_Precision Then
                    'Ajouter une erreur de précision
                    pPoint.ID = pNewPoint.ID

                    'Projeter les points
                    pPoint.Project(m_SpatialReferenceProj)
                    pNewPoint.Project(m_SpatialReferenceProj)
                    'Calculer la distance
                    pProxOp = CType(pPoint, IProximityOperator)
                    dDistance = pProxOp.ReturnDistance(pNewPoint)

                    'Vérifier si la distance est plus grande que la tolérance de recherche
                    If dDistance > m_TolRechercheOri Then
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + m_TolAdjacenceOri.ToString("0.0##") + ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.FeatureA = pFeature
                        oErreur.ElementA = m_ListeElementTraiter.Count
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        m_ErreurFeaturePrecision.Add(oErreur)

                    Else
                        'Conserver l'erreur dans une structure
                        oErreur = New Structure_Erreur
                        oErreur.Description = "Ne touche pas à la limite (" & dDistance.ToString("0.0##") & "<=" & m_TolRechercheOri.ToString("0.0##") & ")"
                        oErreur.IdentifiantA = sIdentifiant
                        oErreur.PointA = pNewPoint
                        oErreur.FeatureA = pFeature
                        oErreur.ElementA = m_ListeElementTraiter.Count
                        oErreur.Distance = dDistance
                        oErreur.PointB = pPoint

                        'Ajouter la structure d'erreur dans la collection
                        m_ErreurFeaturePrecision.Add(oErreur)
                    End If

                    'Conserver l'dentifiant du point d'adjacence en erreur
                    If Not m_ListeErreurPointAdjacent.Contains(pNewPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pNewPoint.ID, pNewPoint.ID.ToString)
                End If
            End If

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
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
            pPointColl = CType(m_LimiteDecoupageAvecPoint, IPointCollection)

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
                    pPointA.Project(m_SpatialReferenceProj)
                    pPointB.Project(m_SpatialReferenceProj)

                    'Interface pour calculer la distance d'adjacence
                    pProxOp = CType(pPointA, IProximityOperator)
                    'Calculer la distance d'adjacence
                    dDistance = pProxOp.ReturnDistance(pPointB)

                    'Sortir si la distance est plus grande que la tolérance d'adjacence
                    If dDistance <= m_TolAdjacenceOri Then
                        'Définir la liste des sommets d'éléments adjacents
                        qElementCollA = CType(m_ListeElementPointAdjacent.Item(Math.Abs(pPointA.ID)), Collection)
                        'Définir l'élément
                        qElementTraiterA = CType(m_ListeElementTraiter.Item(qElementCollA.Item(1)), Structure_Element_Traiter)
                        pFeatureA = qElementTraiterA.FeatureClass.GetFeature(qElementTraiterA.OID)
                        'Définir la valeur de l'identifiant
                        sIdentifiantA = fsIdentifiantDecoupage(pFeatureA, m_IdentifiantDecoupage)

                        'Définir la liste des sommets d'éléments adjacents
                        qElementCollB = CType(m_ListeElementPointAdjacent.Item(Math.Abs(pPointB.ID)), Collection)
                        'Définir l'élément
                        qElementTraiterB = CType(m_ListeElementTraiter.Item(qElementCollB.Item(1)), Structure_Element_Traiter)
                        pFeatureB = qElementTraiterB.FeatureClass.GetFeature(qElementTraiterB.OID)
                        'Définir la valeur de l'identifiant
                        sIdentifiantB = fsIdentifiantDecoupage(pFeatureB, m_IdentifiantDecoupage)

                        'Vérifier si une erreur de déplacement doit être identifiée
                        'Si l'adjacence est unique, on peut déplacer si les deux sommets ont seulement 1 élément
                        If m_AdjacenceUnique = False Or (m_AdjacenceUnique And qElementCollA.Count = 1 And qElementCollB.Count = 1) Then
                            'Conserver l'erreur dans une structure
                            oErreur = New Structure_Erreur
                            oErreur.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + m_TolAdjacenceOri.ToString("0.0##") + ")"
                            oErreur.IdentifiantA = sIdentifiantA
                            oErreur.PointA = pPointA
                            oErreur.FeatureA = pFeatureA
                            oErreur.IdentifiantB = sIdentifiantB
                            oErreur.PointB = pPointB
                            oErreur.FeatureB = pFeatureB
                            oErreur.Distance = dDistance

                            'Ajouter la structure d'erreur dans la collection
                            m_ErreurFeatureAdjacence.Add(oErreur)
                            'Conserver l'dentifiant du point d'adjacence en erreur
                            If Not m_ListeErreurPointAdjacent.Contains(pPointA.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPointA.ID, pPointA.ID.ToString)
                            'Conserver l'dentifiant du point d'adjacence en erreur
                            If Not m_ListeErreurPointAdjacent.Contains(pPointB.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPointB.ID, pPointB.ID.ToString)
                        End If
                    End If
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit For
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
            pPointColl = CType(m_LimiteDecoupageAvecPoint, IPointCollection)

            'Afficher le message de lecture des éléments
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Trouver les éléments adjacents aux points d'adjacence isolés ..."
            'Afficher la barre de progression
            InitBarreProgression(0, pPointColl.PointCount - 2, pTrackCancel)

            'Traiter tous les points de découpage
            For i = 0 To pPointColl.PointCount - 2
                'Définir le point en erreur
                pPointA = pPointColl.Point(i)

                'Vérifier si le point est en erreur
                If m_ErreurFeatureAdjacence.Contains(pPointA.ID.ToString) Then
                    'Définir l'erreur A
                    oErreurA = CType(m_ErreurFeatureAdjacence.Item(pPointA.ID.ToString), Structure_Erreur)

                    'Trouver l'élément adjacent
                    For j = i + 1 To pPointColl.PointCount - 1
                        'Définir le point d'adjacence correspondant
                        pPointB = pPointColl.Point(j)

                        'Projeter les points
                        pPointA.Project(m_SpatialReferenceProj)
                        pPointB.Project(m_SpatialReferenceProj)
                        'Interface pour calculer la distance d'adjacence
                        pProxOp = CType(pPointA, IProximityOperator)
                        'Calculer la distance d'adjacence
                        dDistance = pProxOp.ReturnDistance(pPointB)

                        'Sortir si la distance est plus grande que la tolérance d'adjacence
                        'Debug.Print(dDistance.ToString("0.0#######"))
                        If dDistance > m_TolAdjacenceOri Then Exit For

                        'Vérifier si le point correspondant est aussi en erreur
                        If m_ErreurFeatureAdjacence.Contains(pPointB.ID.ToString) Then
                            'Définir l'erreur B
                            oErreurB = CType(m_ErreurFeatureAdjacence.Item(pPointB.ID.ToString), Structure_Erreur)

                            'Vérifier si c'est la même classe
                            If oErreurA.FeatureA.Class.AliasName = oErreurB.FeatureA.Class.AliasName Or m_ClasseDifferente Then
                                'Ajouter le lien d'adjacence dans l'erreur
                                oErreurA.Description = "Sommet d'élément adjacent à déplacer (" + dDistance.ToString("0.0##") + "<=" + m_TolAdjacenceOri.ToString("0.0##") + ")"
                                oErreurA.IdentifiantB = oErreurB.IdentifiantA
                                oErreurA.FeatureB = oErreurB.FeatureA
                                oErreurA.PointB = oErreurB.PointA
                                oErreurA.Distance = dDistance

                                'Retirer l'élément A d'adjacence isolé en erreur
                                m_ErreurFeatureAdjacence.Remove(pPointA.ID.ToString)

                                'Retirer l'élément B d'adjacence isolé en erreur
                                m_ErreurFeatureAdjacence.Remove(pPointB.ID.ToString)

                                'Ajouter la structure d'erreur dans la collection
                                m_ErreurFeatureAdjacence.Add(oErreurA)

                                'Sortir de la boucle lorsque l'élément est trouvé
                                Exit For
                            End If
                        End If
                    Next
                End If

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit For
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
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
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Message = "Valider l'adjacence des éléments de la liste des points d'adjacence ..."
            'Afficher la barre de progression
            InitBarreProgression(0, m_ListeElementPointAdjacent.Count, pTrackCancel)

            'Traiter tous les points d'adjacence
            For i = 1 To m_ListeElementPointAdjacent.Count
                'Définir le point d'adjacence
                pPoint = CType(m_ListePointAdjacence.Geometry(i - 1), IPoint)

                'Si le point d'adjacence n'a pas déjà une erreur
                If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then
                    'Définir la liste des sommets d'éléments adjacents
                    qElementColl = CType(m_ListeElementPointAdjacent.Item(i), Collection)

                    'Traiter tous les éléments adjacents pour un point d'adjacence
                    For j = 1 To qElementColl.Count
                        'Définir l'élément
                        qElementTraiter = CType(m_ListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)
                        pFeatureA = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                        pDatasetA = CType(pFeatureA.Class, IDataset)

                        'Définir la valeur de l'identifiant
                        sIdentifiantA = fsIdentifiantDecoupage(pFeatureA, m_IdentifiantDecoupage)

                        'Initialiser le nombre d'éléments adjacents
                        nNbAdjacent = 0
                        'Vérifier l'adjacence de tous les éléments
                        For k = 1 To qElementColl.Count
                            'Définir l'élément adjacent
                            qElementTraiter = CType(m_ListeElementTraiter.Item(qElementColl.Item(k)), Structure_Element_Traiter)
                            pFeatureB = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)
                            pDatasetB = CType(pFeatureB.Class, IDataset)

                            'Définir la valeur de l'identifiant adjacent
                            sIdentifiantB = fsIdentifiantDecoupage(pFeatureB, m_IdentifiantDecoupage)

                            'Vérifier si ce n'est pas le même OBJECTID ou si ce n'est pas la même classe, ou si ce n'est pas la même Géodatabase
                            If pFeatureA.OID <> pFeatureB.OID Or pFeatureA.Class.AliasName <> pFeatureB.Class.AliasName Or pDatasetA.Workspace.PathName <> pDatasetB.Workspace.PathName Then
                                'Vérifier si c'est la même classe ou si si on permet une classe différente
                                If pFeatureA.Class.AliasName = pFeatureB.Class.AliasName Or m_ClasseDifferente = True Then
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
                            If m_AdjacenceUnique Then
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
                                If Not m_ErreurFeatureAdjacence.Contains(sClef) Then m_ErreurFeatureAdjacence.Add(oErreur, sClef)
                                'Conserver l'dentifiant du point d'adjacence en erreur
                                If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                            End If

                            'Vérifier si l'élément possède un seul élément adjacent
                        ElseIf nNbAdjacent = 1 Then
                            'Vérifier si plusieurs éléments adjacents et si seulement l'adjacence unique est permise
                            If qElementColl.Count > 2 And m_AdjacenceUnique Then
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
                                If Not m_ErreurFeatureAdjacence.Contains(sClef) Then m_ErreurFeatureAdjacence.Add(oErreur, sClef)
                                'Conserver l'dentifiant du point d'adjacence en erreur
                                If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
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
                                        If Not m_ErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then m_ErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                        'Conserver l'dentifiant du point d'adjacence en erreur
                                        If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)

                                        'Si le point d'adjacence ne correspond pas une extrémité de la ligne et si doit tenir compte des identifiants
                                    ElseIf m_SansIdentifiant = False Then
                                        'Conserver l'erreur dans une structure
                                        oErreur = New Structure_Erreur
                                        oErreur.Description = "Aucun élément adjacent mais n'est pas une extrémité de ligne"
                                        oErreur.IdentifiantA = sIdentifiantA
                                        oErreur.PointA = pPoint
                                        oErreur.FeatureA = pFeatureA
                                        oErreur.ValeurA = nNbAdjacent.ToString
                                        oErreur.Distance = 0

                                        'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                        If Not m_ErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then m_ErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                        'Conserver l'dentifiant du point d'adjacence en erreur
                                        If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
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
                                    If Not m_ErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then m_ErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                    'Conserver l'dentifiant du point d'adjacence en erreur
                                    If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                                End If

                                'Si deux éléments
                            ElseIf qElementColl.Count = 2 Then
                                'Si on doit tenir compte des identifiants
                                If m_SansIdentifiant = False Then
                                    'Conserver l'erreur dans une structure
                                    oErreur = New Structure_Erreur
                                    oErreur.Description = "Aucun élément adjacent mais un autre élément est présent (Nb éléments=" & qElementColl.Count.ToString & ")"
                                    oErreur.IdentifiantA = sIdentifiantA
                                    oErreur.PointA = pPoint
                                    oErreur.FeatureA = pFeatureA
                                    oErreur.ValeurA = nNbAdjacent.ToString
                                    oErreur.Distance = 0

                                    'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                    If Not m_ErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then m_ErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                    'Conserver l'dentifiant du point d'adjacence en erreur
                                    If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                                End If

                                'Si plusieurs éléments
                            ElseIf qElementColl.Count > 2 Then
                                'Si on doit tenir compte des identifiants ou si seulement l'adjacence unique est permise 
                                If m_SansIdentifiant = False Or m_AdjacenceUnique Then
                                    'Conserver l'erreur dans une structure
                                    oErreur = New Structure_Erreur
                                    oErreur.Description = "Aucun élément adjacent mais plusieurs éléments présents (Nb éléments=" & qElementColl.Count.ToString & ")"
                                    oErreur.IdentifiantA = sIdentifiantA
                                    oErreur.PointA = pPoint
                                    oErreur.FeatureA = pFeatureA
                                    oErreur.ValeurA = nNbAdjacent.ToString
                                    oErreur.Distance = 0

                                    'Ajouter la structure d'erreur dans la collection avec sa clef de recherche
                                    If Not m_ErreurFeatureAdjacence.Contains(pPoint.ID.ToString) Then m_ErreurFeatureAdjacence.Add(oErreur, pPoint.ID.ToString)
                                    'Conserver l'dentifiant du point d'adjacence en erreur
                                    If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                                End If
                            End If
                        End If
                    Next
                End If

                'Récupération de la mémoire disponible
                GC.Collect()

                'Vérifier si un Cancel a été effectué
                If pTrackCancel.Continue = False Then Exit For
            Next

            'Cacher la barre de progression
            If pTrackCancel.Progressor IsNot Nothing Then pTrackCancel.Progressor.Hide()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
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
            For i = 1 To m_AttributAdjacence.Count
                'Définir le nom de l'attribt à traiter
                sNomAttribut = CType(m_AttributAdjacence.Item(i), String)

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
                    If Not m_ErreurFeatureAttribut.Contains(sClef) Then
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
                        m_ErreurFeatureAttribut.Add(oErreur, sClef)
                        'Conserver l'dentifiant du point d'adjacence en erreur
                        If Not m_ListeErreurPointAdjacent.Contains(pPoint.ID.ToString) Then m_ListeErreurPointAdjacent.Add(pPoint.ID, pPoint.ID.ToString)
                    End If
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
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
                If pHitTest.HitTest(pPoint, m_TolRecherche, esriGeometryHitPartType.esriGeometryPartVertex, _
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
                    If pProxOp.ReturnDistance(pFeatureB.Shape) <= m_TolRecherche Then
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
                    If pProxOp.ReturnDistance(pFeatureB.Shape) <= m_TolRecherche Then
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
    ''' Routine qui permet de transformer les tolrérances de recherche et d'adjacence en géographique à partir d'un enveloppe d'une de zone travail.
    '''</summary>
    '''
    '''<param name="pEnvelope"> Envelope correspondant à la zone de travail.</param>
    ''' 
    Public Sub TransformerTolerances(ByVal pEnvelope As IEnvelope)
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
                m_SpatialReferenceProj = pSpatialRefFact.CreateSpatialReference(3979)

                'Projeter la limite dans la référence spatiale LCC NAD83 CSRS:3979
                pEnvelopeProj.Project(m_SpatialReferenceProj)

                'Définir la distance projetée
                pProximityOp = CType(pEnvelopeProj.LowerLeft, IProximityOperator)
                dDistanceProj = pProximityOp.ReturnDistance(pEnvelopeProj.UpperRight)

                'Définir la distance géographique
                pProximityOp = CType(pEnvelope.LowerLeft, IProximityOperator)
                dDistanceGeog = pProximityOp.ReturnDistance(pEnvelope.UpperRight)

                'Redéfinir les tolérances
                m_TolRecherche = (m_TolRechercheOri * dDistanceGeog) / dDistanceProj
                m_TolAdjacence = (m_TolAdjacenceOri * dDistanceGeog) / dDistanceProj
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

    '''<summary>
    ''' Routine qui permet de corriger toutes les erreurs de précision.
    ''' On déplace les sommets de tous les éléments en erreur présents vers le point d'adjacence correspondant.
    ''' 
    '''<param name="bFenetre">Indique si on doit corriger seulement les erreurs présentes dans la fenêtre graphique.</param>
    ''' 
    '''</summary>
    ''' 
    Public Sub CorrigerErreurPrecision(Optional ByVal bFenetre As Boolean = False)
        'Déclarer les variables de travail
        Dim qElementTraiter As Structure_Element_Traiter = Nothing  'Objet contenant l'information de l'élément.
        Dim oErreur As Structure_Erreur = Nothing   'Structure contenant une erreur de précision
        Dim pFeature As IFeature = Nothing          'Interface ESRI contenant l'élément en erreur
        Dim pEditor As IEditor = Nothing            'Interface ESRI pour effectuer l'édition des éléments.
        Dim pPointA As IPoint = Nothing             'Interface contenant le point en erreur
        Dim pPointB As IPoint = Nothing             'Interface contenant le point en erreur
        Dim i As Integer = Nothing                  'Compteur

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()
            End If

            'Traiter toutes les erreurs de précision
            For i = 1 To m_ErreurFeaturePrecision.Count
                'Définir la structure d'erreur
                oErreur = CType(m_ErreurFeaturePrecision.Item(i), Structure_Erreur)

                'Définir l'élément

                'Définir l'élément en erreur
                pFeature = oErreur.FeatureA
                qElementTraiter = CType(m_ListeElementTraiter(oErreur.ElementA), Structure_Element_Traiter)
                pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)

                'Définir le point en erreur
                pPointA = oErreur.PointB

                'Définir le point d'adjacence
                pPointB = oErreur.PointA

                'S'il faut vérifier si le point en erreur est à l'intérieur de la fenêtre
                If bFenetre Then
                    'Vérifier si le point en erreur est à l'intérieur de la fenêtre
                    If pPointA.X > m_MxDocument.ActiveView.Extent.XMin And pPointA.X < m_MxDocument.ActiveView.Extent.XMax _
                    And pPointA.Y > m_MxDocument.ActiveView.Extent.YMin And pPointA.Y < m_MxDocument.ActiveView.Extent.YMax Then
                        'Déplacer le point de l'élément en erreur vers le point d'adjacence
                        Call DeplacerSommet(pFeature, pPointA, pPointB, m_TolRecherche)
                    End If

                    'Sinon
                Else
                    'Déplacer le point de l'élément en erreur vers le point d'adjacence
                    Call DeplacerSommet(pFeature, pPointA, pPointB, m_TolRecherche)
                End If
            Next

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Terminer l'opération UnDo
                pEditor.StopOperation("Corriger erreur précision")
            End If
            'Désactiver l'interface
            pEditor = Nothing

            'Rafraîchir l'affichage
            m_MxDocument.FocusMap.ClearSelection()
            m_MxDocument.ActiveView.Refresh()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            qElementTraiter = Nothing
            oErreur = Nothing
            pFeature = Nothing
            pEditor = Nothing
            pPointA = Nothing
            pPointB = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de corriger toutes les erreurs d'adjacence.
    ''' On déplace les sommets de tous les éléments en erreur présents vers le point d'adjacence correspondant.
    ''' 
    '''<param name="bFenetre">Indique si on doit corriger seulement les erreurs présentes dans la fenêtre graphique.</param>
    ''' 
    '''</summary>
    ''' 
    Public Sub CorrigerErreurAdjacence(Optional ByVal bFenetre As Boolean = False)
        'Déclarer les variables de travail
        Dim pEnumFeature As IEnumFeature = Nothing      'Interface ESRI utilisé pour extraire les éléments sélectionnés.
        Dim pEnumFeatureSet As IEnumFeatureSetup = Nothing  'Interface ESRI utilisé pour extraire les éléments sélectionnés.
        Dim qElementTraiter As Structure_Element_Traiter = Nothing 'Contient un élément à traiter.
        Dim oErreur As Structure_Erreur = Nothing   'Structure contenant une erreur d'adjacence
        Dim qElementColl As Collection = Nothing    'Objet qui contient la liste des éléments adjacents
        Dim pFeature As IFeature = Nothing          'Interface ESRI contenant l'élément à déplacer
        Dim pEditor As IEditor = Nothing            'Interface ESRI pour effectuer l'édition des éléments.
        Dim i As Integer = Nothing                  'Compteur
        Dim j As Integer = Nothing                  'Compteur

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()
            End If

            'Traiter tous les points de découpage
            For i = 1 To m_ErreurFeatureAdjacence.Count
                'Définir l'erreur A
                oErreur = CType(m_ErreurFeatureAdjacence.Item(i), Structure_Erreur)

                'Vérifier si l'erreur est un élément isolé
                If oErreur.PointB Is Nothing Then
                    'Sélectionner tous les éléments au point d'adjacence
                    m_MxDocument.FocusMap.SelectByShape(oErreur.PointA, Nothing, False)

                    'Si plus d'un élément est trouvé
                    If m_MxDocument.FocusMap.SelectionCount > 1 Then
                        'Interface pour extraire les éléments sélectionnés
                        pEnumFeature = CType(m_MxDocument.ActiveView.FocusMap.FeatureSelection, IEnumFeature)
                        'Interface pour extraire toutes les valeurs d'attributs
                        pEnumFeatureSet = CType(pEnumFeature, IEnumFeatureSetup)

                        'Extraire le premier élément
                        pEnumFeature.Reset()
                        pFeature = pEnumFeature.Next

                        'Traiter tous les éléments
                        Do Until pFeature Is Nothing
                            'Déplacer le point de l'élément en erreur vers le point d'adjacence
                            Call DeplacerSommet(pFeature, oErreur.PointA, oErreur.PointA, m_TolRecherche)

                            'Extraire le prochain élément
                            pFeature = pEnumFeature.Next
                        Loop
                    End If

                    'Vérifier si l'erreur n'est pas un élément isolé
                Else
                    'Vérifier s'il faut vérifier si le point en erreur est à l'intérieur de la fenêtre
                    If bFenetre Then
                        'Vérifier si le point en erreur est à l'intérieur de la fenêtre
                        If oErreur.PointA.X > m_MxDocument.ActiveView.Extent.XMin And oErreur.PointA.X < m_MxDocument.ActiveView.Extent.XMax _
                        And oErreur.PointA.Y > m_MxDocument.ActiveView.Extent.YMin And oErreur.PointA.Y < m_MxDocument.ActiveView.Extent.YMax Then
                            'Vérifier si le point d'adjacence est un sommet virtuel car il ne faut pas le déplacer
                            If oErreur.PointA.ID < 0 Then
                                'Définir la liste des éléments du point d'adjacence B
                                qElementColl = CType(m_ListeElementPointAdjacent.Item(oErreur.PointB.ID.ToString), Collection)

                                'Traiter tous les éléments du point d'adjacence B
                                For j = 1 To qElementColl.Count
                                    'Définir l'élément à traiter
                                    'pFeature = CType(qElementColl.Item(j), IFeature)
                                    qElementTraiter = CType(m_ListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)
                                    pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)

                                    'Déplacer le point de l'élément en erreur vers le point d'adjacence
                                    Call DeplacerSommet(pFeature, oErreur.PointB, oErreur.PointA, m_TolRecherche)
                                Next

                                'Si ce n'est pas un sommet virtuel
                            Else
                                'Définir la liste des élément du point d'adjacence A
                                qElementColl = CType(m_ListeElementPointAdjacent.Item(oErreur.PointA.ID.ToString), Collection)

                                'Traiter tous les éléments du point d'adjacence A
                                For j = 1 To qElementColl.Count
                                    'Définir l'élément à traiter
                                    'pFeature = CType(qElementColl.Item(j), IFeature)
                                    qElementTraiter = CType(m_ListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)
                                    pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)

                                    'Déplacer le point de l'élément en erreur vers le point d'adjacence
                                    Call DeplacerSommet(pFeature, oErreur.PointA, oErreur.PointB, m_TolRecherche)
                                Next
                            End If
                        End If

                        'Si on tient pas compte de la fenêtre
                    Else
                        'Vérifier si le point d'adjacence est un sommet virtuel car il ne faut pas le déplacer
                        If oErreur.PointA.ID < 0 Then
                            'Définir la liste des éléments du point d'adjacence B
                            qElementColl = CType(m_ListeElementPointAdjacent.Item(oErreur.PointB.ID.ToString), Collection)

                            'Traiter tous les éléments du point d'adjacence B
                            For j = 1 To qElementColl.Count
                                'Définir l'élément à traiter
                                'pFeature = CType(qElementColl.Item(j), IFeature)
                                qElementTraiter = CType(m_ListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)
                                pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)

                                'Déplacer le point de l'élément en erreur vers le point d'adjacence
                                Call DeplacerSommet(pFeature, oErreur.PointB, oErreur.PointA, m_TolRecherche)
                            Next

                            'Si ce n'est pas un sommet virtuel
                        Else
                            'Définir la liste des élément du point d'adjacence A
                            qElementColl = CType(m_ListeElementPointAdjacent.Item(oErreur.PointA.ID.ToString), Collection)

                            'Traiter tous les éléments du point d'adjacence A
                            For j = 1 To qElementColl.Count
                                'Définir l'élément à traiter
                                'pFeature = CType(qElementColl.Item(j), IFeature)
                                qElementTraiter = CType(m_ListeElementTraiter.Item(qElementColl.Item(j)), Structure_Element_Traiter)
                                pFeature = qElementTraiter.FeatureClass.GetFeature(qElementTraiter.OID)

                                'Déplacer le point de l'élément en erreur vers le point d'adjacence
                                Call DeplacerSommet(pFeature, oErreur.PointA, oErreur.PointB, m_TolRecherche)
                            Next
                        End If
                    End If
                End If
            Next

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Terminer l'opération UnDo
                pEditor.StopOperation("Corriger erreur adjacence")
            End If
            'Désactiver l'interface
            pEditor = Nothing

            'Rafraîchir l'affichage
            m_MxDocument.FocusMap.ClearSelection()
            m_MxDocument.ActiveView.Refresh()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEnumFeature = Nothing
            pEnumFeatureSet = Nothing
            qElementTraiter = Nothing
            oErreur = Nothing
            qElementColl = Nothing
            pFeature = Nothing
            pEditor = Nothing
            i = Nothing
            j = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de corriger toutes les erreurs d'attribut.
    ''' On remplace la valeur de l'attribut en erreur.
    ''' 
    '''<param name="bFenetre">Indique si on doit corriger seulement les erreurs présentes dans la fenêtre graphique.</param>
    ''' 
    '''</summary>
    ''' 
    Public Sub CorrigerErreurAttribut(Optional ByVal bFenetre As Boolean = False)
        'Déclarer les variables de travail
        Dim oErreur As Structure_Erreur = Nothing   'Objet contenant toute l'information pour corriger l'erreur d'attribut
        Dim pEditor As IEditor = Nothing            'Interface ESRI pour effectuer l'édition des éléments.
        Dim pWrite As IFeatureClassWrite = Nothing  'Interface qui permet d'écrire la correction
        Dim i As Integer = Nothing                  'Compteur

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()
            End If

            'Traiter toutes les erreurs d'attribut
            For i = 1 To m_ErreurFeatureAttribut.Count
                'Définir le point en erreur
                oErreur = CType(m_ErreurFeatureAttribut.Item(i), Structure_Erreur)

                'Vérifier s'il faut vérifier si le point en erreur est à l'intérieur de la fenêtre
                If bFenetre Then
                    'Vérifier si le point en erreur est à l'intérieur de la fenêtre
                    If oErreur.PointA.X > m_MxDocument.ActiveView.Extent.XMin And oErreur.PointA.X < m_MxDocument.ActiveView.Extent.XMax _
                    And oErreur.PointA.Y > m_MxDocument.ActiveView.Extent.YMin And oErreur.PointA.Y < m_MxDocument.ActiveView.Extent.YMax Then
                        'Vérifier si on doit corriger la valeur de l'élément A
                        If oErreur.ValeurA <> "." And oErreur.ValeurB <> "" And oErreur.ValeurB <> "." Then
                            'Corriger la valeur de l'élément A en fonction de la valeur de l'élément B
                            oErreur.FeatureA.Value(oErreur.PosAttA) = oErreur.ValeurB
                            'Conserver la correction
                            pWrite = CType(oErreur.FeatureA.Class, IFeatureClassWrite)
                            pWrite.WriteFeature(oErreur.FeatureA)

                            'Vérifier si on doit corriger la valeur de l'élément B
                        ElseIf oErreur.ValeurB <> "." And oErreur.ValeurA <> "" And oErreur.ValeurA <> "." Then
                            'Corriger la valeur de l'élément B en fonction de la valeur de l'élément A
                            oErreur.FeatureB.Value(oErreur.PosAttB) = oErreur.ValeurA
                            'Conserver la correction
                            pWrite = CType(oErreur.FeatureB.Class, IFeatureClassWrite)
                            pWrite.WriteFeature(oErreur.FeatureB)
                        End If
                    End If
                Else
                    'Vérifier si on doit corriger la valeur de l'élément A
                    If oErreur.ValeurA <> "." And oErreur.ValeurB <> "" And oErreur.ValeurB <> "." Then
                        'Corriger la valeur de l'élément A en fonction de la valeur de l'élément B
                        oErreur.FeatureA.Value(oErreur.PosAttA) = oErreur.ValeurB
                        'Conserver la correction
                        pWrite = CType(oErreur.FeatureA.Class, IFeatureClassWrite)
                        pWrite.WriteFeature(oErreur.FeatureA)

                        'Vérifier si on doit corriger la valeur de l'élément B
                    ElseIf oErreur.ValeurB <> "." And oErreur.ValeurA <> "" And oErreur.ValeurA <> "." Then
                        'Corriger la valeur de l'élément B en fonction de la valeur de l'élément A
                        oErreur.FeatureB.Value(oErreur.PosAttB) = oErreur.ValeurA
                        'Conserver la correction
                        pWrite = CType(oErreur.FeatureB.Class, IFeatureClassWrite)
                        pWrite.WriteFeature(oErreur.FeatureB)
                    End If
                End If
            Next

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Terminer l'opération UnDo
                pEditor.StopOperation("Corriger erreur attribut")
            End If
            'Désactiver l'interface
            pEditor = Nothing

            'Rafraîchir l'affichage
            m_MxDocument.FocusMap.ClearSelection()
            m_MxDocument.ActiveView.Refresh()

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            oErreur = Nothing
            pEditor = Nothing
            pWrite = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de déplacer le sommet de l'élément en erreur du pointA vers le pointB selon la tolérence active. 
    '''</summary>
    '''
    '''<param name="pFeature">Interface ESRI contenant l'élément à corriger.</param>
    '''<param name="pPointA">Interface ESRI contenant le point de départ.</param>
    '''<param name="pPointB">Interface ESRI contenant le point d'arrivé.</param>
    '''<param name="dTol">Contient la tolérence de recherche.</param>
    ''' 
    Public Sub DeplacerSommet(ByRef pFeature As IFeature, ByVal pPointA As IPoint, ByVal pPointB As IPoint, Optional ByVal dTol As Double = 5.0)
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie de l'élément.
        Dim pGeodataset As IGeoDataset = Nothing        'Interface contenant la référence spatial de la classe.
        Dim pGeometryDef As IGeometryDef = Nothing      'Interface ESRI qui permet de vérifier la présence du Z et du M.
        Dim pWrite As IFeatureClassWrite = Nothing      'Interface qui permet d'écrire un élément.
        Dim bModif As Boolean = False                   'Indique si une modification a été effectuée.

        Try
            'Interface contenant la référence spatiale
            pGeodataset = CType(pFeature.Class, IGeoDataset)

            'Définir la géométrie de l'élément
            pGeometry = pFeature.ShapeCopy

            'Déplacer et vérifier si le sommet a été déplacé
            If DeplacerSommetGeometrie(pGeometry, pPointA, pPointB, dTol, True, True) Then
                'Vérifier si la géométrie est vide
                If pGeometry.IsEmpty Then
                    'Détruire l'élément
                    pFeature.Delete()

                    'si la géométrie n'est pas vide
                Else
                    'Indiquer qu'il y a eu une modification
                    bModif = True
                    'Projet la géométrie
                    pGeometry.Project(pGeodataset.SpatialReference)

                    'Interface pour vérifier la présence du Z et du M
                    pGeometryDef = RetournerGeometryDef(CType(pFeature.Class, IFeatureClass))
                    'Vérifier la présence du Z
                    If pGeometryDef.HasZ Then
                        'Traiter le Z
                        Call TraiterZ(pGeometry)
                    End If
                    'Vérifier la présence du M
                    If pGeometryDef.HasM Then
                        'Traiter le M
                        Call TraiterM(pGeometry)
                    End If

                    'Changer la géométrie de l'élément
                    pFeature.Shape = pGeometry
                    Try
                        'Sauver la modification
                        pWrite = CType(pFeature.Class, IFeatureClassWrite)
                        pWrite.WriteFeature(pFeature)
                    Catch ex As Exception
                        'Message d'erreur
                        Err.Raise(vbObjectError + 1, "", "ERREUR : Vous devez enlever la classe de topologie dans votre BD !")
                    End Try
                End If
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pGeodataset = Nothing
            pGeometryDef = Nothing
            pWrite = Nothing
            bModif = Nothing
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
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pDataset = Nothing
            nPosAtt = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de retourner le buffer de la géométrie spécifiée en paramètre selon une largeur.
    ''' 
    ''' Si la géométrie est vide, le buffer sera vide.
    ''' Si la largeur est nulle, la largeur changée pour la valeur 0.001.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI de la Géométrie à traiter.</param>
    '''<param name="dLargeur">Largeur utilisée pour créer le buffer de la géométrie.</param>
    ''' 
    '''<returns>La fonction va retourner un "IPolygon". Sinon "Nothing".</returns>
    '''
    Public Function fpBufferGeometrie(ByVal pGeometry As IGeometry, ByVal dLargeur As Double) As IPolygon
        'Déclarer les variables de travail
        Dim pPolygon As IPolygon = Nothing              'Interface contenant le buffer de la géométrie
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface utilisé pour créer le buffer
        Dim pMultipoint As IMultipoint = Nothing        'Interface contenant le point
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter le point dans le Multipoint
        Dim i As Integer = Nothing                      'Compteur

        'Définir la valeur par défaut
        fpBufferGeometrie = Nothing

        Try
            'Vérifier si la largeur est nulle
            If dLargeur <= 0 Then dLargeur = 0.001

            'Vérifier si la géométrie est vide ou qu'il n'y a pas de largeur
            If pGeometry.IsEmpty Then
                'Créer un polygon vide
                pPolygon = CType(New Polygon, IPolygon)
                pPolygon.SpatialReference = pGeometry.SpatialReference

                'Vérifier si la géométrie est un Bag
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                'Créer un polygon vide
                pPolygon = CType(New Polygon, IPolygon)
                pPolygon.SpatialReference = pGeometry.SpatialReference
                'Traiter toutes les composantes de géométries
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Créer le buffer
                    pGeomColl = CType(fpBufferGeometrie(pGeomColl.Geometry(i), dLargeur), IGeometryCollection)
                    'Ajouter le polygon précédent
                    pGeomColl.AddGeometry(pPolygon)
                Next
                'Définir le buffer
                pPolygon = CType(pGeomColl, IPolygon)

                'Vérifier si la géométrie est un point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Créer un multipoint vide
                pMultipoint = CType(New Multipoint, IMultipoint)
                pMultipoint.SpatialReference = pGeometry.SpatialReference
                'Ajouter le point dans le multipoint
                pGeomColl = CType(pMultipoint, IGeometryCollection)
                pGeomColl.AddGeometry(pGeometry)
                'Interface pour créer le buffer
                pTopoOp = CType(pMultipoint, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
                'Créer le buffer
                pPolygon = CType(pTopoOp.Buffer(dLargeur), IPolygon)

                'Pour les autres types
            Else
                'Simplifier la géométrie
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
                'Créer le buffer
                pPolygon = CType(pTopoOp.Buffer(dLargeur), IPolygon)
            End If

            'Retourner la même géométrie
            fpBufferGeometrie = pPolygon

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pPolygon = Nothing
            pTopoOp = Nothing
            pMultipoint = Nothing
            pGeomColl = Nothing
            i = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter le Z d'une Géométrie.
    '''</summary>
    '''
    '''<param name=" pGeometry "> Interface contenant la géométrie à traiter.</param>
    '''<param name=" dZ "> Contient la valeur du Z.</param>
    '''
    Public Sub TraiterZ(ByRef pGeometry As IGeometry, Optional ByVal dZ As Double = 0)
        'Déclarer les variables de travail
        Dim pZAware As IZAware = Nothing                'Interface ESRI utilisée pour traiter le Z.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface qui permet d'accéder aux géométries
        Dim pPointColl As IPointCollection = Nothing    'Interface qui permet d'accéder aux sommets de la géométrie
        Dim pPoint As IPoint = Nothing                  'Interface qui permet de modifier le Z
        Dim pZ As IZ = Nothing                          'Interface utilisé pour calculer le Z invalide
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Interface pour traiter le Z
            pZAware = CType(pGeometry, IZAware)

            'Vérifier la présence du Z
            If pGeometry.SpatialReference.HasZPrecision Then
                'Ajouter le 3D
                pZAware.ZAware = True
                'Vérifier si on doit corriger le Z
                If pZAware.ZSimple = False Then
                    'Vérifier si la géométrie est un Point
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Définir le point
                        pPoint = CType(pGeometry, IPoint)
                        'Interface pour traiter le Z
                        pZAware = CType(pPoint, IZAware)
                        'Vérifier si on doit corriger le Z
                        If pZAware.ZSimple = False Then
                            'Définir le Z du point
                            pPoint.Z = dZ
                        End If

                        'Vérifier si la géométrie est un Bag
                    ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                        'Interface utilisé pour accéder aux géométries
                        pGeomColl = CType(pGeometry, IGeometryCollection)
                        'Traiter toutes les géométries
                        For i = 0 To pGeomColl.GeometryCount - 1
                            'Traiter le Z
                            Call TraiterZ(pGeomColl.Geometry(i), dZ)
                        Next

                        'Vérifier si la géométrie est un Multipoint,une Polyline ou un Polygon
                    Else
                        Try
                            'Interface pour corriger le Z par interpollation
                            pZ = CType(pGeometry, IZ)
                            'Corriger le Z invalide par interpollation
                            pZ.CalculateNonSimpleZs()
                        Catch ex As Exception
                            'On ne fait rien
                        End Try

                        'Vérifier si on doit corriger le Z
                        If pZAware.ZSimple = False Then
                            'Interface utilisé pour accéder aux sommets de la géométrie
                            pPointColl = CType(pGeometry, IPointCollection)
                            'Traiter tous les sommets de la géométrie
                            For i = 0 To pPointColl.PointCount - 1
                                'Interface pour définir le Z
                                pPoint = pPointColl.Point(i)
                                'Interface pour traiter le Z
                                pZAware = CType(pPoint, IZAware)
                                'Vérifier si on doit corriger le Z
                                If pZAware.ZSimple = False Then
                                    'Définir le Z du point
                                    pPoint.Z = dZ
                                End If
                                'Conserver les modifications
                                pPointColl.UpdatePoint(i, pPoint)
                            Next
                        End If
                    End If
                End If

                'Si aucun Z
            Else
                'Enlever le 3D
                pZAware.ZAware = True
                pZAware.DropZs()
                pZAware.ZAware = False
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pZAware = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pPoint = Nothing
            pZ = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de traiter le M d'une Géométrie.
    '''</summary>
    '''
    '''<param name=" pGeometry "> Interface contenant la géométrie à traiter.</param>
    '''<param name=" dM "> Contient la valeur du M.</param>
    '''
    Public Sub TraiterM(ByRef pGeometry As IGeometry, Optional ByVal dM As Double = 0)
        'Déclarer les variables de travail
        Dim pMAware As IMAware = Nothing                'Interface ESRI utilisée pour traiter le M.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface qui permet d'accéder aux géométries
        Dim pPointColl As IPointCollection = Nothing    'Interface qui permet d'accéder aux sommets de la géométrie
        Dim pPoint As IPoint = Nothing                  'Interface qui permet de modifier le M
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Interface pour traiter le M
            pMAware = CType(pGeometry, IMAware)

            'Vérifier la présence du M
            If pGeometry.SpatialReference.HasMPrecision Then
                'Ajouter le M
                pMAware.MAware = True
                'Corriger le M au besoin
                If pMAware.MSimple = False Then
                    'Vérifier si la géométrie est un Point
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Définir le point
                        pPoint = CType(pGeometry, IPoint)
                        'Interface pour traiter le M
                        pMAware = CType(pPoint, IMAware)
                        'Vérifier si on doit corriger le M
                        If pMAware.MSimple = False Then
                            'Définir le Z du point
                            pPoint.M = dM
                        End If

                        'Vérifier si la géométrie est un Bag
                    ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                        'Interface utilisé pour accéder aux géométries
                        pGeomColl = CType(pGeometry, IGeometryCollection)
                        'Traiter toutes les géométries
                        For i = 0 To pGeomColl.GeometryCount - 1
                            'Traiter le M
                            Call TraiterM(pGeomColl.Geometry(i), dM)
                        Next

                        'Vérifier si la géométrie est un Multipoint,une Polyline ou un Polygon
                    Else
                        'Interface utilisé pour accéder aux sommets de la géométrie
                        pPointColl = CType(pGeometry, IPointCollection)
                        'Traiter tous les sommets de la géométrie
                        For i = 0 To pPointColl.PointCount - 1
                            'Interface pour définir le M
                            pPoint = pPointColl.Point(i)
                            'Interface pour traiter le M
                            pMAware = CType(pPoint, IMAware)
                            'Vérifier si on doit corriger le M
                            If pMAware.MSimple = False Then
                                'Définir le Z du point
                                pPoint.M = dM
                            End If
                            'Conserver les modifications
                            pPointColl.UpdatePoint(i, pPoint)
                        Next
                    End If
                End If

                'Si aucun M
            Else
                'Enlever le M
                pMAware.MAware = True
                pMAware.DropMs()
                pMAware.MAware = False
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pMAware = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pPoint = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de transformer une géométrie point, ligne ou surface en MultiPoint. 
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à convertir en Multipoint.</param>
    ''' 
    '''<returns>La fonction va retourner un "IMultiPoint" contenant un ou plusieurs points. Sinon "Nothing".</returns>
    '''
    Public Function fpConvertirEnMultiPoint(ByVal pGeometry As IGeometry) As IMultipoint
        'Déclarer les variables de travail
        Dim pMultiPoint As IMultipoint = Nothing        'Interface contenant le nouveau Multipoint
        Dim pPointColl As IPointCollection = Nothing    'Interface utilisé pour ajouter le point dans le Multipoint

        'Définir la valeur par défaut
        fpConvertirEnMultiPoint = Nothing

        Try
            'Créer un nouveau MultiPoint vide
            pMultiPoint = CType(New Multipoint, IMultipoint)

            'Définir la référence spatiale du MultiiPoint
            pMultiPoint.SpatialReference = pGeometry.SpatialReference

            'Interface pour ajouter le point dans le MultiPoint
            pPointColl = CType(pMultiPoint, IPointCollection)

            'Vérifier si la géométrie est un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Ajouter le point dans le MultiPoint
                pPointColl.AddPoint(CType(pGeometry, IPoint))

                'Si la géométrie n'est pas un point
            Else
                'Ajouter les points de la géométrie dans le MultiPoint
                pPointColl.AddPointCollection(CType(pGeometry, IPointCollection))
            End If

            'Retourner le nouveau polygone
            fpConvertirEnMultiPoint = CType(pPointColl, IMultipoint)

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pPointColl = Nothing
            pMultiPoint = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de déplacer un sommet existant sur une géométrie en fonction d'un point, d'une tolérance de recherche
    ''' et d'un point de déplacement.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à traiter.</param>
    '''<param name="pPointA">Interface ESRI contenant le sommet à rechercher sur la géométrie selon une tolérance.</param> 
    '''<param name="pPointB">Interface ESRI contenant le sommet résultant du sommet trouvé.</param> 
    '''<param name="dTolerance">Contient la tolérance de recherche du sommet à rechercher sur la géométrie.</param>
    '''<param name="bSimplifier">Indiquer si on doit simplifier la géométrie après avoir déplacé le sommet.</param>
    '''
    '''<returns>TRUE pour indiquer si le sommet a été déplacé, FALSE sinon.</returns>
    '''
    Public Function DeplacerSommetGeometrie(ByRef pGeometry As IGeometry, ByVal pPointA As IPoint, ByVal pPointB As IPoint, _
    Optional ByVal dTolerance As Double = 0, Optional ByVal bSeulementUn As Boolean = False, _
    Optional ByVal bSimplifier As Boolean = True) As Boolean
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface utilisée pour extraire la limite d'une géométrie
        Dim pPath As IPath = Nothing                    'Interface ESRI utilisé pour vérifier si la partie de la géométrie est fermée
        Dim pClone As IClone = Nothing                  'Interface ESRI utilisé pour cloner une géométrie
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface ESRI utilisé pour accéder aux parties de la géométrie
        Dim pPointColl As IPointCollection = Nothing    'Interface ESRI utilisé pour détruire un sommet
        Dim pHitTest As IHitTest = Nothing              'Interface pour tester la présence du sommet recherché
        Dim pNewPoint As IPoint = Nothing               'Interface contenant le sommet trouvé
        Dim pProxOp As IProximityOperator = Nothing     'Interface qui permet de calculer la distance
        Dim dDistance As Double = Nothing               'Interface contenant la distance calculée entre le point de recherche et le sommet trouvé
        Dim nNumeroPartie As Integer = Nothing          'Numéro de partie trouvée
        Dim nNumeroSommet As Integer = Nothing          'Numéro de sommet de la partie trouvée
        Dim bCoteDroit As Boolean = Nothing             'Indiquer si le point trouvé est du côté droit de la géométrie
        Dim j As Integer = Nothing                      'Compteur

        Try
            'Initialiser la valeur de retour
            DeplacerSommetGeometrie = False

            'Projeter les sommets
            pPointA.Project(pGeometry.SpatialReference)
            pPointB.Project(pGeometry.SpatialReference)

            'Interface pour extraire le sommet et le numéro de sommet de proximité
            pHitTest = CType(pGeometry, IHitTest)

            'Vérifier s'il s'agit d'un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Interface pour calculer la distance
                pProxOp = CType(pPointA, IProximityOperator)

                'Vérifier la distance
                If pProxOp.ReturnDistance(pGeometry) <= dTolerance Then
                    'Interface pour cloner la géométrie
                    pClone = CType(pPointB, IClone)

                    'Redéfinir la géométrie
                    pGeometry = CType(pClone.Clone, IGeometry)

                    'Indiquer que le sommet a été déplacé
                    DeplacerSommetGeometrie = True
                End If

                'Vérifier s'il s'agit d'un multipoint
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Rechercher le point par rapport à chaque sommet de la géométrie
                If pHitTest.HitTest(pPointA, dTolerance, esriGeometryHitPartType.esriGeometryPartVertex, _
                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Interface pour détruire le sommet de proximité
                    pPointColl = CType(pGeometry, IPointCollection)

                    'Déplacer le sommet
                    pPointColl.UpdatePoint(nNumeroSommet, pPointB)

                    'Indiquer que le sommet a été déplacé
                    DeplacerSommetGeometrie = True
                End If

                'Vérifier s'il s'agit d'une ligne ou d'une surface
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Or _
            pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Rechercher le point par rapport à chaque sommet de la géométrie
                If pHitTest.HitTest(pPointA, dTolerance, esriGeometryHitPartType.esriGeometryPartVertex, _
                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Interface pour accéder aux parties de la géométrie
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Interface pour vérifier si la partie est fermée
                    pPath = CType(pGeomColl.Geometry(nNumeroPartie), IPath)

                    'Interface pour déplacer le sommet
                    pPointColl = CType(pPath, IPointCollection)

                    'Déplacer le sommet trouvé
                    pPointColl.UpdatePoint(nNumeroSommet, pPointB)

                    'Vérifier si seulement un sommet doit être traité
                    If bSeulementUn Then
                        'Vérifier si la partie de la ligne est fermé et que le numéro de sommet est le premier
                        If pPath.IsClosed And nNumeroSommet = 0 Then
                            'Mettre à jour le dernier sommet
                            pPointColl.UpdatePoint(pPointColl.PointCount - 1, pPointB)
                        End If

                        'Si tous les sommets doivent être traité
                    Else
                        'Interface pour calculer la distance
                        pProxOp = CType(pPointA, IProximityOperator)
                        'Redéfinir le numéro de partie au début
                        nNumeroPartie = 0
                        'Traiter tous les Path
                        Do Until (nNumeroPartie > pGeomColl.GeometryCount - 1)
                            'Interface pour traiter chaque Path
                            pPath = CType(pGeomColl.Geometry(nNumeroPartie), IPath)
                            'Interface pour déplacer le sommet
                            pPointColl = CType(pPath, IPointCollection)

                            'Traiter tous les sommets du Path
                            For j = 0 To pPointColl.PointCount - 1
                                'Vérifier la distance
                                If pProxOp.ReturnDistance(pPointColl.Point(j)) <= dTolerance Then
                                    'Déplacer le sommet trouvé
                                    pPointColl.UpdatePoint(j, pPointB)
                                End If
                            Next

                            'Vérifier la longueur du Path
                            If pPath.Length = 0 Then
                                'Détruire le path
                                pGeomColl.RemoveGeometries(nNumeroPartie, 1)
                            Else
                                'Changer de numéro de partie
                                nNumeroPartie = nNumeroPartie + 1
                            End If
                        Loop
                    End If

                    'Indiquer que le sommet a été détruit
                    DeplacerSommetGeometrie = True

                    'Sinon c'est l'absence d'un sommet
                ElseIf (pHitTest.HitTest(pPointA, dTolerance, esriGeometryHitPartType.esriGeometryPartBoundary, _
                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit)) Then
                    'Interface pour accéder aux parties de la géométrie
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Interface pour vérifier si la partie est fermée
                    pPath = CType(pGeomColl.Geometry(nNumeroPartie), IPath)

                    'Interface pour déplacer le sommet
                    pPointColl = CType(pPath, IPointCollection)

                    'Insérer un nouveau sommet
                    pPointColl.InsertPoints(nNumeroSommet + 1, 1, pPointB)

                    'Indiquer que le sommet a été détruit
                    DeplacerSommetGeometrie = True
                End If
            End If

            'Vérifier si un déplacement a eu lieu et que ce n'est pas un point
            If DeplacerSommetGeometrie And pGeometry.GeometryType <> esriGeometryType.esriGeometryPoint And bSimplifier Then
                'Simplifier la géométrie
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
            End If

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pClone = Nothing
            pPath = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pHitTest = Nothing
            pNewPoint = Nothing
            pProxOp = Nothing
            dDistance = Nothing
            nNumeroPartie = Nothing
            nNumeroSommet = Nothing
            bCoteDroit = Nothing
            j = Nothing
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
    Private Sub InitBarreProgression(ByVal iMin As Integer, ByVal iMax As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pStepPro As IStepProgressor = Nothing   'Interface qui permet de modifier les paramètres de la barre de progression.

        Try
            'sortir si le progressor est absent
            If pTrackCancel.Progressor Is Nothing Then Exit Sub

            'Interface pour modifier les paramètres de la barre de progression.
            pTrackCancel.Progressor = m_Application.StatusBar.ProgressBar
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

    ''' <summary> 
    ''' Cette fonction permet de retourner la définition de la géométrie à partir de la classe afin de vérifier la présence du Z et M.
    ''' </summary>
    ''' <param name="pFeatureClass"></param>
    ''' <returns>IGeometryDef</returns>
    Public Function RetournerGeometryDef(ByVal pFeatureClass As IFeatureClass) As IGeometryDef
        Dim shapeFieldName As String = pFeatureClass.ShapeFieldName
        Dim fields As IFields = pFeatureClass.Fields
        Dim geometryIndex As Integer = fields.FindField(shapeFieldName)
        Dim field As IField = fields.Field(geometryIndex)
        Dim geometryDef As IGeometryDef = field.GeometryDef
        Return geometryDef
    End Function
End Module
