'**
'Nom de la composante : citsCls_Pixel2Polygon.vb
'
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.DataSourcesRaster

'''<summary>
''' Classe qui permet de transformer chaque pixel d'une image en polygon.
'''</summary>
'''
'''<remarks>
'''Auteur : Michel Pothier
'''Date : 27 avril 2011
'''</remarks>
''' 
Public Class clsPixel2Polygon
    'Déclarer les variables globales
    '''<summary>Interface contenant les paramètres d'affichage de l'image matricielle.</summary>
    Private gpRasterLayer As IRasterLayer = Nothing

    '''<summary>Interface ESRI contenant la couleur RGB de texte par défaut.</summary>
    Private gpRgbColorText As IRgbColor = Nothing
    '''<summary>Interface ESRI contenant la couleur RGB de point par défaut.</summary>
    Private gpRgbColorPoint As IRgbColor = Nothing
    '''<summary>Interface ESRI contenant la couleur RGB de ligne par défaut.</summary>
    Private gpRgbColorPolyline As IRgbColor = Nothing
    '''<summary>Interface ESRI contenant la couleur RGB de surface par défaut.</summary>
    Private gpRgbColorPolygon As IRgbColor = Nothing

    '''<summary>Interface ESRI contenant le symbole de texte par défaut.</summary>
    Private gpSymbolText As ITextSymbol = Nothing
    '''<summary>Interface ESRI contenant le symbole de point par défaut.</summary>
    Private gpSymbolPoint As ISimpleMarkerSymbol = Nothing
    '''<summary>Interface ESRI contenant le symbole de ligne par défaut.</summary>
    Private gpSymbolPolyline As ISimpleLineSymbol = Nothing
    '''<summary>Interface ESRI contenant le symbole de surface par défaut.</summary>
    Private gpSymbolPolygon As ISimpleFillSymbol = Nothing

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet de construire et initialiser la classe.
    '''</summary>
    '''
    '''<param name="pRasterLayer">Interface contenant les paramètres d'affichage de l'image matricielle.</param>
    ''' 
    Public Sub New(ByVal pRasterLayer As IRasterLayer)
        'Définir l'image et le nombre de pixel à calculer pour une ligne (X) et une colonne (Y).
        gpRasterLayer = pRasterLayer

        'Initialiser les couleurs et les symboles
        Call Init()
    End Sub

    '''<summary>
    '''Permet de vider la mémoire.
    '''</summary>
    '''
    Protected Overrides Sub finalize()
        'Libérer les ressources
        gpRasterLayer = Nothing
        gpRgbColorText = Nothing
        gpRgbColorPoint = Nothing
        gpRgbColorPolyline = Nothing
        gpRgbColorPolygon = Nothing
        gpSymbolText = Nothing
        gpSymbolPoint = Nothing
        gpSymbolPolyline = Nothing
        gpSymbolPolygon = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    '''Propriété qui permet de définir et retourner l'interface contenant les paramètres d'affichage de l'image matricielle.
    '''</summary>
    ''' 
    Public Property RasterLayer() As IRasterLayer
        Get
            RasterLayer = gpRasterLayer
        End Get
        Set(ByVal value As IRasterLayer)
            gpRasterLayer = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner la couleur RGB de texte par défaut.
    '''</summary>
    ''' 
    Public Property RgbColorText() As IRgbColor
        Get
            RgbColorText = gpRgbColorText
        End Get
        Set(ByVal value As IRgbColor)
            gpRgbColorText = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner la couleur RGB de poinr par défaut.
    '''</summary>
    ''' 
    Public Property RgbColorPoint() As IRgbColor
        Get
            RgbColorPoint = gpRgbColorPoint
        End Get
        Set(ByVal value As IRgbColor)
            gpRgbColorPoint = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner la couleur RGB de ligne par défaut.
    '''</summary>
    ''' 
    Public Property RgbColorPolyline() As IRgbColor
        Get
            RgbColorPolyline = gpRgbColorPolyline
        End Get
        Set(ByVal value As IRgbColor)
            gpRgbColorPolyline = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner la couleur RGB de surface par défaut.
    '''</summary>
    ''' 
    Public Property RgbColorPolygon() As IRgbColor
        Get
            RgbColorPolygon = gpRgbColorPolygon
        End Get
        Set(ByVal value As IRgbColor)
            gpRgbColorPolygon = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner le symbole de texte par défaut.
    '''</summary>
    ''' 
    Public Property SymbolText() As ITextSymbol
        Get
            SymbolText = gpSymbolText
        End Get
        Set(ByVal value As ITextSymbol)
            gpSymbolText = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner le symbole de point par défaut.
    '''</summary>
    ''' 
    Public Property SymbolPoint() As ISimpleMarkerSymbol
        Get
            SymbolPoint = gpSymbolPoint
        End Get
        Set(ByVal value As ISimpleMarkerSymbol)
            gpSymbolPoint = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner le symbole de ligne par défaut.
    '''</summary>
    ''' 
    Public Property SymbolPolyline() As ISimpleLineSymbol
        Get
            SymbolPolyline = gpSymbolPolyline
        End Get
        Set(ByVal value As ISimpleLineSymbol)
            gpSymbolPolyline = value
        End Set
    End Property

    '''<summary>
    '''Propriété qui permet de définir et retourner le symbole de surface par défaut.
    '''</summary>
    ''' 
    Public Property SymbolPolygon() As ISimpleFillSymbol
        Get
            SymbolPolygon = gpSymbolPolygon
        End Get
        Set(ByVal value As ISimpleFillSymbol)
            gpSymbolPolygon = value
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"

    '''<summary>
    ''' Routine qui permet d'afficher les valeurs des pixels et leurs formes selon un nombre maximum de pixels à afficher en X et Y.
    '''</summary>
    '''
    '''<param name="pActiveView">Interface contenant la vue d'affichage active.</param>
    '''<param name="pStatusBar">Interface contenant la barre de statut.</param>
    '''<param name="nNbPixelX">Contient le nombre de pixel à calculer pour une ligne (X).</param>
    '''<param name="nNbPixelY">Contient le nombre de pixel à calculer pour une colonne (Y).</param>
    ''' 
    Public Sub AfficherValeurPixel(ByVal pActiveView As IActiveView, ByVal pStatusBar As IStatusBar, _
    Optional ByVal nNbPixelX As Integer = 36, Optional ByVal nNbPixelY As Integer = 36)
        'Déclarer les variables de travail
        Dim pSpatialReference As ISpatialReference = Nothing    'Interface contenant la référence spatiale
        Dim pRaster As IRaster2 = Nothing           'Interface utilisé pour extraire les coordonnées image
        Dim pRasterProps As IRasterProps = Nothing  'Interface contenant le spropriété de l'image
        Dim pPolygon As IPolygon = Nothing          'Interface contenant le polygon représentant la forme du pixel
        Dim pPoint As IPoint = Nothing              'Interface contenant le centre du pixel
        Dim pEnvelope As IEnvelope = Nothing        'Interface contenant l'enveloppe de la fenêtre active.
        Dim pUpperLeft As IPoint = Nothing          'Interface contenant le point du coin supérieur gauche de la vue active
        Dim pLowerRight As IPoint = Nothing         'Interface contenant le point du coin inférieur droit de la vue active
        Dim pClone As IClone = Nothing              'Interface qui permet de clone le point
        Dim vValue As Object = Nothing              'Contient la valeur du pixel traité
        Dim x As Double, y As Double                'Contient la coordonnée X et Y correcpondant au centre du pixel
        Dim dSizeX As Double, dSizeY As Double      'Contient la demi-taille du pixel
        Dim nStepX As Integer, nStepY As Integer    'Contient le pas en X et Y d'affichage des pixels
        Dim i As Integer, j As Integer              'Compteur X et Y
        Dim a As Integer, b As Integer              'Contient la limite minimum et maximum en X (ligne) des coordonnées image
        Dim c As Integer, d As Integer              'Contient la limite minimum et maximum en Y (colonne) des coordonnées image

        Try
            'Définir la référence spatial d'affichage
            pSpatialReference = pActiveView.FocusMap.SpatialReference

            'Interface contenant l'image matricielle
            pRaster = CType(gpRasterLayer.Raster, IRaster2)

            'Interface pour extraire les propriétés de l'image
            pRasterProps = CType(pRaster, IRasterProps)

            'Définir la demi-taille du pixel
            dSizeX = ((pRasterProps.Extent.XMax - pRasterProps.Extent.XMin) / pRasterProps.Width) / 2
            dSizeY = ((pRasterProps.Extent.YMax - pRasterProps.Extent.YMin) / pRasterProps.Height) / 2

            'Extraire l'enveloppe de la fenêtre visible
            pEnvelope = pActiveView.Extent
            pEnvelope.Project(pRasterProps.SpatialReference)

            'Définir le point correspondant au coin supérieur gauche
            pUpperLeft = pEnvelope.UpperLeft
            pUpperLeft.Project(pRasterProps.SpatialReference)

            'Définir le point correspondant au coin inférieur droit
            pLowerRight = pEnvelope.LowerRight
            pLowerRight.Project(pRasterProps.SpatialReference)

            'Définir le point utilisé pour contruire les points du polygone et le centre du pixel
            pPoint = pActiveView.Extent.UpperLeft
            pPoint.Project(pRasterProps.SpatialReference)
            pClone = CType(pPoint, IClone)

            'Retourner la position en pixel du coin supérieur gauche
            pRaster.MapToPixel(pUpperLeft.X, pUpperLeft.Y, a, c)
            'Retourner la position en pixel du coin inférieur droit
            pRaster.MapToPixel(pLowerRight.X, pLowerRight.Y, b, d)

            'Définir le pas utilisé en X et Y en fonction du nombre de pixel permis afin de ne pas surcharger l'affichage
            nStepX = CInt(((b - a) / nNbPixelX) + 0.5)
            If nStepX = 0 Then nStepX = 1
            nStepY = CInt(((d - c) / nNbPixelY) + 0.5)
            If nStepY = 0 Then nStepY = 1

            'Afficher le pas utilisé en X et Y dans la barre de statut
            pStatusBar.Message(0) = gpRasterLayer.Name & ", X:" & Str(nNbPixelX) & " (1/" & Str(nStepX) & ")," _
            & " Y:" & Str(nNbPixelY) & " (1/" & Str(nStepY) & ")"

            'Définir le début et fin des coordonnées images X et Y à traiter
            If a < 0 Then a = 0
            If c < 0 Then c = 0
            If b < a Then b = a + 1
            If d < c Then d = c + 1
            If b > pRasterProps.Width - 1 Then b = pRasterProps.Width - 1
            If d > pRasterProps.Height - 1 Then d = pRasterProps.Height - 1

            'Traiter toutes les coordonnées images en X
            For i = a To b Step nStepX
                'Traiter toutes les coordonnées images en Y
                For j = c To d Step nStepY
                    'Extraire la valeur du pixel
                    vValue = pRaster.GetPixelValue(0, i, j)

                    'Vérifier si la valeur est nulle
                    If vValue IsNot Nothing Then
                        'Extraire la coordonnée X et Y de la carte à partir de la ligne (i) et la colonne (j)
                        pRaster.PixelToMap(i, j, x, y)

                        'Créer le point du centre du pixel
                        pPoint = CType(pClone.Clone, IPoint)
                        pPoint.X = x
                        pPoint.Y = y

                        'Créer le polygon représentant la forme du pixel
                        pPolygon = CreerPolygonPixel(pSpatialReference, pPoint, dSizeX, dSizeY)

                        'Projeter le centre du pixel selon la référence spatiale de la fenêtre d'affichage
                        pPoint.Project(pSpatialReference)

                        'Afficher la valeur, le centre et le polygone du pixel dans la fenêtre d'affichage
                        Call DessinerPixel(pActiveView.ScreenDisplay, pPolygon, pPoint, CInt(vValue).ToString)
                    End If
                Next j
            Next i

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pSpatialReference = Nothing
            pRaster = Nothing
            pRasterProps = Nothing
            pPolygon = Nothing
            pPoint = Nothing
            pEnvelope = Nothing
            pUpperLeft = Nothing
            pLowerRight = Nothing
            pClone = Nothing
            vValue = Nothing
        End Try
    End Sub
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet d'initialiser les couleurs et les symboles d'affichage.
    '''</summary>
    ''' 
    Private Sub Init()
        'Déclarer les variables de travail
        Try
            'Créer la couleur pour le texte
            gpRgbColorText = New RgbColor
            gpRgbColorText.Red = 255

            'Créer la couleur pour le point
            gpRgbColorPoint = New RgbColor
            gpRgbColorPoint.Red = 200
            gpRgbColorPoint.Blue = 200
            gpRgbColorPoint.Green = 200

            'Créer la couleur pour la ligne
            gpRgbColorPolyline = New RgbColor
            gpRgbColorPolyline.Red = 200
            gpRgbColorPolyline.Blue = 200
            gpRgbColorPolyline.Green = 200

            'Créer la couleur pour la surface
            gpRgbColorPolygon = New RgbColor
            gpRgbColorPolygon.Red = 200
            gpRgbColorPolygon.Blue = 200
            gpRgbColorPolygon.Green = 200

            'Créer le symbole pour le texte
            gpSymbolText = New TextSymbol
            gpSymbolText.Color = gpRgbColorText
            gpSymbolText.Font.Bold = True
            gpSymbolText.HorizontalAlignment = esriTextHorizontalAlignment.esriTHACenter
            gpSymbolText.Size = 10

            'Créer le symbole pour le point
            gpSymbolPoint = New SimpleMarkerSymbol
            gpSymbolPoint.Color = gpRgbColorPoint
            gpSymbolPoint.Style = esriSimpleMarkerStyle.esriSMSCircle
            gpSymbolPoint.Size = 1

            'Définir la symbologie pour la ligne
            gpSymbolPolyline = New SimpleLineSymbol
            gpSymbolPolyline.Color = gpRgbColorPolyline

            'Créer le symbole pour la surface
            gpSymbolPolygon = New SimpleFillSymbol
            gpSymbolPolygon.Color = gpRgbColorPolygon
            gpSymbolPolygon.Outline = gpSymbolPolyline
            gpSymbolPolygon.Style = esriSimpleFillStyle.esriSFSNull

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de créer un polygon représentant la forme d'un pixel.
    '''</summary>
    '''
    '''<param name="pSpatialReference">Interface contenant la référence spatiale du polygon à créer.</param>
    '''<param name="pCentrePixel">Interface contenant le Point du centre du pixel.</param>
    '''<param name="dSizeX">Contient la moitié de la taille du pixel en X.</param>
    '''<param name="dSizeY">Contient la moitié de la taille du pixel en Y.</param>
    '''
    '''<returns>"IPolygon" représentant le pixel, "Nothing" sinon.</returns>
    ''' 
    Private Function CreerPolygonPixel(ByVal pSpatialReference As ISpatialReference, ByVal pCentrePixel As IPoint, _
    ByVal dSizeX As Double, ByVal dSizeY As Double) As IPolygon
        'Déclarer les variables de travail
        Dim pPolygon As IPolygon = Nothing              'Interface contenant le polygon représentant la forme du pixel
        Dim pPointColl As IPointCollection = Nothing    'Interface qui permet de construire le polygone
        Dim pPoint As IPoint = Nothing                  'Interface contenant un point du polygone
        Dim pClone As IClone = Nothing                  'Interface qui permet de clone le point

        'Définir la valeur de retour par défaut
        CreerPolygonPixel = Nothing

        Try
            'Interface qui permet de clone le centre du pixel
            pClone = CType(pCentrePixel, IClone)

            'Créer un nouveau polygon vide
            pPolygon = New Polygon
            pPolygon.SpatialReference = pSpatialReference

            'Interface pour ajouter les sommets au polygon
            pPointColl = CType(pPolygon, IPointCollection)

            'Ajouter le premier sommet
            pPoint = CType(pClone.Clone, IPoint)
            pPoint.X = pCentrePixel.X - dSizeX
            pPoint.Y = pCentrePixel.Y + dSizeY
            pPoint.Project(pSpatialReference)
            pPointColl.AddPoint(pPoint)

            'Ajouter le deuxième sommet
            pPoint = CType(pClone.Clone, IPoint)
            pPoint.X = pCentrePixel.X + dSizeX
            pPoint.Y = pCentrePixel.Y + dSizeY
            pPoint.Project(pSpatialReference)
            pPointColl.AddPoint(pPoint)

            'Ajouter le troisième sommet
            pPoint = CType(pClone.Clone, IPoint)
            pPoint.X = pCentrePixel.X + dSizeX
            pPoint.Y = pCentrePixel.Y - dSizeY
            pPoint.Project(pSpatialReference)

            'Ajouter le quatrième sommet
            pPointColl.AddPoint(pPoint)
            pPoint = CType(pClone.Clone, IPoint)
            pPoint.X = pCentrePixel.X - dSizeX
            pPoint.Y = pCentrePixel.Y - dSizeY
            pPoint.Project(pSpatialReference)
            pPointColl.AddPoint(pPoint)

            'Fermer le polygon
            pPolygon.Close()

            'Retourner le polygone
            CreerPolygonPixel = pPolygon

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pPolygon = Nothing
            pPointColl = Nothing
            pPoint = Nothing
            pClone = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de dessiner un pixel en utilisant un Polygon, un point et un texte contenant la valeur du pixel.
    '''</summary>
    '''
    '''<param name="pScreenDisplay">Interface contenant la fenêtre d'affichage.</param>
    '''<param name="pPolygon">Interface contenant le Polygon représentant le pixel.</param>
    '''<param name="pPoint">Interface contenant le Point du centre du pixel.</param>
    '''<param name="sTexte">Contient le texte à afficher et représentant la valeur du pixel.</param>
    ''' 
    Private Sub DessinerPixel(ByVal pScreenDisplay As IScreenDisplay, ByVal pPolygon As IPolygon, ByVal pPoint As IPoint, ByVal sTexte As String)
        'Déclarer les variables de travail
        Try

            'Afficher le polygone, le point du centre et le texte dans la vue active
            With pScreenDisplay
                'Débuter l'affichage
                .StartDrawing(pScreenDisplay.hDC, -1)

                'Afficher la géométrie avec sa symbologie dans la vue active
                .SetSymbol(CType(gpSymbolPolygon, ISymbol))
                .DrawPolygon(pPolygon)

                'Afficher la géométrie avec sa symbologie dans la vue active
                .SetSymbol(CType(gpSymbolPoint, ISymbol))
                .DrawPoint(pPoint)

                'Afficher le texte avec sa symbologie dans la vue active
                .SetSymbol(CType(gpSymbolText, ISymbol))
                .DrawText(pPoint, sTexte)

                'Terminer l'affichage
                .FinishDrawing()
            End With

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
        End Try
    End Sub
#End Region
End Class

