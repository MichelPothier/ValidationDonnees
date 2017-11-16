Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.DataSourcesGDB
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

'**
'Nom de la composante : modExecuter.vb 
'
'''<summary>
''' Module principale permettant l'ex�cution de la validation des contraintes d'int�grit� pr�sentes dans une table de contraintes.
''' 
''' Les 11 param�tres de la ligne de commande du programme sont les suivants:
''' -------------------------------------------------------------------------
''' 1-Geodatabase des classes: Nom de la G�odatabase contenant les classes � traiter.
'''                            Obligatoire, D�faut : 
''' 2-Table des contraintes  : Nom de la table contenant l'information sur les contraintes d'int�grit� spatiales.
'''                            Obligatoire, D�faut : CONTRAINTE_INTEGRITE_SPATIALE
''' 3-Filtre des contraintes : Filtre effectu� sur la table des contraintes pour lesquelles on veut effectuer le traitement.
'''                            Optionnel, D�faut : # (Aucun filtre)
''' 4-Layer de d�coupage     : Nom du Layer de d�coupage dans lequel les �l�ments de d�coupage s�lectionn�s seront utilis�s pour effectuer le traitement.
'''                            Optionnel, D�faut : 
''' 5-Attribut de d�coupage  : Nom de l'attribut du Layer de d�coupage dans lequel les identifiants seront utilis�s pour effectuer le traitement.
'''                            Optionnel, D�faut : 
''' 6-R�pertoire des erreurs : Nom du r�pertoire dans lequel les erreurs seront �crites.
'''                            Optionnel, D�faut : 
''' 7-Rapport d'erreurs      : Nom du rapport dans lequel l'information sur les erreurs sera �crite.
'''                            Optionnel, D�faut : 
''' 8-Fichier journal        : Nom du fichier journal dans lequel l'information sur l'ex�cution du traitement sera �crite.
'''                            Optionnel, D�faut : 
''' 9-Courriel               : Addresses courriels des personnes � qui on veut envoyer le rapport d'erreurs.
'''                            Optionnel, D�faut : 
''' 10-Table des statistiques: Nom de la table contenant l'information sur les statistiques d'erreurs et de traitements.
'''                            Optionnel, D�faut : 
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 09 mai 2016
''' 
''' Pour augmenter la m�moire en 32-bits apr�s avoir compil� le programme:
''' "C:\Program Files (x86)\Microsoft Visual Studio 12.0\VC\bin\editbin.exe" /LARGEADDRESSAWARE S:\applications\gestion_bdg\pro\Geotraitement\exe\ValiderContrainte.exe
'''</remarks>
''' 
Module modExecuter
    '''<summary> Interface d'initialisation des licences ESRI.</summary>
    Private m_AOLicenseInitializer As LicenseInitializer = New ValiderContrainte.LicenseInitializer()
    '''<summary>Nom de la G�odatabase contenant les classes � valider.</summary>
    Private m_NomGeodatabaseClasses As String = "Database Connections\BDRS_PRO_BDG.sde"
    '''<summary>Nom de la table contenant l'information sur les contraintes d'int�grit�.</summary>
    Private m_NomTableContraintes As String = "CONTRAINTE_INTEGRITE_SPATIALE"
    '''<summary>Filter des contraintes � traiter.</summary>
    Private m_FiltreContraintes As String = ""
    '''<summary>Nom de la classe de d�coupage dans lequel les �l�ments de d�coupage s�lectionn�s seront utilis�s pour effectuer le traitement.</summary>
    Private m_NomClasseDecoupage As String = ""
    '''<summary>Nom de l'attribut du Layer de d�coupage dans lequel les identifiants seront utilis�s pour effectuer le traitement.</summary>
    Private m_NomAttributDecoupage As String = ""
    '''<summary>Nom du r�pertoire dans lequel les erreurs seront �crites.</summary>
    Private m_NomRepertoireErreurs As String = ""
    '''<summary>Nom du rapport d'erreurs dans lequel seront �crites les statistiques d'erreurs.</summary>
    Private m_NomRapportErreurs As String = ""
    '''<summary>Nom du fichier journal dans lequel l'ex�cution du traitement sera inscrit.</summary>
    Private m_NomFichierJournal As String = ""
    '''<summary>Addresses courriels des personnes � qui on veut envoyer le rapport d'ex�cution du traitement.</summary>
    Private m_Courriel As String = ""
    '''<summary>Nom de la table contenant l'information sur les statistiques d'erreurs et de traitements.</summary>
    Private m_NomTableStatistiques As String = ""
    '''<summary>Interface qui permet d'annuler l'ex�cution du traitement en inteactif.</summary>
    Private m_TrackCancel As ITrackCancel = Nothing

    <STAThread()> _
    Sub Main()
        'D�clarer les variables de travail
        Dim oTableContraintes As clsTableContraintes = Nothing  'Objet qui permet d'ex�cuter la validation des contraintes d'int�grit�.

        Try
            'ESRI License Initializer generated code.
            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeStandard}, _
                                                         New esriLicenseExtensionCode() {})

            'Valider les parametres de la ligne de commandes
            Call ValiderParametres()

            'Permettre d'annuler la s�lection avec la touche ESC
            m_TrackCancel = New CancelTracker
            m_TrackCancel.CancelOnKeyPress = False
            m_TrackCancel.CancelOnClick = False

            'Objet qui permet d'ex�cuter le traitement de validation des contraintes d'int�grit�
            oTableContraintes = New clsTableContraintes(m_NomGeodatabaseClasses, m_NomTableContraintes, m_FiltreContraintes)
            oTableContraintes.NomClasseDecoupage = m_NomClasseDecoupage
            oTableContraintes.NomAttributDecoupage = m_NomAttributDecoupage
            oTableContraintes.NomRepertoireErreurs = m_NomRepertoireErreurs
            oTableContraintes.NomRapportErreurs = m_NomRapportErreurs
            oTableContraintes.NomFichierJournal = m_NomFichierJournal
            oTableContraintes.Courriel = m_Courriel
            oTableContraintes.NomTableStatistiques = m_NomTableStatistiques
            oTableContraintes.TrackCancel = m_TrackCancel
            oTableContraintes.Commande = System.Environment.CommandLine

            'Ex�cuter le traitement de validation des contraintes d'int�grit�
            oTableContraintes.Executer()

            'Retourner le code d'ex�ution du traitement
            System.Environment.Exit(oTableContraintes.CodeErreur)

        Catch ex As Exception
            'Afficher l'erreur
            Console.WriteLine(ex.Message)
            '�crire le message d'erreur
            'File.AppendAllText("D:\Erreur_" & System.DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".log", ex.Message & vbCrLf)
            'Retourner le code d'�chec du traitement
            System.Environment.Exit(-1)
        Finally
            'Vider la m�moire
            oTableContraintes = Nothing
            'Fermer ad�quatement l'application utilisant les licences ESRI
            m_AOLicenseInitializer.ShutdownApplication()
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider et d�finir les param�tres du programme.
    '''</summary>
    '''
    Sub ValiderParametres()
        'D�claration des variables de travail
        Dim args() As String = System.Environment.GetCommandLineArgs()  'Contient les param�tres de la ligne de commandes

        Try
            'Valider le param�tre de la Geodatabase des classes � valider.
            Call ValiderParametreGeodatabaseClasses(args)

            'Valider le param�tre de la table contenant les contraintes d'int�grit�.
            Call ValiderParametreTableContraintes(args)

            'Valider le param�tre du filtre des contraintes.
            Call ValiderParametreFiltreContraintes(args)

            'Valider le param�tre contenant les identifiants de d�coupage � valider
            Call ValiderParametreClasseDecoupage(args)

            'Valider le param�tre de l'attribut contenant l'identifiant de d�coupage � valider
            Call ValiderParametreAttributDecoupage(args)

            'Valider le param�tre du r�pertoire dans lequel les erreurs seront �crites
            Call ValiderParametreRepertoireErreurs(args)

            'Valider le param�tre du rapport dans lequel les statistiques d'erreurs seront �crites
            Call ValiderParametreRapportErreurs(args)

            'Valider le param�tre du fichier journal dans lequel l'information sur l'ex�cution du traitement sera �crite.
            Call ValiderParametreFichierJournal(args)

            'Valider le param�tre des courriels qui recevront le rapport d'ex�cution du traitement
            Call ValiderParametreCourriels(args)

            'Valider le param�tre de la table contenant les statistiques.
            Call ValiderParametreTableStatistiques(args)

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la m�moire
            args = Nothing
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre du nom de la Geodatabase des classes � valider.
    '''</summary>
    '''
    Sub ValiderParametreGeodatabaseClasses(ByVal args() As String)
        Try
            'V�rifier si le param�tre de la Geodatabase des classes est pr�sent.
            If args.Length - 1 > 0 Then
                'D�finir le nom de la G�odatbase des classes
                m_NomGeodatabaseClasses = args(1)

            Else
                'Retourner l'erreur
                Err.Raise(-1, "ValiderParametreGeodatabaseClasses", "ERREUR : Le param�tre de la Geodatabase des classes est absent !")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre de la table des contraintes d'int�grit�.
    '''</summary>
    '''
    Sub ValiderParametreTableContraintes(ByVal args() As String)
        Try
            'V�rifier si le param�tre de la table des contraintes d'int�grit� est pr�sent.
            If args.Length - 1 > 1 Then
                'D�finir le nom de la table des contraintes
                m_NomTableContraintes = args(2)

            Else
                'Retourner l'erreur
                Err.Raise(-1, "ValiderParametreTableContraintes", "ERREUR : Le param�tre de la table des contraintes est absent !")
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre du filtre des contraintes � valider.
    '''</summary>
    '''
    Sub ValiderParametreFiltreContraintes(ByVal args() As String)
        Try
            'V�rifier si le param�tre de la liste des groupes est pr�sent.
            If args.Length - 1 > 2 Then
                'V�rifier si le param�tre contient un filtre des contraintes
                If args(3) <> "#" Then
                    'D�finir le filtre des contraintes
                    m_FiltreContraintes = args(3)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre du nom du Layer contenant les identifiants de d�coupage � valider.
    '''</summary>
    '''
    Sub ValiderParametreClasseDecoupage(ByVal args() As String)
        Try
            'Valider le param�tre du nom du Layer contenant les identifiants de d�coupage � valider
            If args.Length - 1 > 3 Then
                'V�rifier si le param�tre contient le nom de la classe de d�coupage
                If args(4) <> "#" Then
                    'D�finir le nom de la classe de d�coupage 
                    m_NomClasseDecoupage = args(4)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre du nom de l'attribut contenant l'identifiant de d�coupage � valider.
    '''</summary>
    '''
    Sub ValiderParametreAttributDecoupage(ByVal args() As String)
        Try
            'Valider le param�tre du nom de l'attribut contenant l'identifiant de d�coupage � valider
            If args.Length - 1 > 4 Then
                'V�rifier si le param�tre contient le nom de l'attribut de d�coupage
                If args(5) <> "#" Then
                    'D�finir le nom de l'attribut de d�coupage
                    m_NomAttributDecoupage = args(5)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre du r�pertoire dans lequel les erreurs seront �crites.
    '''</summary>
    '''
    Sub ValiderParametreRepertoireErreurs(ByVal args() As String)
        Try
            'Valider le param�tre du r�pertoire ou de la g�odatabase dans lequel les erreurs seront �crites
            If args.Length - 1 > 5 Then
                'V�rifier si le param�tre contient le nom du r�pertoire ou de la g�odatabase d'erreurs
                If args(6) <> "#" Then
                    'D�finir le nom du r�pertoire ou de la g�odatabase d'erreurs
                    m_NomRepertoireErreurs = args(6)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre du rapport dans lequel les statistiques d'erreurs seront �crites.
    '''</summary>
    '''
    Sub ValiderParametreRapportErreurs(ByVal args() As String)
        Try
            'Valider le param�tre du rapport d'erreurs
            If args.Length - 1 > 6 Then
                'V�rifier si le param�tre contient le nom du rapport d'erreurs
                If args(7) <> "#" Then
                    'D�finir le nom du rapport d'erreurs
                    m_NomRapportErreurs = args(7)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre du fichier journal dans lequel l'information sur l'ex�cution du traitement sera �crite.
    '''</summary>
    '''
    Sub ValiderParametreFichierJournal(ByVal args() As String)
        Try
            'Valider le param�tre du r�pertoire dans lequel les erreurs seront �crites
            If args.Length - 1 > 7 Then
                'V�rifier si le param�tre contient le nom du fichier journal
                If args(8) <> "#" Then
                    'D�finir le nom du fichier journal
                    m_NomFichierJournal = args(8)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre des courriels qui recevront le rapport d'ex�cution du traitement.
    '''</summary>
    '''
    Sub ValiderParametreCourriels(ByVal args() As String)
        Try
            'Valider le param�tre des courriels qui recevront le rapport d'ex�cution du traitement
            If args.Length - 1 > 8 Then
                'V�rifier si le param�tre contient les courriels pour lesquels on veut envoyer le rapport d'erreurs
                If args(9) <> "#" Then
                    'D�finir les courriels pour lesquels on veut envoyer le rapport d'erreurs
                    m_Courriel = args(9)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub

    '''<summary>
    '''Routine qui permet de valider le param�tre de la table des statistiques.
    '''</summary>
    '''
    Sub ValiderParametreTableStatistiques(ByVal args() As String)
        Try
            'Valider le param�tre de la table des statistiques
            If args.Length - 1 > 9 Then
                'V�rifier si le param�tre contient le nom de la table des statistiques
                If args(10) <> "#" Then
                    'D�finir le nom de la table des statistiques
                    m_NomTableStatistiques = args(10)
                End If
            End If

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        End Try
    End Sub
End Module
