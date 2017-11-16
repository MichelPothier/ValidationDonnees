'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Desktop.AddIns
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports System
Imports System.Collections.Generic

Namespace My
    
    '''<summary>
    '''A class for looking up declarative information in the associated configuration xml file (.esriaddinx).
    '''</summary>
    Friend Module ThisAddIn
        
        Friend ReadOnly Property Name() As String
            Get
                Return "BarreLayerSelection"
            End Get
        End Property
        
        Friend ReadOnly Property AddInID() As String
            Get
                Return "{30265e03-9a23-403a-aed1-3c694f1f54f9}"
            End Get
        End Property
        
        Friend ReadOnly Property Company() As String
            Get
                Return "MPO"
            End Get
        End Property
        
        Friend ReadOnly Property Version() As String
            Get
                Return "1.0"
            End Get
        End Property
        
        Friend ReadOnly Property Description() As String
            Get
                Return "Permet d’effectuer différents traitements (Créer, Retirer, Détruire, Union, Inter"& _ 
                    "section, Différence) relatifs à la conservation des éléments selon la sélection "& _ 
                    "d’éléments et/ou la liste d’identifiants d’éléments contenus dans les « FeatureL"& _ 
                    "ayers »."
            End Get
        End Property
        
        Friend ReadOnly Property Author() As String
            Get
                Return "Michel Pothier"
            End Get
        End Property
        
        Friend ReadOnly Property [Date]() As String
            Get
                Return "27/08/2013"
            End Get
        End Property
        
        <System.Runtime.CompilerServices.ExtensionAttribute()>  _
        Friend Function ToUID(ByVal id As String) As ESRI.ArcGIS.esriSystem.UID
            Dim uid As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UIDClass()
            uid.Value = id
            Return uid
        End Function
        
        '''<summary>
        '''A class for looking up Add-in id strings declared in the associated configuration xml file (.esriaddinx).
        '''</summary>
        Friend Class IDs
            
            '''<summary>
            '''Returns 'MPO_BarreLayerSelection_cmdCreerListeIdentifiant', the id declared for Add-in Button class 'cmdCreerListeIdentifiant'
            '''</summary>
            Friend Shared ReadOnly Property cmdCreerListeIdentifiant() As String
                Get
                    Return "MPO_BarreLayerSelection_cmdCreerListeIdentifiant"
                End Get
            End Property
            
            '''<summary>
            '''Returns 'MPO_BarreLayerSelection_cmdRetirerListeIdentifiant', the id declared for Add-in Button class 'cmdRetirerListeIdentifiant'
            '''</summary>
            Friend Shared ReadOnly Property cmdRetirerListeIdentifiant() As String
                Get
                    Return "MPO_BarreLayerSelection_cmdRetirerListeIdentifiant"
                End Get
            End Property
            
            '''<summary>
            '''Returns 'MPO_BarreLayerSelection_cmdDetruireListeIdentifiant', the id declared for Add-in Button class 'cmdDetruireListeIdentifiant'
            '''</summary>
            Friend Shared ReadOnly Property cmdDetruireListeIdentifiant() As String
                Get
                    Return "MPO_BarreLayerSelection_cmdDetruireListeIdentifiant"
                End Get
            End Property
            
            '''<summary>
            '''Returns 'MPO_BarreLayerSelection_cmdUnionSelection', the id declared for Add-in Button class 'cmdUnionSelection'
            '''</summary>
            Friend Shared ReadOnly Property cmdUnionSelection() As String
                Get
                    Return "MPO_BarreLayerSelection_cmdUnionSelection"
                End Get
            End Property
            
            '''<summary>
            '''Returns 'MPO_BarreLayerSelection_cmdIntersectSelection', the id declared for Add-in Button class 'cmdIntersectSelection'
            '''</summary>
            Friend Shared ReadOnly Property cmdIntersectSelection() As String
                Get
                    Return "MPO_BarreLayerSelection_cmdIntersectSelection"
                End Get
            End Property
            
            '''<summary>
            '''Returns 'MPO_BarreLayerSelection_cmdDifferenceSelection', the id declared for Add-in Button class 'cmdDifferenceSelection'
            '''</summary>
            Friend Shared ReadOnly Property cmdDifferenceSelection() As String
                Get
                    Return "MPO_BarreLayerSelection_cmdDifferenceSelection"
                End Get
            End Property
        End Class
    End Module
    
Friend Module ArcMap
  Private s_app As ESRI.ArcGIS.Framework.IApplication
  Private s_docEvent As ESRI.ArcGIS.ArcMapUI.IDocumentEvents_Event

  Public ReadOnly Property Application() As ESRI.ArcGIS.Framework.IApplication
    Get
      If s_app Is Nothing Then
        s_app = TryCast(Internal.AddInStartupObject.GetHook(Of ESRI.ArcGIS.ArcMapUI.IMxApplication)(), ESRI.ArcGIS.Framework.IApplication)
                If s_app Is Nothing Then
                    Dim editorHost As ESRI.ArcGIS.Editor.IEditor = Internal.AddInStartupObject.GetHook(Of ESRI.ArcGIS.Editor.IEditor)()
                    If editorHost IsNot Nothing Then s_app = editorHost.Parent
                End If
            End If

            Return s_app
        End Get
    End Property

    Public ReadOnly Property Document() As ESRI.ArcGIS.ArcMapUI.IMxDocument
        Get
            If Application IsNot Nothing Then
                Return TryCast(Application.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
            End If

            Return Nothing
        End Get
    End Property
    Public ReadOnly Property ThisApplication() As ESRI.ArcGIS.ArcMapUI.IMxApplication
        Get
            Return TryCast(Application, ESRI.ArcGIS.ArcMapUI.IMxApplication)
        End Get
    End Property
    Public ReadOnly Property DockableWindowManager() As ESRI.ArcGIS.Framework.IDockableWindowManager
        Get
            Return TryCast(Application, ESRI.ArcGIS.Framework.IDockableWindowManager)
        End Get
    End Property

    Public ReadOnly Property Events() As ESRI.ArcGIS.ArcMapUI.IDocumentEvents_Event
        Get
            s_docEvent = TryCast(Document, ESRI.ArcGIS.ArcMapUI.IDocumentEvents_Event)
            Return s_docEvent
        End Get
    End Property

    Public ReadOnly Property Editor() As ESRI.ArcGIS.Editor.IEditor
        Get
            Dim editorUID As New ESRI.ArcGIS.esriSystem.UID
            editorUID.Value = "esriEditor.Editor"
            Return TryCast(Application.FindExtensionByCLSID(editorUID), ESRI.ArcGIS.Editor.IEditor)
        End Get
    End Property
End Module

Namespace Internal
  <ESRI.ArcGIS.Desktop.AddIns.StartupObjectAttribute(), _
   Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), _
   Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()> _
  Partial Public Class AddInStartupObject
    Inherits ESRI.ArcGIS.Desktop.AddIns.AddInEntryPoint

    Private m_addinHooks As List(Of Object)
    Private Shared _sAddInHostManager As AddInStartupObject

    Public Sub New()

    End Sub

    <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Protected Overrides Function Initialize(ByVal hook As Object) As Boolean
      Dim createSingleton As Boolean = _sAddInHostManager Is Nothing
      If createSingleton Then
        _sAddInHostManager = Me
        m_addinHooks = New List(Of Object)
        m_addinHooks.Add(hook)
      ElseIf Not _sAddInHostManager.m_addinHooks.Contains(hook) Then
        _sAddInHostManager.m_addinHooks.Add(hook)
      End If

      Return createSingleton
    End Function

    <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Protected Overrides Sub Shutdown()
      _sAddInHostManager = Nothing
      m_addinHooks = Nothing
    End Sub

    <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Friend Shared Function GetHook(Of T As Class)() As T
      If _sAddInHostManager IsNot Nothing Then
        For Each o As Object In _sAddInHostManager.m_addinHooks
          If TypeOf o Is T Then
            Return DirectCast(o, T)
          End If
        Next
      End If

      Return Nothing
    End Function

    ''' <summary>
    ''' Expose this instance of Add-in class externally
    ''' </summary>
    Public Shared Function GetThis() As AddInStartupObject
      Return _sAddInHostManager
    End Function

  End Class
End Namespace

End Namespace
