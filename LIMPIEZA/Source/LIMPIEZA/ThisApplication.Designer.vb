'
' Created by SharpDevelop.
' User: torneroloco
' Date: 10/12/2021
' Time: 10:23
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Option Strict Off
Option Explicit On

'''
<Global.System.CLSCompliantAttribute(false)>  _
Partial Public NotInheritable Class ThisApplication
    Inherits Autodesk.Revit.UI.Macros.ApplicationEntryPoint
    
    Public Event Startup As System.EventHandler
    
    Public Event Shutdown As System.EventHandler
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub OnStartup()
        Globals.ThisApplication = Me
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub FinishInitialization()
        MyBase.FinishInitialization
        Me.OnStartup
        If (Not (Me.StartupEvent) Is Nothing) Then
            RaiseEvent Startup(Me, System.EventArgs.Empty)
        End If
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub OnShutdown()
        If (Not (Me.ShutdownEvent) Is Nothing) Then
            RaiseEvent Shutdown(Me, System.EventArgs.Empty)
        End If
        MyBase.OnShutdown
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides ReadOnly Property PrimaryCookie() As String
        Get
            Return "ThisApplication"
        End Get
    End Property
End Class

'''
Partial Public NotInheritable Class Globals
    
    Private Shared _ThisApplication As ThisApplication
    
    Friend Shared Property ThisApplication() As ThisApplication
        Get
            Return _ThisApplication
        End Get
        Set
            If (_ThisApplication Is Nothing) Then
                _ThisApplication = value
            Else
                Throw New System.NotSupportedException
            End If
        End Set
    End Property
End Class
