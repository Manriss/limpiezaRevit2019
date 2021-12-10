'
' Created by SharpDevelop.
' User: GomezM2
' Date: 27/01/2015
' Time: 8:59
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Imports System
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI.Selection
Imports System.Collections.Generic
Imports System.Linq
Imports Microsoft .VisualBasic 


<Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)> _
<Autodesk.Revit.DB.Macros.AddInId("0C708221-AED8-41E5-80D4-E4BEAE9C4ECE")> _
Partial Public Class ThisApplication

    Private Sub Module_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
	
    End Sub
	
    Private Sub Module_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
	
    End Sub
	
	Public Sub LIMPIARFILTRO()
		'elimina todos los filtros que no estan en uso, o en una plantilla, tambien excluye los que contienen QTO
		Try
		Dim contador As Integer 
		Dim ch As Boolean 
		Dim vis As New FilteredElementCollector (ActiveUIDocument .Document )
		vis.OfCategory (BuiltInCategory .OST_Views)	'todas las vistas y plantillas de vista
		Dim fil2 As New FilteredElementCollector (ActiveUIDocument .Document )
		fil2.OfClass (GetType (filterelement))	'todos filtros de proyecto
		Dim report As String 'mensaje final de salida
		Dim idlist As New List (Of ElementId) 'colección de ID a borrar
		
		Dim filuso As  IList (Of String)=New List (Of String)	' nombre filtros en uso
		Dim filsinuso As  IList (Of String)=New List (Of String)	' nombre filtros SIN uso
		
		Dim filterid As ICollection(Of ElementId)
		Dim fil As FilterElement 
		For Each a As Autodesk .Revit.db.View In vis
			filterid=a.GetFilters 
			
			For Each b As ElementId In filterid
				fil=ActiveUIDocument.Document .GetElement (b)	'nombre del filtro en uso
				
						filuso.Add (fil.Name)	'todos los filtros en uso repetidos
					
				Next
				
				
			Next
			For Each h As FilterElement In fil2
				ch=true
				For Each f As String In filuso
					If h.Name =f  Then	'si el nombre del filtro no esta entre los utilizados
						
						ch=false
					Else
						
						
					End If
				Next
				If ch=True Then
					filsinuso .Add(h.Name)
				End If
			Next
			
			For Each d As String In filsinuso 
				For Each dd As FilterElement In fil2
					If d=dd.Name Then 'quiere decir que hay que borrarlo
						idlist .Add (dd.Id)	'lista con los ID a borrar
					End If
				Next
				contador =contador+1
			report= report & d &vbCr 'string de todos los filtros que se van a borrar
			
			Next
			Dim dialogo As New TaskDialog ("ATENCION")
			dialogo.MainInstruction ="ATENCION, SE VAN A BORRAR LOS SIGUIENTES FILTROS: (" & contador & "/" & idlist.Count & ")"
			dialogo.MainContent =report
			
			dialogo.CommonButtons =Autodesk .Revit .UI.TaskDialogCommonButtons .Cancel Or Autodesk .Revit .UI.TaskDialogCommonButtons .Ok
			dialogo.Title ="ATENCION"
			Dim resultado As Autodesk .Revit.UI.TaskDialogResult =dialogo.Show
			If resultado = vbOK Then 'vamos a borrar
				Using cam As New Transaction (ActiveUIDocument .Document ,"transa")
					If cam.Start =TransactionStatus .Started  Then
						ActiveUIDocument .Document .Delete (idlist )	'borrado
					End If
					Cam.Commit 
				End Using
			End If
		Catch ex As exception
			TaskDialog .Show ("ERROR",ex.ToString )
			End try
	End Sub
	
	Public Sub LIMPIAR()
		Try
			
		
		Dim mens As String
		Dim template As New List (Of ElementId )	'todas las templates de proyecto
		Dim templateuso As New List (Of Elementid)	'templates en uso
		Dim vistas As New FilteredElementCollector (ActiveUIDocument .Document )
		vistas.OfCategory (BuiltInCategory .OST_Views )
		For Each el As Autodesk .Revit .DB.View In vistas
			If el.IsTemplate Then	'todas las templates
				template .Add(el.Id)
			End If
			If el.ViewTemplateId.IntegerValue  <> -1 Then	'si la vista tiene plantilla asignada
				templateuso.Add (el.ViewTemplateId  )	'guardo el id de la plantilla, esta en uso
				
			End If
		Next
		Dim differenceQuery = template .Except(templateuso )
		
		For Each elem As ElementId In differenceQuery 
			mens=mens & ActiveUIDocument .Document .GetElement (elem).Name  & vbCr
		Next
		Dim dialogo As New TaskDialog ("ATENCION")
		dialogo .Title ="ATENCION"
		dialogo.MainIcon  = TaskDialogIcon .TaskDialogIconWarning 
		dialogo.TitleAutoPrefix =False
		dialogo.MainInstruction ="Las siguientes plantillas no estan asociadas a ninguna vista, se borraran." & vbCr & mens
		dialogo .CommonButtons =Autodesk .Revit.UI.TaskDialogCommonButtons .Ok Or Autodesk .Revit.UI.TaskDialogCommonButtons .Cancel 
		Dim resultado As Autodesk .Revit.UI.TaskDialogResult =dialogo .Show
		If resultado =vbOK Then	'ok a borrar
		Using tr As New Transaction (ActiveUIDocument .Document ,"ww")
			If tr.Start =TransactionStatus .Started  Then
				'borramos plantillas
					ActiveUIDocument .Document .Delete (differenceQuery.ToList  )
				
			End If
			tr.commit
		End Using
		TaskDialog .Show ("ATENCION","Se han borrado " & differenceQuery .Count & " plantillas de vista.")
		
		End If
		
		Catch ex As Exception 
			TaskDialog.Show ("ERROR", "Se ha producido un error desconocido")
		End Try
	End Sub
	End class