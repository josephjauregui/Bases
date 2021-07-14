#tag Class
Protected Class ClaseErrorConexion
Implements iClaseErrorConexion
	#tag Method, Flags = &h0
		Function Error() As Boolean
		  Return mError
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Error(Assigns nuevoValor As Boolean)
		  mError = nuevoValor
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ErrorDetalles() As String
		  Return mErrorDetalles
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ErrorDetalles(Assigns nuevoValor As String)
		  mErrorDetalles = nuevoValor
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ErrorMensaje() As String
		  Return mErrorMensaje
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ErrorMensaje(Assigns nuevoValor As String)
		  mErrorMensaje = nuevoValor
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ErrorMuestraDialogo(titulo as string = "SOLUTICA© ")
		  Dim  d As New MessageDialog            
		  Dim  b As MessageDialogButton               
		  
		  d.Title = titulo + "Lamentablemente ocurrió un error..."
		  d.Icon = MessageDialog.GraphicStop       
		  d.ActionButton.Caption = "Aceptar"
		  
		  d.Message = "Error #" + mErrorNumero.ToText + " "  + mErrorMensaje
		  d.Explanation = mErrorDetalles
		  
		  b = d.ShowModal                 
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ErrorNumero() As Integer
		  Return mErrorNumero
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ErrorNumero(Assigns nuevoValor As Integer)
		  mErrorNumero = nuevoValor
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mError As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mErrorDetalles As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mErrorMensaje As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mErrorNumero As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
