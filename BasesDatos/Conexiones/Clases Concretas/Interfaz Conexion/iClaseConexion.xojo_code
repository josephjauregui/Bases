#tag Interface
Protected Interface iClaseConexion
	#tag Method, Flags = &h0
		Sub Constructor(parametros As string)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Desconectar()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Error() As Boolean
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ErrorDetalles() As String
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ErrorMensaje() As String
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ErrorMuestraDialogo(titulo as string = "SOLUTICA© ")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ErrorNumero() As Integer
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerBasesDatos() As Dictionary
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerCampos(BaseDatos As String, Tabla As String) As Dictionary
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerTablas(BaseDatos As String) As Dictionary
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Tipo() As TiposConexionEnum
		  
		End Function
	#tag EndMethod


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
End Interface
#tag EndInterface
