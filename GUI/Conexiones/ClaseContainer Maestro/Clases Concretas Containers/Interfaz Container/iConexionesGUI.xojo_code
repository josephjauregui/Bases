#tag Interface
Protected Interface iConexionesGUI
	#tag Method, Flags = &h0
		Sub Close()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EmbedWithin(containingControl As RectControl, left As Integer = 0, top As Integer = 0, width As Integer = - 1, height As Integer = - 1)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EmbedWithin(containingWindow As Window, left As Integer = 0, top As Integer = 0, width As Integer = - 1, height As Integer = - 1)
		  
		End Sub
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
