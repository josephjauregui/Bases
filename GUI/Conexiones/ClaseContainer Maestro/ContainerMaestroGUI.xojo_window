#tag Window
Begin ContainerControl ContainerMaestroGUI
   AcceptFocus     =   False
   AcceptTabs      =   True
   AutoDeactivate  =   True
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   Compatibility   =   ""
   DoubleBuffer    =   False
   Enabled         =   True
   EraseBackground =   True
   HasBackColor    =   False
   Height          =   430
   HelpTag         =   ""
   InitialParent   =   ""
   Left            =   0
   LockBottom      =   False
   LockLeft        =   False
   LockRight       =   False
   LockTop         =   False
   TabIndex        =   0
   TabPanelIndex   =   0
   TabStop         =   True
   Top             =   0
   Transparent     =   True
   UseFocusRing    =   False
   Visible         =   True
   Width           =   330
   Begin ClaseTipoConexionGUI TipoConexionGUI
      AcceptFocus     =   False
      AcceptTabs      =   True
      AutoDeactivate  =   True
      BackColor       =   &cFFFFFF00
      Backdrop        =   0
      DoubleBuffer    =   False
      Enabled         =   True
      EraseBackground =   True
      HasBackColor    =   False
      Height          =   110
      HelpTag         =   ""
      InitialParent   =   ""
      Left            =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   2
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   0
      Transparent     =   True
      UseFocusRing    =   False
      Visible         =   True
      Width           =   330
   End
End
#tag EndWindow

#tag WindowCode
	#tag Method, Flags = &h21
		Private Sub ConectadoAction(sender As iConexionesGUI, conexion As iClaseConexion)
		  #Pragma Unused sender
		  #Pragma Unused conexion
		  
		  RaiseEvent Conectado(Conexion)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub EventDesconectadoAction(sender As iConexionesGUI)
		  #Pragma Unused sender
		  RaiseEvent EventDesconectado
		  
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event Conectado(Conexion As iClaseConexion)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event EventDesconectado()
	#tag EndHook


	#tag Property, Flags = &h21
		Private ConexionGUI As iConexionesGUI
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Error As iClaseErrorConexion
	#tag EndProperty


#tag EndWindowCode

#tag Events TipoConexionGUI
	#tag Event , Description = 436164612076657A207175652073652063616D626961206C612073656C65636369C3B36E20656E20656C20506F707570204D656EC3BA2E
		Sub Seleccionado(TipoConexion As TiposConexionEnum)
		  //-- LA interface evita la preguntadera! iConexionGUI
		  // esta llamando a la propiedad del tipoconexionGUI
		  
		  If Not(ConexionGUI = Nil) Then
		    Select Case TipoConexion
		    Case TiposConexionEnum.TipoConexionSQL
		      ConexionGUI = New ClaseConexionSQLGUI
		      'RemoveHandler ClaseConexionSQLGUI(ConexionGUI).Conectado, AddressOf ConectadoAction
		    Case TiposConexionEnum.TipoConexionSQLite
		      ConexionGUI = New ClaseConexionSQLiteGUI
		      'RemoveHandler ClaseConexionSQLiteGUI(ConexionGUI).Conectado, AddressOf ConectadoAction
		    Case TiposConexionEnum.TipoConexionAccess
		      ConexionGUI = New ClaseConexionAccessGUI
		      'RemoveHandler ClaseConexionAccessGUI(ConexionGUI).Conectado, AddressOf ConectadoAction
		    End Select
		    
		    ConexionGUI.Close
		    ConexionGUI = Nil
		  End IF
		  
		  //-- Crea la nueva GUI para el tipo de Conexion
		  //-- Fabrica de GUIs
		  
		  Select Case TipoConexion
		  Case TiposConexionEnum.TipoConexionSQL
		    ConexionGUI = New ClaseConexionSQLGUI
		    
		    
		    AddHandler ClaseConexionSQLGUI(ConexionGUI).Conectado, WeakAddressOf ConectadoAction
		    'AddHandler ClaseConexionSQLGUI.EventDesconectado, WeakAddressOf EventDesconectadoAction
		    
		    
		  Case TiposConexionEnum.TipoConexionSQLite
		    ConexionGUI = New ClaseConexionSQLiteGUI
		    AddHandler ClaseConexionSQLiteGUI(ConexionGUI).Conectado, AddressOf ConectadoAction
		  Case TiposConexionEnum.TipoConexionAccess
		    ConexionGUI = New ClaseConexionAccessGUI
		    AddHandler ClaseConexionAccessGUI(ConexionGUI).Conectado, AddressOf ConectadoAction
		  End Select
		  
		  ConexionGUI.EmbedWithin(Self,0,Me.Height,me.Width,Self.Height-me.Height)
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="300"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="300"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="InitialParent"
		Group="Position"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Left"
		Visible=true
		Group="Position"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Top"
		Visible=true
		Group="Position"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockLeft"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockTop"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockRight"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockBottom"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabPanelIndex"
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabIndex"
		Visible=true
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabStop"
		Visible=true
		Group="Position"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Enabled"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="AutoDeactivate"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HelpTag"
		Visible=true
		Group="Appearance"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="UseFocusRing"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackColor"
		Visible=true
		Group="Background"
		InitialValue="&hFFFFFF"
		Type="Color"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		Type="Picture"
		EditorType="Picture"
	#tag EndViewProperty
	#tag ViewProperty
		Name="AcceptFocus"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="AcceptTabs"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="EraseBackground"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Transparent"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="DoubleBuffer"
		Visible=true
		Group="Windows Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
#tag EndViewBehavior
