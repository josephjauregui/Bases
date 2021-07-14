#tag Window
Begin ContainerControl ClaseConexionAccessGUI Implements iConexionesGUI
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
   Height          =   320
   HelpTag         =   ""
   InitialParent   =   ""
   Left            =   0
   LockBottom      =   True
   LockLeft        =   True
   LockRight       =   True
   LockTop         =   True
   TabIndex        =   0
   TabPanelIndex   =   0
   TabStop         =   True
   Top             =   0
   Transparent     =   True
   UseFocusRing    =   False
   Visible         =   True
   Width           =   330
   Begin Label UsuarioLabel
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Usuario:"
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   139
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   113
   End
   Begin Label TipoConexionLabel
      AutoDeactivate  =   True
      Bold            =   True
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Conexión a Microsoft Access©"
      TextAlign       =   1
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   50
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   290
   End
   Begin Label ConectarseLabel
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Conectarse..."
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   20.0
      TextUnit        =   0
      Top             =   20
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   290
   End
   Begin Label ArchivoLabel
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Archivo de la base de datos:"
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   81
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   290
      Begin CheckBox CheckBoxUsuario
         AutoDeactivate  =   True
         Bold            =   False
         Caption         =   "Activar Usuario"
         DataField       =   ""
         DataSource      =   ""
         Enabled         =   True
         Height          =   20
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "ArchivoLabel"
         Italic          =   False
         Left            =   209
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   2
         State           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         TextFont        =   "System"
         TextSize        =   0.0
         TextUnit        =   0
         Top             =   81
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   101
      End
   End
   Begin TextField ArchivoTextField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "Ruta y Nombre del Archivo (*.accdb, *.mdb)"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   False
      Format          =   ""
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   105
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   268
   End
   Begin TextField UsuarioTextField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "Nombre del usuaario"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   False
      Format          =   ""
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Admin"
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   165
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   290
   End
   Begin Label ContrasenaLabel
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Contraseña:"
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   199
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   290
   End
   Begin TextField ContrasenaTextField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "Contraseña"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   False
      Format          =   ""
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   True
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   225
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   290
   End
   Begin PushButton ConectarPushButton
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Conectar"
      Default         =   True
      Enabled         =   False
      Height          =   32
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   110
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   268
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   110
   End
   Begin PushButton BuscarArchivoPushButton
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "..."
      Default         =   True
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   288
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   105
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   22
   End
   Begin Canvas CanvasBD
      AcceptFocus     =   False
      AcceptTabs      =   False
      AutoDeactivate  =   True
      Backdrop        =   2125283327
      DoubleBuffer    =   False
      Enabled         =   True
      EraseBackground =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   275
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   10
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   20
      Transparent     =   True
      UseFocusRing    =   True
      Visible         =   False
      Width           =   35
   End
End
#tag EndWindow

#tag WindowCode
	#tag Method, Flags = &h0
		Sub Close()
		  Super.Close
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ConectadoAdentro()
		  ConectarPushButton.Caption = "Desconectar"
		  BuscarArchivoPushButton.Enabled = False
		  ContrasenaTextField.Enabled = False
		  UsuarioTextField.Enabled = False
		  CheckBoxUsuario.Enabled = False
		  ConectarseLabel.Text = "Conectado..."
		  CanvasBD.Visible = True
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  //-- Inicializa la ruta del archivo
		  
		  mArchivoBaseDatos = SpecialFolder.Documents
		  ArchivoTextField.Text = mArchivoBaseDatos.NativePath
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Desconectado()
		  ConectarPushButton.Caption = "Conectar"
		  BuscarArchivoPushButton.Enabled = true
		  CheckBoxUsuario.Enabled = True
		  CheckBoxUsuario.Value = False
		  CanvasBD.Visible = False
		  ConectarseLabel.Text = "Conectarse..."
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EmbedWithin(containingControl As RectControl, left As Integer = 0, top As Integer = 0, width As Integer = - 1, height As Integer = - 1)
		  Super.EmbedWithin(containingControl, left, top, width, height)
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EmbedWithin(containingWindow As Window, left As Integer = 0, top As Integer = 0, width As Integer = - 1, height As Integer = - 1)
		  Super.EmbedWithin(containingWindow, left, top, width, height)
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SeConecto()
		  If ConectarPushButton.Caption = "Conectar" Then
		    
		    Dim Conexion As iClaseConexion 
		    Dim Parametros As String = "DatabaseFile=" + ArchivoTextField.Text + ";"
		    Dim CadenaConexion As String 
		    
		    CadenaConexion = CadenaConexion + "username=" + UsuarioTextField.Text + ";"
		    CadenaConexion = CadenaConexion + "password=" + ContrasenaTextField.Text + ";"
		    
		    Conexion = New ClaseConexionAccess(Parametros)
		    
		    If Conexion.Error Then
		      Dim Error As iClaseErrorConexion = New ClaseErrorConexion
		      ClaseErrorConexion(Error).Error = True
		      ClaseErrorConexion(Error).ErrorNumero = Conexion.ErrorNumero
		      ClaseErrorConexion(Error).ErrorMensaje = Conexion.ErrorMensaje 
		      ClaseErrorConexion(Error).ErrorDetalles = Conexion.ErrorDetalles
		      
		      'RaiseEvent NoConectado(Error)
		    Else 
		      ConectadoAdentro()
		      RaiseEvent Conectado(Conexion)
		      Conexion = Nil
		    End If
		  Else
		    Desconectado()
		    'RaiseEvent Desconectar
		  End If
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event Conectado(Conexion As iClaseConexion)
	#tag EndHook


	#tag Property, Flags = &h21
		Private mArchivoBaseDatos As FolderItem
	#tag EndProperty


#tag EndWindowCode

#tag Events CheckBoxUsuario
	#tag Event
		Sub Action()
		  if CheckBoxUsuario.Value Then
		    UsuarioTextField.Enabled = True
		    ContrasenaTextField.Enabled = True
		  else
		    UsuarioTextField.Enabled = False
		    ContrasenaTextField.Enabled = False
		  end if
		  
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ConectarPushButton
	#tag Event
		Sub Action()
		  SeConecto()
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events BuscarArchivoPushButton
	#tag Event
		Sub Action()
		  Dim dlg As New OpenDialog
		  
		  dlg.InitialDirectory = mArchivoBaseDatos
		  
		  dlg.Title = "Selecciona el archivo de la base de datos de Access"
		  
		  dlg.SuggestedFileName = "*.db"
		  dlg.PromptText = "Archivo de Access"
		  dlg.ActionButtonCaption = "Seleccionar"
		  dlg.CancelButtonCaption = "Cancelar"
		  dlg.MultiSelect = False
		  dlg.Filter = BasesDatosTypes.Acces.Extensions
		  
		  mArchivoBaseDatos = dlg.ShowModal
		  
		  If mArchivoBaseDatos <> Nil Then
		    If not mArchivoBaseDatos.Directory Then
		      ArchivoTextField.Text =  mArchivoBaseDatos.NativePath
		      ConectarPushButton.Enabled = True
		    Else
		      ConectarPushButton.Enabled = False
		    End If
		  End If
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
