#tag Class
Protected Class ClaseConexionAccess
Inherits ClaseErrorConexion
Implements iClaseConexion
	#tag Method, Flags = &h21
		Private Sub Conectar()
		  mConectado = False
		  
		  Dim ConnectionString As String = kCONNECTION_STRING
		  Dim ArchivoAccess As New FolderItem(mDatabaseFile)
		  
		  Try
		    
		    //-- Valida el Archivo Existe
		    If Not (ArchivoAccess.Exists) Then
		      Error = True
		      ErrorNumero = -300
		      ErrorMensaje = "El archivo de la base de datos no existe."
		      ErrorDetalles = ArchivoAccess.NativePath
		      Return
		    End If
		    
		    //-- Valida que el archivo se pueda leer 
		    If Not(ArchivoAccess.IsReadable) Then
		      Error = True
		      ErrorNumero = -400
		      ErrorMensaje = "No cuenta con permisos de lectura para el archivo de la base de datos."
		      ErrorDetalles = ArchivoAccess.NativePath
		      Return
		    End If
		    
		    '-- Crea el connection string para el odbc de access
		    ConnectionString = ConnectionString.Replace("<DatabaseFile>", mDatabaseFile)
		    ConnectionString = ConnectionString.Replace("<UserName>", mUserName)
		    ConnectionString = ConnectionString.Replace("<Password>", mPassword)
		    
		    innerConexion.DataSource = ConnectionString
		    
		    mConectado = innerConexion.Connect
		    
		    if innerConexion.Error Then
		      Error = True
		      ErrorNumero = innerConexion.ErrorCode
		      ErrorMensaje = "Error al tratar de conectar con la base de datos de Access."
		      ErrorDetalles = innerConexion.ErrorMessage
		    End if
		    
		  Catch 
		    Error = True
		    ErrorNumero = -999
		    ErrorMensaje = "Error inesperado al tratar de conectar con la base de datos de Access."
		    ErrorDetalles = ArchivoAccess.NativePath
		  End Try
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(parametros As string)
		  // Part of the iClaseConexion interface.
		  mConectado = False
		  LeerParametros(parametros)
		  innerConexion = New ODBCDatabase
		  Conectar()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Desconectar()
		  If mConectado Then innerConexion.Close
		  mConectado = False
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  // Part of the iClaseConexion interface.
		  if mConectado Then Desconectar
		  innerConexion = Nil
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub LeerParametros(parametros As string)
		  Try
		    Dim Partes() As String = Parametros.Split(";")
		    For Each Parte As String In Partes
		      If Parte <> "" Then 
		        Dim ParteValor() As String = Parte.Trim.Split("=")
		        Select Case ParteValor(0).Uppercase
		        Case "DATABASEFILE"
		          mDatabaseFile = ParteValor(1)
		        Case "USERNAME"
		          mUserName = ParteValor(1)
		        Case "PASSWORD"
		          mPassword = ParteValor(1)
		        End Select
		      End If
		    Next Parte
		  Catch
		    Error = True
		    ErrorNumero = -999
		    ErrorMensaje = "Error inesperado."
		    ErrorDetalles = "Al tratar de interpretar los parámetros de la conexión."
		  End Try
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerBasesDatos() As Dictionary
		  // Part of the iClaseConexion interface.
		  //-- Sólo existe una base de datos
		  //-- Para obtener el nombre lo sacaré del nombre del archivo
		  Dim Resultado As New Dictionary
		  
		  If Not mConectado then Return Resultado
		  
		  Try
		    Dim Archivo As New FolderItem(mDatabaseFile)
		    Resultado.Value(1) = Archivo.DisplayName
		  Catch
		    Error = True
		    ErrorNumero = -600
		    ErrorMensaje = "Error inesperado al obtener el nombre de la base de datos."
		    ErrorDetalles = "Archivo: " + mDatabaseFile
		  End Try
		  
		  Return Resultado
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerCampos(BaseDatos As String, Tabla As String) As Dictionary
		  // Part of the iClaseConexion interface.
		  
		  #Pragma Unused BaseDatos
		  #Pragma Unused Tabla
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerTablas(BaseDatos As String) As Dictionary
		  // Part of the iClaseConexion interface.
		  #Pragma Unused BaseDatos
		  
		  If Not mConectado then Return Nil
		  
		  Dim Resultado As New Dictionary
		  Dim NumTab As Integer = 0
		  Dim rs As RecordSet = innerConexion.TableSchema
		  
		  If innerConexion.Error then
		    ErrorNumero = innerConexion.ErrorCode
		    ErrorMensaje = "Error al ejecutar la solicitud de datos."
		    ErrorDetalles = innerConexion.ErrorMessage
		    Return Nil
		  End if
		  
		  //-- Busca las Tablas
		  While Not rs.EOF
		    Dim NombreTabla As String = rs.Field("TableName").StringValue
		    NumTab = NumTab + 1
		    Resultado.Value(NumTab) = NombreTabla
		    rs.MoveNext
		  Wend
		  
		  rs.Close
		  rs = Nil
		  
		  Return Resultado
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Tipo() As TiposConexionEnum
		  Return TiposConexionEnum.TipoConexionAccess
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private innerConexion As ODBCDatabase
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mConectado As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDatabaseFile As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPassword As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mUserName As String = "Admin"
	#tag EndProperty


	#tag Constant, Name = kCONNECTION_STRING, Type = String, Dynamic = False, Default = \"Driver\x3D{Microsoft Access Driver (*.mdb\x2C *.accdb)};Dbq\x3D<DatabaseFile>;Uid\x3D<UserName>;Pwd\x3D<Password>;", Scope = Private
	#tag EndConstant


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
