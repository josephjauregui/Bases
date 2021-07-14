#tag Class
Protected Class ClaseConexionSQLite
Inherits ClaseErrorConexion
Implements iClaseConexion
	#tag Method, Flags = &h21
		Private Sub Conectar()
		  mConectado = False
		  innerConexion.DatabaseFile = new FolderItem(mDatabaseFile)
		  
		  //-- Valida el Archivo Existe
		  If Not (innerConexion.DatabaseFile.Exists) Then
		    Error = True
		    
		    ErrorMensaje = "El archivo de la base de datos no existe."
		    ErrorDetalles = innerConexion.DatabaseFile.NativePath
		    Return
		  End If
		  
		  //-- Valida que el archivo se pueda leer 
		  If Not(innerConexion.DatabaseFile.IsReadable) Then
		    Error = True
		    ErrorMensaje = "No cuenta con permisos de lectura para el archivo de la base de datos."
		    ErrorDetalles = innerConexion.DatabaseFile.NativePath
		    Return
		  End If
		  
		  try
		    
		    mConectado = innerConexion.Connect
		    
		    If innerConexion.Error Then
		      Error = True
		      ErrorMensaje = innerConexion.ErrorMessage
		      ErrorDetalles = innerConexion.DatabaseFile.NativePath
		      Return
		    End If
		    
		    innerConexion.MultiUser = True
		    
		  Catch
		    
		    Error = True
		    ErrorMensaje = "Error inesperado. "
		    ErrorDetalles = innerConexion.DatabaseFile.NativePath
		    
		    mConectado = False
		    
		  end try
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(parametros As string)
		  '-- Se elabora un Constructor generico para que combine con otras bases de datos
		  '-- Se pasan los valores de la conexion como cadena en pares
		  '-- Ejemplo: parametros = "servidor=.\compac;usuario=sa;contrasena=Compac08;puerto=0"
		  
		  mConectado = False
		  LeerParametros(Parametros)
		  innerConexion = New SQLiteDatabase
		  Conectar()
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Desconectar()
		  innerConexion.Close
		  mConectado = False
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  Desconectar
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
		  Dim Resultado As New Dictionary
		  If Not(mConectado) then Return Resultado
		  
		  Dim rs As RecordSet = innerConexion.SQLSelect(kQUERY_DATABASES)
		  
		  Dim NumReg As Integer = 0
		  While Not rs.EOF
		    NumReg = NumReg + 1
		    Resultado.Value(NumReg) = rs.Field("name").StringValue
		    rs.MoveNext
		  Wend
		  
		  rs.Close
		  rs = Nil
		  
		  Return Resultado
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerCampos(BaseDatos As String, Tabla As String) As Dictionary
		  #Pragma Unused BaseDatos
		  #Pragma Unused Tabla
		  
		  Dim Resultado As New Dictionary
		  If Not mConectado then Return Resultado
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerTablas(BaseDatos As String) As Dictionary
		  #Pragma Unused BaseDatos 
		  Dim Resultado As New Dictionary
		  
		  If Not mConectado then Return Resultado
		  
		  Dim Query As String = kQUERY_TABLES
		  Dim rs As RecordSet = innerConexion.SQLSelect(Query)
		  
		  Dim NumTab As Integer = 0
		  
		  While Not rs.EOF
		    NumTab = NumTab + 1
		    Resultado.Value(NumTab) = rs.Field("name").StringValue
		    rs.MoveNext
		  Wend
		  
		  rs.Close
		  rs = Nil
		  
		  Return Resultado
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Tipo() As TiposConexionEnum
		  Return TiposConexionEnum.TipoConexionSQLite
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		innerConexion As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mConectado As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDatabaseFile As String
	#tag EndProperty


	#tag Constant, Name = kQUERY_DATABASES, Type = String, Dynamic = False, Default = \"PRAGMA database_list;", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kQUERY_FIELDS, Type = String, Dynamic = False, Default = \"Aca va la otra query", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kQUERY_TABLES, Type = String, Dynamic = False, Default = \"SELECT \tname \rFROM \tsqlite_master \rWHERE \ttype\x3D\'table\' AND \r\tNOT( name LIKE \'sqlite%\');", Scope = Private
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
