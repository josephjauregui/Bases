#tag Class
Protected Class ClaseConexionSQL
Inherits ClaseErrorConexion
Implements iClaseConexion
	#tag Method, Flags = &h21
		Private Sub Conectar()
		  'conexion  
		  Try
		    innerConexion.Host = mHost
		    innerConexion.UserName = mUserName
		    innerConexion.Password = mPassword
		    innerConexion.DatabaseName = "master"
		    innerConexion.TimeOut = 2
		    
		    '-- Revisa el puerto 
		    If Not (mPort = 0) Then innerConexion.Port = mPort
		    
		    '-- Trata de conectar
		    If Not innerConexion.Connect Then
		      '-- No pudo conectarse
		      Error = True
		      ErrorNumero = innerConexion.ErrorCode
		      ErrorMensaje = "Error al conectar con el servidor de SQL Server " + mHost + "."
		      ErrorDetalles = innerConexion.ErrorMessage
		      mConectado = False
		      Return
		    End If
		    
		    '-- Conectado
		    mConectado= True
		    
		  Catch 
		    '-- Error desconocido
		    Error = True
		    ErrorNumero = innerConexion.ErrorCode
		    ErrorMensaje = "Error inesperado al conectar con el servidor de SQL Server " + mHost + "."
		    ErrorDetalles = innerConexion.ErrorMessage
		    mConectado= False
		    Return
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(parametros As string)
		  '-- Se elabora un Constructor generico para que combine con otras bases de datos
		  '-- Se pasan los valores de la conexion como cadena en pares
		  '-- Ejemplo: parametros = "servidor=.\compac;usuario=sa;contrasena=Compac08;puerto=0"
		  
		  mConectado = False
		  LeerParametros(Parametros)
		  innerConexion = New MSSQLServerDatabase
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
		  if mConectado Then Desconectar
		  innerConexion = Nil
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub LeerParametros(parametros As string)
		  Dim Partes() As String = Parametros.Split(";")
		  For Each Parte As String In Partes
		    
		    If Not (Parte = "") Then
		      //-- No está vacio el parámetro
		      
		      Dim ParteValor() As String = Parte.Trim.Split("=")
		      
		      Select Case ParteValor(0).Uppercase
		      Case "HOST"
		        mHost = ParteValor(1)
		      Case "USERNAME"
		        mUserName = ParteValor(1)
		      Case "PASSWORD"
		        mPassword = ParteValor(1)
		      Case "PORT"
		        mPort = ParteValor(1).Val
		      End Select
		      
		    End If
		    
		  Next Parte
		  
		  
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
		  Dim Resultado As New Dictionary
		  If Not mConectado then Return Resultado
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ObtenerTablas(BaseDatos As String) As Dictionary
		  Dim Resultado As New Dictionary
		  
		  If Not mConectado then Return Resultado
		  
		  Dim Query As String = kQUERY_TABLES.Replace("<DATABASE_NAME>", BaseDatos)
		  Dim rs As RecordSet = innerConexion.SQLSelect(Query)
		  
		  Dim NumTab As Integer = 0
		  
		  While Not rs.EOF
		    NumTab = NumTab + 1
		    Resultado.Value(NumTab) = rs.Field("TABLE_NAME").StringValue
		    rs.MoveNext
		  Wend
		  
		  rs.Close
		  rs = Nil
		  
		  Return Resultado
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Tipo() As TiposConexionEnum
		  Return TiposConexionEnum.TipoConexionSQL
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		innerConexion As MSSQLServerDatabase
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mConectado As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHost As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPassword As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPort As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mUserName As String
	#tag EndProperty


	#tag Constant, Name = kQUERY_DATABASES, Type = String, Dynamic = False, Default = \"SELECT name FROM master.dbo.sysdatabases", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kQUERY_FIELDS, Type = String, Dynamic = False, Default = \"Aca va la otra query", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kQUERY_TABLES, Type = String, Dynamic = False, Default = \"SELECT TABLE_NAME \rFROM [<DATABASE_NAME>].INFORMATION_SCHEMA.TABLES \rWHERE TABLE_TYPE \x3D \'BASE TABLE\'", Scope = Private
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
