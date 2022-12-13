<%
    sub OpenConnection(strServer, strDB, strUser, strPass, intConnTimeOut, intCommTimeOut)
        Dim strConnection, ObjConn
        
        strConnection = "DRIVER={SQL Server Native Client 11.0};SERVER={Server};User ID={User};Password={Pass};Database={strDB}" 
        strConnection = Replace(Replace(Replace(Replace(strConnection,"{Server}", strServer),"{User}", strUser),"{Pass}", strPass),"{strDB}", strDB)

        set ObjConn = Server.CreateObject("ADODB.Connection")
        ObjConn.ConnectionTimeout = intConnTimeOut 
        ObjConn.CommandTimeout = intCommTimeOut        
        ObjConn.Open(strConnection)
    end sub

    Sub CloseConnection()
        ObjConn.Close()
        set ObjConn = nothing
    end Sub   
%>