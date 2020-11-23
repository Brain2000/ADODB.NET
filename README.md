# ADODB.NET
Fully managed .NET rendition of ADODB Connections/Recordsets/Commands with both client and server side cursors.

# Example:
```
using (var DB = new ADODB.Connection())
{
    if (DB.Open("PROVIDER=sqloledb;NETWORK=DBMSSOCN;SERVER=MyServer\SQLInstance;DATABASE=coverstone;UID=coverstone;PWD=xxxxxxxxxxxxxxx"))
    {
        using (DYN = new ADODB.Recordset())
        {
            //will always use client side cursor for static/readonly, so no need to specify
            DYN.Open("SELECT * FROM User WHERE UserName='Fred'", DB, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly);
            if (!DYN.EOF)
            {
                string name = (string)DYN["UserName"].Value;
                int age = (int)DYN["Age"].Value;
            }
            DYN.Close();
        }
    }

    using (DYN = new ADODB.Recordset())
    {
        //will always use server side cursor by default for non-static/readonly
        //Recordsets will inherit this from the Connection, this can be overriden by setting CursorLocation on either the connection, or the Recordset
        //once you override CursorLocation on a Recordset, it will no longer inherit from the Connection object
        DYN.Open("SELECT * FROM User WHERE 0=1", DB, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic);
        DYN.AddNew();
        DYN["User_Id"] = 1234L;
        DYN["UserName"] = "Fred";
        DYN.Update();
        DYN.Close();
    }
    
    DB.Close(); //if you forget to close this, the Dispose() will also handle it
}
```
