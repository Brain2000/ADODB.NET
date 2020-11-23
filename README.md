# ADODB.NET
Fully managed .NET rendition of ADODB Connections/Recordsets/Commands with both client and server side cursors.

So why not just use the built in ADODB COM object? It crashes from time to time if using multiple threads. Even if each created instance stays on its own thread and is not shared, it will still crash.

This version is rock solid.  Note: it doesn't mean this is thread safe where you can pass instances between threads and call the functions simultaneously, as the underlying ADO.NET is also not thread safe in this fashion. You can have 1000 threads creating their own instances simultaneously and there will not be any crashing. Or a thread can hand an instance to another thread, so long as the original thread does not call any functions simultaneously with the secondary thread!

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
