# ADODB.NET
Fully managed .NET rendition of ADODB Connections/Recordsets/Commands with both client and server side cursors.

So why not just use the built in ADODB COM object? It crashes from time to time if using multiple threads. Even if each created instance stays on its own thread and is not shared, it will still crash.

This version is rock solid.  Note: it doesn't mean this is thread safe where you can pass instances between threads and call the functions simultaneously, as the underlying ADO.NET is also not thread safe in this fashion. You can have 1000 threads creating their own instances simultaneously and there will not be any crashing. Or a thread can hand an instance to another thread, so long as the original thread does not call any functions simultaneously with the secondary thread!

There are a number of things that this rendition does not fully support, such as batch optimistic (actually, I'm not even sure what that does).

Quite a bit of this has been tested against the COM ADODB version, and it even mirrors a lot of the errors verbatim, and the circumstances that they occur under.
It was designed to be a drop-in replacement for COM ADODB with minimal changes.

Server side cursors were reverse engineered into their internal stored procedures that they call. They are still one of the safest ways to run a single update against a database without worry of accidentally updating other rows because of a malformed WHERE clause that matches too many items. If you try to run Updates via a client side recordset, it will work, but I believe ADO.NET will throw errors unless you have also selected all primary key fields.

There are also some enhancements in this code:

1) The "Recordset.updatedFields" property is a dictionary that keeps track of fields that have been updated. It's smart enough to know if you update a field and set it to the same value that is already set. This in turn is used when running updates, as it it optimized and will only update fields that have actually changed.

2) The "Recordset.accessedFields" property is a dictionary that keeps track of fields that have been simply accessed. This can be useful if you want to keep track of fields that are touched in order to optimize queries later on.

3) Connection.AcquireAppLock( ) and Connection.ReleaseAppLock( ) are two functions that can be used to synchronize calls to a database across a server farm. It's like a Mutex that exists across servers.

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
        DYN["User_Id"].Value = 1234L;
        DYN["UserName"].Value = "Fred";
        DYN.Update();
        DYN.Close();
    }
    
    DB.Close(); //if you forget to close this, the Dispose() will also handle it
}
```
