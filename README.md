# vb6-tcp-client-server
VB6 CLient-Server example for OstroSoft Winsock Component

Example demonstrates using the component in client-server scenario.
Server would echo a client request prefixed by "you sent " (echo server).

On project start, an instance of the component listens on port 22222 (can be changed in code). 
From the server form you can start pre-defined number of clients (3 default is 3).

On connection request server redirects the request to one of its available listeners, while keep listening on port 22222. If no listener is available, server will create a new one.

From this point on, listener will handle interaction with the corresponding client. Once client disconnects, listener becomes available for a new connection.

Since VB does not accept events from object arrays, it's necessary to create a separate class for the listener (clsTCPServerListener) triggering events on the main form (frmTCPServer).

Please keep in mind that sockets are resource-heavy and there is a limit on how many you can have opened (32767, to be precise). Don't forget to properly close client connection (with CloseWinsock method), so listener would become available for the next request.

The latest version of OstroSoft Winsock Component is available for download at http://www.ostrosoft.com/oswinsck.aspx
