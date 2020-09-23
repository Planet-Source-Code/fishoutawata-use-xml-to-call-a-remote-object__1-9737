Thank you for downloading this example.  Below are the instructions for
installation....

Included in this Download.

---- Source Code ----
[Project1] -- SoapDest.vbp
SoapDest.vbp
cStringTest.cls

[Project2] -- RPC.vbp
RPC.vbp
MethodWrapper.cls

[Project3] -- prjTest.vbp
prjTest.vbp
frmMain.frm

[Asp File]
rpcxml.asp


---- Binary Relese ----
SoapDest.dll
RPC.dll


[Required to run this Example]
Web Server that supports ASP (i.e. Personal Web Server, IIS)


[Installation]

1. Copy SoapDest.dll and RPC.dll to your system directory (i.e. windows\system).  
 
2. Register both dlls with regsvr32 ( i.e. regsvr32 (yourdll.dll) )

3. Copy rpcxml.asp to your webserver's html directory (i.e. c:\inetpub\wwwroot)

4. Modify Click event of Send Button in prjTest to point to rpcxml.asp on your server. (or you could just leave it as is and test against my web server) :)

5. Run prjTest and see if it works.


 