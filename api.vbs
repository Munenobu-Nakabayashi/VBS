option explicit

dim args
dim method
dim url
dim oHTTP

set args = wscript.arguments

if args.count < 1 then
   wscript.echo "No URL provided"
   wscript.quit
end if

set oHTTP = createObject("MSXML2.ServerXMLHTTP")

method = "GET"
url    = args(0)

wscript.echo "method: " & method
wscript.echo "URL: " & url

oHTTP.open method, url, false
oHTTP.send

wscript.echo "Status: " & oHTTP.status & " " & oHTTP.statusText
wscript.echo oHTTP.responseText