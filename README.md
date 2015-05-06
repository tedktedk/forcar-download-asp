
#Forçando Downloads usando ASP

Utilizando o Content-Type "application/x-msdownload" para forçar os downloads

```sh
public function download(arquivo, pasta)
    dim objStream
    set objStream = server.createObject("ADODB.Stream")
    with (response)
        .buffer = true  
        .addHeader "Content-Type","application/x-msdownload"
        .addHeader "Content-Disposition","attachment; filename="&arquivo  
        .flush  
    end with
    with (objStream)
        .open  
        .type = 1  
        .loadFromFile server.mapPath(pasta)
    end with
    response.binaryWrite objStream.read
    set objStream = nothing
    response.flush
end function
```
