option explicit 
dim testetecla ,objFSO, Wshs, usrProfile, varShell, vardata, frase, novaSubPasta        

testetecla = msgbox ("Sr. usu�rio: Bem vindo. Clique em Sim para gravar um arquivo de backup.",3,"SISTEMA DE ALUNOS")

Set varShell = wscript.CreateObject("WScript.Shell")   'Cria uma inst�ncia de objeto  

if testetecla = 2 then wscript.quit 
if testetecla = 6 then   ' o arquivo backup ser� criado...           

    varData = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now)  
    'msgbox varData  

    frase = "cmd  /k  CD  C:\&  mkdir  c:\copia\" & vardata & " & exit"   ' monta string cria��o nova pasta     
    'msgbox frase 

    varShell.run frase,1,true   ' Cria no drive "c" uma pasta de nome "copia" e sub pasta "yyyymmddhhmmss"                         ' Far� backup do arquivo  

    novaSubPasta = "c:\copia\" & varData & "\anterior.accdb"     
    'msgbox novaSubPasta

    Set objFSO = CreateObject("Scripting.FileSystemObject")     
    Set Wshs = WScript.CreateObject("WScript.Shell")     
    usrProfile = Wshs.ExpandEnvironmentStrings("%UserProfile%")     

    objFSO.CopyFile "C:\Users\LeoBalestra\Desktop\Trabalho\Setings\TrabalhoFinal.accdb" , novaSubPasta      
    msgbox "Arquivo de backup gravado..."            

end if  

                           'execu��o do Access
Set varShell = wscript.CreateObject("WScript.Shell")   'Cria uma inst�ncia de objeto  
varShell.run ("""msaccess.exe""C:\Users\LeoBalestra\Desktop\Trabalho\Setings\TrabalhoFinal.accdb"),3,true     
if testetecla = 6 then wscript.Echo "Sr. Usu�rio... foi gravado o arquivo backup, favor providenciar armazenamento em local seguro!" 
'msgbox "Opera��o realizada com sucesso!",0,"Sistema Banco de Dados"