#+--------------------------------------------------Descrição-------------------------------------------------------+
#| - Lista ou pesquisa impressoras da rede       		                                                            |
#+----------------------------------------------------Autor---------------------------------------------------------+
#| - Wellington        									         	                                                |
#+----------------------------------------------Sistemas testados---------------------------------------------------+
#| - Windows 10(x64)		                    				 		                                            |
#+-------------------------------------------------IMPORTANTE-------------------------------------------------------+
#| - Precisa ser salvo na unidade C:\	                                                                            |
#| - Executar o 'impressoras.bat'                                                                                   |								               
#+------------------------------------------------------------------------------------------------------------------+

$servidor = "10.39.0.126"

function listar-impressoras{
	Clear-Host
	Write-Host "======================================"
	Write-Host "        Lista de impressora           "
	Write-Host "======================================"
	
	Get-WmiObject -Class Win32_Printer -ComputerName $servidor `
	| Select-Object -Property @{N='Nome';E={$_.sharename}}, @{N='IP';E={$_.portName}} | Format-Table
	
	Write-Host "======================================"
	Read-Host -Prompt "Pressione ENTER para continuar"
}

function pesqusiar-impressora{
	Clear-Host
	Write-Host "======================================"
	Write-Host "        Pesquisar impressora          "
	Write-Host "======================================"
	
	$nameImpressora = Read-Host "Informe o nome da impressora"
	Get-WmiObject -Class Win32_Printer -ComputerName $servidor | Where{$_.Name -like "*$nameImpressora*"} `
	| Select-Object -Property @{N='Nome';E={$_.sharename}}, @{N='IP';E={$_.portName}} | Format-Table
	
	Write-Host "======================================"
	Read-Host -Prompt "Pressione ENTER para continuar"
}

do{
	Clear-Host
	Write-Host "======================================"
	Write-Host "            Impressoras               "
	Write-Host "======================================"
 
	Write-Host "1 - Listar impressoras"
	Write-Host "2 - Pesquisar impressora"
	Write-Host "Q - Sair"
	Write-Host "======================================"
	$selecao = Read-Host "Escolha uma opcao"
	
    switch ($selecao){
		'1' {listar-impressoras} 
		'2' {pesqusiar-impressora} 
    }
 }until($selecao -eq 'q')
 
 Set-ExecutionPolicy -ExecutionPolicy Default -Scope Process -Force
