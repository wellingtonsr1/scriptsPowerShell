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

$Title = "Impressoras" # Título do menu
# Lista com os elementos para a criação do menu
$listMenuElements = @("               $Title                    ", "============================================", 
								"1 - Pesquisar impressora", "2 - Listar impressoras", "Q - Sair")
$server = "localhost"

function list-printers($subtitle){ 
	Clear-Host
	Write-Host "============================================" 
	Write-Host "               $($subtitle)                 "
	Write-Host "============================================" 
	
	# Pega a lista de impressoras e exibe na tela
	Get-WmiObject -Class Win32_Printer -ComputerName $server `
	| Select-Object -Property @{N='Nome';E={$_.name}}, @{N='IP';E={$_.portName}} | Format-Table
	
	Write-Host "============================================" 
	Read-Host -Prompt "Pressione ENTER para continuar"
}

function search-printer($subtitle){
	do{
		Clear-Host
		Write-Host "============================================"
		Write-Host "           $($subtitle)                  "
		Write-Host "============================================"
		$printerName = Read-Host "Nome da impressora ['Q' pra sair]"
		if($printerName -eq 'q'){return}
	}until($printerName -ne '')
	
	# Pega a lista de impressoras (nome e ip)
	$data = Get-WmiObject -Class Win32_Printer -ComputerName $server | Where{$_.Name -like "*$printerName"} `
	| Select-Object -Property name, portName 
	
	# Retornou algum valor?
	if(!$data){
		Write-Host ""
		Write-Host "Impressora nao encontrada!" -ForegroundColor yellow
		Write-Host "============================================" 
		Write-Host ""
		Read-Host -Prompt "Pressione ENTER para continuar"
	}else{
		# Corrige e/ou testa o formato do endereço ip...eliminando valores numéricos extras, ou algum outro caracter.
		$data.portName -match '(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)'
		$ipAddress = $Matches[0]
		
		# Exibe na tela
		$data | Format-Table -AutoSize
		Write-Host "============================================" 
		Write-Host ""
		
		$choice = Read-Host -Prompt "Abrir pagina web da impressora? (S/N)" 
		
		#Abre a página web da impressora
		#if($choice -eq 's'){start "http://$($data.portName.Replace('_1', ''))"}
		if($choice -eq 's'){start "http://$($ipAddress)"}
	}
}

do{
	Clear-Host
	# Cria o menu
	Write-Host "============================================" 
	forEach ($element in $listMenuElements){
		Write-Host $element 
	}
	Write-Host "============================================" 
	$choice = Read-Host "Escolha uma opcao"
    switch ($choice){
		'1' {search-printer -subtitle "Pesquisar impressora"} 
		'2' {list-printers -subtitle "Listar impressora"} 
    }
 }until($choice -eq 'q')
 
 Set-ExecutionPolicy -ExecutionPolicy Default -Scope Process -Force
