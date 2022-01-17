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

$Titulo = "Impressoras" # Título do menu
# Lista com os elementos para a criação do menu
$listaElementosMenu = @("            $Titulo                   ", "======================================", 
								"1 - Listar impressoras", "2 - Pesquisar impressora", "Q - Sair")
$servidor = "10.39.0.126" # Endereço do servior que contém as impressoras instaladas


function exibir($selecao, $subtitulo){
	Clear-Host
	Write-Host "======================================"
	Write-Host "         $subtitulo                   "
	Write-Host "======================================"
	
	if($selecao -eq '1'){
		$nomeImpressora = ""
	}elseif($selecao -eq '2'){
		$nomeImpressora = Read-Host "Informe o nome da impressora"
	}
	
	# Cria e exibe uma lista com as impressoras
	Get-WmiObject -Class Win32_Printer -ComputerName $servidor | Where{$_.Name -like "*$nomeImpressora*"} `
	| Select-Object -Property @{N='Nome';E={$_.sharename}}, @{N='IP';E={$_.portName}} | Format-Table -AutoSize
	
	Write-Host "======================================"
	Read-Host -Prompt "Pressione ENTER para continuar" 
}

do{
	Clear-Host
	# Cria o menu
	Write-Host "======================================"
	forEach ($elemento in $listaElementosMenu){
		Write-Host $elemento
	}
	Write-Host "======================================"
	
	$selecao = Read-Host "Escolha uma opcao"
	switch ($selecao){
		'1' {exibir -selecao $selecao -subtitulo "Lista de impressora"} 
		'2' {exibir -selecao $selecao -subtitulo "Pesquisar impressora"} 
    }
 }until($selecao -eq 'q')
 
 # Configura a política de execução como a padrão do sistema
 Set-ExecutionPolicy -ExecutionPolicy Default -Scope Process -Force
