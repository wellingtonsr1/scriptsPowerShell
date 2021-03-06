#+--------------------------------------------------Descrição-------------------------------------------------------+
#| - Lista ou pesquisa impressoras da rede       		                                                            |
#+----------------------------------------------------Autor---------------------------------------------------------+
#| - Wellington        									         	                                                |
#+----------------------------------------------Sistemas testados---------------------------------------------------+
#| - Windows 10(x64)		                    				 		                                            |
#+-------------------------------------------------IMPORTANTE-------------------------------------------------------+
#| - Precisa ser salvo na unidade C:\	                                                                            |
#| - Clicar com o direito, Executar com o powershell                                                                                   |								               
#+------------------------------------------------------------------------------------------------------------------+

$Title = "Impressoras" # Título do menu
# Lista com os elementos para a criação do menu
$listMenuElements = @("               $Title                    ", "============================================", 
								"1 - Pesquisar impressora", "2 - Listar impressoras", "Q - Sair")
$server = "SERVER_IP"
$r = '(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)'


# Retirado de https://github.com/LeDragoX/Win-10-Smart-Debloat-Tools
function Request-PrivilegesElevation(){
	If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
		Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit
	}
}


function List-Printers($subtitle){ 
	Clear-Host
	Write-Host "============================================" 
	Write-Host "               $($subtitle)                 "
	Write-Host "============================================" 
	
	# Pega a lista de impressoras (Nome e IP)
	$data = Get-WmiObject -Class Win32_Printer -ComputerName $server `
	| Select-Object -Property @{N='Nome';E={$_.name}}, @{N='IP';E={$_.portName}}
	
    # Cria uma estrututa de dados com o nome e ip da impressora
    $listPrinterData = @{}
	foreach ($item in $data) {
		if($item.IP -match $r){$listPrinterData.add($item.Nome, $Matches[0])}
	}

	# Exibe os dados na tela
    foreach ($nameAndIp in $listPrinterData.GetEnumerator() | Sort-Object -Property Value) {
        Write-Host $nameAndIp.key '-->' $nameAndIp.value -ForegroundColor green
    }

	Write-Host "============================================" 
	Read-Host -Prompt "Pressione ENTER para continuar"
}

function Search-Printer($subtitle){
	do{
		Clear-Host
		Write-Host "============================================"
		Write-Host "           $($subtitle)                  "
		Write-Host "============================================"
		$printerName = Read-Host "Nome da impressora ['Q' pra sair]"
		if($printerName -eq 'q'){return}
	}until($printerName -ne '')
	
	# Pega a lista de impressoras (Nome e IP)
	$data = Get-WmiObject -Class Win32_Printer -ComputerName $server | Where{$_.Name -like "*$printerName"} `
	| Select-Object -Property name, portName 
	
	# Retornou algum valor?
	if(!$data){
		Write-Host ""
		Write-Host "Impressora não encontrada!" -ForegroundColor yellow
		Write-Host "============================================" 
		Write-Host ""
		Read-Host -Prompt "Pressione ENTER para continuar"
	}else{
		# Corrige e/ou testa o formato do endereço ip...eliminando valores numéricos extras, ou algum outro caracter.
		$null = $data.portName -match $r
		$ipAddress = $Matches[0]
		
        # Cria uma estrututa de dados com o nome e ip da impressora
        $listPrinterData = @{}
		$listPrinterData.add($data.name, $ipAddress)

        # Exibe os dados na tela
        foreach ($nameAndIp in $listPrinterData.GetEnumerator()) {
            Write-Host ""
            Write-Host $nameAndIp.key '-->' $nameAndIp.value -ForegroundColor green
            Write-Host ""
        }
		
		Write-Host "============================================" 
		Write-Host ""
		
		$choice = Read-Host -Prompt "Abrir página web da impressora? (S/N)" 
		# Abre a página web da impressora
		if($choice -eq 's'){start "http://$($ipAddress)"}
	}
}

function Create-Menu(){
	Request-PrivilegesElevation

	do{
		Clear-Host
		# Cria o menu
		Write-Host "============================================" 
		forEach ($element in $listMenuElements){
			Write-Host $element 
		}
		Write-Host "============================================" 
		$choice = Read-Host "Escolha uma opção"
		switch ($choice){
			'1' {Search-Printer -subtitle "Pesquisar impressora"} 
			'2' {List-Printers -subtitle "Listar impressora"} 
		}
	}until($choice -eq 'q')
 }

Create-Menu