#region 01 - Anotacoes --------------------------------------------------------------------------------------------------------------------
<# Registrar o Job:
	$trigger0400 = New-JobTrigger -Daily -At 4:00AM
	Register-ScheduledJob -Name SJ7-GestaoGrs -ScriptBlock{ D:\Apps\GestaoGrupos.ps1 } -Trigger $trigger0400
	Get-ScheduledJob | Get-JobTrigger	#ou Get-JobTrigger -Name SchJob1
	Get-ScheduledJob SJ7-GestaoGrs | Get-JobTrigger | Set-JobTrigger -Once -At 12:22PM

Executar sob demanda:
	(Get-ScheduledJob -Name SJ7-GestaoGrs).StartJob()
	Receive-Job -Name SJ7-GestaoGrs
#>
#endregion
#region 01 - Parametros e constantes ------------------------------------------------------------------------------------------------------
if ( $env:computername -eq 'NPAA8050' ) {
	$pastaLocal  = "D:\Apps\Dados\Gerencia"
} else {
	$pastaLocal  = ".\Dados\Gerencia"
}
$grupoSeg = ls $pastaLocal | select @{N='Gru'; E={$_.name.substring(0,$_.name.IndexOf('.'))} } | select -ExpandProperty Gru
#[string[]]$UPSTREAM = Get-Content $pastaLocal\UPSTREAM.TXT
#[string[]]$COMPLETO = Get-Content $pastaLocal\COMPLETO.TXT
#$grupos = @{	'GN_PBI_3SJ0_GERAL_INTER' 		= $INTER
#	'GN_PBI_3SJ0_GERAL_UPSTREAM' 	= $UPSTREAM
#	'GN_PBI_3SJ0_GERAL_COMPLETO' 	= $COMPLETO 	}
#endregion
#region 12 - Credenciais para Conexao
$ConAzureAuto = $true
$UsuarioConectar = "SAPOWERBI"
$usuarioSessao = $env:username
$FilePass1 = "D:\Apps\cred\$($usuarioSessao)_user_azuread_$usuarioConectar.txt" # Salvar a senha: (Get-Credential -Credential $usuarioConectar ).Password | ConvertFrom-SecureString | Out-File $FilePass1
$MeuCred1 = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$usuarioConectar@petrobras.com.br", ( Get-Content $FilePass1 | ConvertTo-SecureString )
#endregion
#region 15 - Connect  AzureAD / AzAccount
if ( $ConAzureAuto ) {
	$conAzure = Connect-AzureAD -Credential $MeuCred1 | fl | Out-String			# MeuCred1 ou MeuCred2
		# Connect-AzAccount -Credential $MeuCred1 | fl | Out-String			# MeuCred1 ou MeuCred2
		#	Connect-AzAccount -DeviceCode
	#write $conAzure
} else {
	try {
		$conAzure = Connect-AzureAD -ErrorAction Ignore
	} catch {
		'erro no Conect-AzureAD. exit'
		exit
	}
}
#endregion
#region 31 - Sincroniza AD x Grupo
foreach ( $grpSeg in $grupoSeg ) {
	[string[]]$gerencias = Get-Content "$pastaLocal\$grpSeg.txt"
	$Usuarios_Completos = @()
	#Obtem Grupo de Visualizadores
	$grupo = Get-AzureADGroup -Filter "DisplayName eq '$grpSeg'"			#$($gp.Name)'"
	#Write-Host "Grupo: $($grupo.DisplayName)"
	#Lista Membros }
	$membros = Get-AzureADGroupMember -ObjectId $grupo.ObjectId -All $true
	Write-Host "Grupo: $($grupo.DisplayName) / Qde Membros: $($membros.Count)"; #read-host
	foreach ($unidade in $gerencias) {
	#Ajuste para verificar se existe * no final do nome do grupo
            if($unidade.Contains('*')) {
                $NovaUnidade = $Unidade.split('*')[0]
                $usuarios = Get-AzureADUser -Filter "startswith(Department,'$NovaUnidade')" -All $true

            } else {
                $usuarios = Get-AzureADUser -Filter "Department eq '$unidade'" -All $true 

            }

			$usuarios = $usuarios | Where-Object PhysicalDeliveryOfficeName -NE $null 
			$Usuarios_Completos += $usuarios
			#Usuarios a Adicionar
			$usuarios | ForEach-Object {
                $adicionar = $false
			    if ( -not $membros ) { $adicionar = $true }
                elseif ( $membros.Count -eq 1 -and ($membros -ne $_) ) { $adicionar = $true }
                    elseif ( $membros.Count -gt 1 -and (-not $membros.Contains($_)) ) { $adicionar = $true }
                if ( $adicionar ) {
					Write-Host "Adicionar " $_.DisplayName "(" $_.Department ")"
					try {
						Add-AzureADGroupMember -ObjectId $grupo.ObjectId -RefObjectId $_.ObjectId
					} catch {}
				}
			}
		}
	}
    #Usuarios a Remover
    #Lista Membros
    $membros = Get-AzureADGroupMember -ObjectId $grupo.ObjectId -All $true
    $membros | ForEach-Object {
        if (-not $usuarios_Completos.Contains($_)){
            Write-Host "Remover " $_.DisplayName "(" $_.Department ")"
            Remove-AzureADGroupMember -ObjectId $grupo.ObjectId -MemberId $_.ObjectId
        }
    }
}
#endregion
