#region 0 - Descrição, Parâmetos, Constantes, Funções ------------------------------------------------------------------------------------------
	#region Descrição
		# Script para criação de PowerBI Service Workspace
		#	Exemplos:
		#				. .\WorkSpace.ps1 -Oferta AS-CS_POWERBI_APL -CodApl YGLQ -Registro S1094089 -ChaveRegistro ZCW5 -Premium CL_002
		#				. .\WorkSpace.ps1 -Oferta AS-CS_POWERBI_APL -CodApl N1JV -Registro S0000011 -ChaveRegistro ESC0 -Premium CL_001 -EmailEnvia 0 -ConAzureManual 1
		#				. .\WorkSpace.ps1 -Oferta AS-CS_POWERBI_GRUPO_SEGURANCA -SufixoGrupoSegGeral TESTE1 -Registro SXXXXX -ChaveRegistro esc0 -CodApl PBIP -EmailEnvia 0
		#
		#	Parâmetros de entrada:
		#		Codigo da Aplicacao ( e a partir dela se obtem as chaves dos gestores da aplicação )
		#
		#	CS_POWERBI_APL				
		#	CS_POWERBI_PREMIUM			
		#	CS_POWERBI_GRUPO_SEGURANCA	
		#
		#   Para criação de WorkSpace de Dados, exemplos:
		#	. .\WorkSpace.ps1 -Oferta AS-CS_POWERBI_WS_DADOS -CodApl CINT -WSDadosNivelProt CONFIDENCIAL
		#	. .\WorkSpace.ps1 -Oferta AS-CS_POWERBI_WS_DADOS -CodApl SBRH -WSDadosNivelProt INTERNO -WSDadosSufixoNome EFETIVO
	#endregion Descrição
	#region Parametros
	Param(
		[Parameter(	Mandatory=$false )]						[switch] $CodApl,						# Código da aplicação
		[Parameter(	Mandatory=$false )]						[string] $ValorCodApl,					# = $(throw "Codigo Apl necessario")
#		[Parameter( ParameterSetName="Apl", Mandatory=$true  )]										#
#		[Parameter( ParameterSetName="Premium", Mandatory=$true  )]									#
#		[Parameter( ParameterSetName="GrupoSeg", Mandatory=$true  )]								#
		[Parameter(	Mandatory=$true )]
			[ValidateSet(	"AS-CS_POWERBI_APL",													# Criar a WorkSpace
							"AS-CS_POWERBI_PREMIUM",												# Associar/Desassociar a capacidade Premium
							"AS-CS_POWERBI_GRUPO_SEGURANCA",										# Cria grupo de segurança extra
							"AS-CS_POWERBI_WS_DADOS"												# Criar WS de Dados
							)]								[string] $Oferta,						#
		[Parameter(	Mandatory=$false )]
			[ValidateSet(	"INTERNO",																# Antigo NP-2
							"GERAL",																# Antigo NP-2
							"CONFIDENCIAL"															# Antigo NP-3
							)]								[String] $WSDadosNivelProt,				# Nivel de Protecao para o nome da WS
		[Parameter(	Mandatory=$false )]						[String] $WSDadosSufixoNome,			# Sufixo do Nome da WS
		[Parameter(	Mandatory=$false )]						[String] $Registro,						# Numero do registro do click
		[Parameter(	Mandatory=$false )]						[String] $ChaveRegistro,				# Chave do usuário solicitante do Click
		[Parameter( Mandatory=$false  )]															#		[Parameter( ParameterSetName="Apl", Mandatory=$true  )]
			[ValidateSet("CL_001","CL_002")]				[string] $Premium,						#		[Parameter( ParameterSetName="Premium", Mandatory=$true  )]
		[Parameter( Mandatory=$false  )]															#
#		[Parameter( ParameterSetName="GrupoSeg", Mandatory=$true  )]								#
															[string] $SufixoGrupoSegGeral,			#	$true		$false
		[Parameter(	Mandatory=$false )] 					[bool]	 $WorkspaceCria = $true,		#	$true		$false	/ Default para 
		[Parameter(	Mandatory=$false )] 					[bool]	 $GrupoSegPadraoCria = $true,	#	$true		$false
		[Parameter(	Mandatory=$false )]						[bool] 	 $EmailEnvia = $true,			#	$true		$false
		[Parameter(	Mandatory=$false )]						[bool] 	 $ConAzureManual = $false,		#	$true		$false
		[Parameter(	Mandatory=$false )]						[switch] $ForcaDebug					# = $(throw "Codigo Apl necessario")
	)
	#"WorkspaceCria: $WorkspaceCria"; if ( $WorkspaceCria) { "WorkspaceCria ok"}; exit
	#endregion Parametros
	#region funcoes
	#foreach ($k in $($et.Keys)) { $et.$k=($et.$k+' '*120).substring(0, 120) }
	function mostra { param( $texto, $nivel )
		$complin=77		#130		145
		#echo $texto; return
		$texto -split("`r`n") | ForEach-Object {				# Quebra as várias linhas
			$linha = $_
			if ( $linha.length -ne 0 ) {	# se não for linha em branco		( -not ($linha -match "^`r$") )
				#echo "/$linha/( $( [int][char]$($linha[0]) ) )/( $($linha.length) )"
				#echo ">     $( ($linha+(' '*199)).substring(0, 121) )<"
				if ($linha -match '^\d\.\d' ){
					Write-Output "$('-'*$($complin-1))"
					#ok echo "*$('-'*$($complin-1))*"
					if ( $linha -match "^9\.9") {
						Write-Output $linha.substring(4, $linha.length-4)
						#ok echo "$( "* $linha $(' '*199)".substring(0, $complin) )*"
					} else {
						Write-Output $linha
					}
				} else {
					Write-Output "$(  "$(if ($nivel){ '  '*($nivel-1) })     $linha")"
					#ok echo "$( "*$(if ($nivel){ '  '*($nivel-1) })     $linha $(' '*199)".substring(0, $complin) )*"
				}
				#$linha.length
			} #else { "lb"}
		}
	}
	function fEnviaEmail { param( $pTipoEmail, $MsgFalha )
		#"pTipoEmail: $pTipoEmail/ MsgFalha: $MsgFalha/"
		if ( $EmailEnvia ) {
			mostra $et.90
			#$chavesEmail = $ChaveRegistro + @($app_chaveRT) + $app_gestor
			$chavesEmail = 	[array]$(if($ChaveRegistro){@($ChaveRegistro)}) +
							[array]$(if($app_chaveRT){@($app_chaveRT)}) +
							[array]$(if($app_gestor){@($app_gestor)})
			#"ChaveRegistro: $ChaveRegistro"; "ChavesEmail: $chavesEmail"
			if ( $debugOn ) {	$chavesEmail = $debugChaves	}
			$chavesIncluidas = @()
			#$idWorkSpace = $wsExistenteId		# $IdWsAtiva	$IdWsNova
			$chavesEmail | ForEach-Object {
				$chave = $_
				if ( $chave ) { 
					if ( -not ($chave -in $chavesIncluidas) ) {	# para evitar duplicidade de inclusão de chave de RT que é a mesma de gestor
						"Enviando email para chave: $chave"
						if ( $Oferta -eq "AS-CS_POWERBI_APL" -or $Oferta -eq "AS-CS_POWERBI_PREMIUM" ) {
							./enviaEmail.ps1 -Oferta $Oferta -CodApl $CodApl -Registro $Registro -Premium $Premium -NomeAplic $app_nome -TipoEmail $pTipoEmail -MsgFalha $MsgFalha -ChaveUsuario $chave
							#-idWorkSpace $idWorkSpace 
						} elseif ( $Oferta -eq "AS-CS_POWERBI_GRUPO_SEGURANCA" ) {
							./enviaEmail.ps1 -Oferta $Oferta -CodApl $CodApl -Registro $Registro                   -NomeAplic $app_nome -TipoEmail $pTipoEmail -MsgFalha $MsgFalha -ChaveUsuario $chave -NomeGrupoSeg $NomeGrupoSeg
						}
					}
				}
				$chavesIncluidas += $chave
			}
		}
	}
	function saierro { param( $num ) 
		mostra $sd.$num
		exit $num
	}
	function saiFalha { param( $num )
		if ( $num -is [int] ) {
			$MsgFalha = $sd.$num
		} else {
			$MsgFalha = $num
			$num = 900
		}
		fEnviaEmail -pTipoEmail 'falha' -MsgFalha $MsgFalha
		# envia email
		exit $num
	}
	#endregion funcoes
	#region Constantes 1
	if ( $Oferta -eq 'AS-CS_POWERBI_WS_DADOS' ) {
		$WorkspaceCria = $true		#?
		$GrupoSegPadraoCria = $false
		$EmailEnvia = $false
	}
	if ( $Oferta -eq 'AS-CS_POWERBI_GRUPO_SEGURANCA' ) {
		$WorkspaceCria = $false
		$GrupoSegPadraoCria = $false
		#$GrupoSegGeralCria = $true
		$SufixoGrupoSegGeral = "_$SufixoGrupoSegGeral"
	}
	if ( $GrupoSegPadraoCria ) {
		$tiposGrupo = @('VISUALIZADOR','CONTRIBUIDOR','MEMBRO')
		$SufixoGrupoSegGeral = ""
	} else {
		$tiposGrupo = @('GERAL')
	}
	$debugOn				= $false		#	$true $false
	if ( $ForcaDebug ) { $debugOn = $true }
	if ( $DebugPreference -eq "Inquire" ) { $debugOn = $true }
	$debugChavesGestor		= @('esc0')		#,'cyl4','u4w4','U4WP','U40A','U45P')	#	@('esc0')	@('esc0','cyl4','u4w4','U4WP','U40A','U45P')
	$debugChavesRT			= @('esc0')												#	@('u4w4')	@('esc0')
	$debugChaves			= $debugChavesGestor + $debugChavesRT					# u4w4 - Rafael, U4WP - Joseane, U40A - Arthur, U45P - Renan
	# Adequar uma WS existente: 
	#	Executar este script com: WorkspaceCria=false GrupoSegPadraoCria=true EmailEnvia=false
	#	Excluir (se possível) a equipe Teams. Observar as pessoas da equipe x pessoas dos novos grupos de segurança
	#	Remover da WS as pessoas e grupos antigos ( PBI_DESENOLVEDOR_TIC, pessoas que estão nos grupos de segurança criados )
	$ConsultaCatOracle		= $false		#	$true		$false
	$ConsultaCatBarramento	= $true			#	$true		$false
	$EquipeTeamsCria		= $false		#	$false
	$GrupoMs365Cria			= $false		#	$false
	$idCapacidade000		= '00000000-0000-0000-0000-000000000000'				# Id da Capacidade Compartilhada
	$idCapacidade			= 'D677CDA4-CB81-4874-9AA7-F47B61168B01'				# Id da Capacidade da POC
	$pbiAdminId				= '2832d2e0-a2ea-4ba3-aa58-c941e5cc6bef'				# (Get-AzureADGroup -SearchString  GN_PETROBRAS_POWER-BI_ADMINS).ObjectId
		# '1d45588d-0bae-486d-b0e8-f91b7c1fe3d5'								# GN_PETROBRAS_POWERBI_ADMINS   ( Sem traco do Power-bi só SaPowerbi )
		# '2832d2e0-a2ea-4ba3-aa58-c941e5cc6bef'	
	$pbiDesenvId 			= '6f54a711-3aa4-4561-9fa0-9ba14243c31f'				# GN_PETROBRAS_POWER-BI_DEV_TIC
		#'96e2ccaf-9fee-45a9-9940-d61f6a666196'									# PBI_DESENOLVEDOR_TIC
	$pbiGovernancaId		= 'c37755b6-6d5b-410d-9644-c0b409449127'				# GN_PETROBRAS_POWER-BI_GOV
	$usuarioSessao			= $env:username	#$usuarioSessao; exit
	$et = [ordered]@{	#	Etapas do Script
		#00	= "0.0 Parametros do script"
		10	= "1.0 Consulta catalogo ( <= Codigo_Apl =>  Nome_Apl, Chave_Gestor )"
		12	= "1.2 Conecta com Powerbi"
		13	= "1.3 Verifica se Workspace existe"
		15	= "1.5 Conecta com AzureAD"
		21	= "2.1 Cria grupo de seguranca"
		25	= "2.5 Cria grupo Microsoft 365 e equipe Teams"
		30	= "3.0 Cria a WorkSpace"
		31	= "3.1 Define WorkSpace: Descricao, Cod_App, Nome_ApP, chave Gestores"
		32	= "3.2 Concede acessos de Admin, Membro e Contributor na WorkSpace"
		33	= "3.3 Concede acesso de Visualizador na WorkSpace"
		#34	= "3.4 Define propriedades: Contact List, OneDrive"
		34	= "3.4 Associa Cacacidade Premium a Workspace recem criada"
		40	= "4.0 Associar/Desassociar Capacidade Premium"
		90	= "9.0 Envio de email"
		99	= "Operacao realizada com sucesso"
	}
	#endregion Constantes 1
	#region 05 - Trata parâmetros --------------------------------------------------------------------------------------------------------------
	if ( "$ValorCodApl" -eq "" ) {
		[String] $CodApl = ""
		#$sd
		saiFalha "220 - Falha: C&oacute;digo da aplica&ccedil;&atilde;o deve ser informado"						# Falha
	}
	[string] $CodApl = $ValorCodApl
	$CodApl = $CodApl.ToUpper()
	#Write-Host "Parametros CodApl: $CodApl/ ValorCodApl: $ValorCodApl"
	if ( ($Oferta -eq 'AS-CS_POWERBI_GRUPO_SEGURANCA') -and ($SufixoGrupoSegGeral -eq '_') ) {
		saiFalha "221 - Falha: Sufixo do grupo de segurança deve ser informado"									# Falha
	}
	#exit
	$SufixoGrupoSegGeral = $SufixoGrupoSegGeral.ToUpper()
	$WSDadosSufixoNome = $WSDadosSufixoNome.ToUpper()
	#endregion 05
	#region Constantes 2 ( Após definiição de CodApl )
	Add-Type -AssemblyName System.Web
	$sd = @{	#	Saídas / Erro => erro no job Ansible / Falha não implica em erro no Job
		103 = "103 - Erro: Grupo Microsoft 365 nao existe"														# erro old
		112 = "112 - Erro na criacao da equipe Teams"															# erro old
		113 = "113 - Erro na modificacao da descricao da WorkSpace"												# erro
		114 = "114 - Erro consultando email do usuario"															# erro
		115 = "115 - Erro adicionando usuario como membro da WorkSpace"											# erro
		116 = "116 - Erro adicionando usuario a equipe do Teams"												# erro
		117 = "117 - Erro adicionando grupo admin na WorkSpace"													# erro
		118 = "118 - Erro adicionando grupo de Viewer na WorkSpace"												# erro
		119 = "119 - Erro na criacao da grupo MS365"															# erro old
		121 = "121 - Erro: Connect Powerbi"																		# erro
		122 = "122 - Erro: Connect AzureAD"																		# erro
		123 = "123 - Erro: Connect Teams"																		# erro
		127 = "127 - Erro obtendo as credenciais para BD"														# erro old
		128 = "128 - Erro obtendo as credenciais para PowerBI"													# erro
		131 = "131 - Erro adicionando grupo Member na WorkSpace"												# erro
		132 = "132 - Erro adicionando grupo Contributor na WorkSpace"											# erro
		143 = "143 - Erro na associacao da Workspace a Premium"													# erro
		145 = "145 - Erro na desassociacao da Workspace a Premium"												# erro
		150 = "150 - Erro na criacao da Workspace"																# erro
		151 = "151 - Erro na verificacao da Workspace"															# erro
		160 = "160 - Erro na criacao do grupo de segurança"														# erro
		165 = "165 - Erro adicionando usuario no grupo de seguranca"											# erro
		167 = "167 - Erro adicionando usuario RT no grupo de seguranca"											# erro
		170 = "170 - Erro adicionando grupo de seguranca com VISUALIZADOR da WorkSpace"							# erro
		171 = "171 - Erro adicionando grupo de seguranca com CONTRIBUIDOR da WorkSpace"							# erro
		172 = "172 - Erro adicionando grupo de seguranca com MEMBRO da WorkSpace"								# erro
		211 = "211 - Falha: Existe Workspace ($CodApl) com o nome solicitado"									# Falha / poderia ser considerado alteração
		212 = "212 - Falha: Aplica&ccedil;&atilde;o $CodApl inexistente ou nome vazio"							# Falha
		213 = "213 - Falha: Aplica&ccedil;&atilde;o $CodApl sem paradigma 'Power BI'"							# Falha
		214 = [System.Web.HttpUtility]::HtmlEncode("214 - Falha: Aplicação $CodApl sem cadastro de responsáveis (gestores).
			A aplicação não possui os responsáveis pela solução cadastrados no Catálogo de Aplicações.
			Por favor, entre em contato com a área de Parceria de Negócio para ajuste do cadastro no Catálogo e em seguida abra um novo chamado.
		")		# Falha
		217 = "217 - Falha: Workspace ($CodApl) inexistente"													# Falha / poderia ser considerado criação
	}
	#endregion Constantes 2
#endregion 0
#region 10 - Consulta Catalogo -----------------------------------------------------------------------------------------------------------------
mostra $et.10
<# sqlplus /nolog
# Desenv: ( DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=npaa5968.petrobras.biz)(PORT=1521)) (CONNECT_DATA=(SID=oradsv11)) )
conn sapowerbi/"senha"@'( DESCRIP...'
# Prod:   ( DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=npaa5787.petrobras.biz)(PORT=1521)) (CONNECT_DATA=(SERVICE_NAME=oraprd11.PETROBRAS.COM.BR)(INSTANCE_NAME=rcprd11a)) )
conn sapowerbi/"senha"@'( DESCRIP...'
#conn cmxl/"senha"@'(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = npaa2654.petrobras.biz)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = vcxlh)(SERVER = DEDICATED)))'
SET PAGESIZE 999
SET LINESIZE 140
SELECT * FROM all_users ORDER BY username;

Desc usr.tabela
select * from usr.tabela
  -- u.USUA_TX_EMAIL Email_Gestor
  -- left join SAGM.USUARIO u on u.USUA_CD_CHAVE = r.RESP_CD_CHAVE_RESPONSAVEL
  a.APLI_SG_APLICACAO Sigla_Aplicacao,
  a.USUA_CD_RESPONSAVEL Chave_Resp_PN,
  a.APLI_DS_APLICACAO Descricao_Aplicacao,
  a.PADE_CD_PARADIGMA_DESENVO Chave_Paradigma,
  p.PADE_DS_PARADIGMA_DESENVO Paradigma_Desenv
#>
# Select Aplicação e Gestores do catálogo
if ( $ConsultaCatOracle ) {
	$queryAplGestor = "select
	  a.APLI_CD_IDENTIFICADOR Codigo
	  ,a.APLI_NM_APLICACAO Nome_Aplicacao
	  ,r.RESP_CD_CHAVE_RESPONSAVEL Chave_Gestor
	  ,a.USUA_CD_SUBSTITUTO Chave_RT
	from	    SAGM.APLICACAO a 
	  left join SAGM.RESPONSAVEL r on a.APLI_CD_APLICACAO = r.APLI_CD_APLICACAO
	  left join SAGM.PARADIGMA_DESENVOLVIMENTO p on a.PADE_CD_PARADIGMA_DESENVO = p.PADE_CD_PARADIGMA_DESENVO 
	where 
		   a.APLI_CD_IDENTIFICADOR = '$CodApl'
	  and  ( p.PADE_DS_PARADIGMA_DESENVO = 'Power BI' ) -- or p.PADE_DS_PARADIGMA_DESENVO = 'Spotfire'  )"
	if ( $debugOn ) {	$t=$queryAplGestor.split( "`r`n" );	$queryAplGestor=$t[0..($t.Count-2)]		}	# elimina a ultima linha

	# Credenciais BD:
	#$usuarioSessao = $env:username
	$username="sapowerbi"
	$FilePass = "D:\Apps\cred\$($usuarioSessao)_user_bd_$username.txt" # Salvar a senha: (Get-Credential -Credential $username ).Password | ConvertFrom-SecureString | Out-File $FilePass
	try {
		$MeuCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, ( Get-Content $FilePass | ConvertTo-SecureString )
	} catch {
		Mostra ( $Error[0] | out-string ) 2
		saierro 127
	}
	$password = $MeuCred.GetNetworkCredential().Password
	#scb $password

	Add-Type -Path "D:\Oracle\oracle.manageddataaccess.core.2.19.90\lib\netstandard2.0\Oracle.ManagedDataAccess.dll"
	$datasource = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=npaa5787.petrobras.biz)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=oraprd11.PETROBRAS.COM.BR)(INSTANCE_NAME=rcprd11a)))"
	$connectionString = 'User Id='+$username + ';Password='+$password + ';Data Source='+$datasource
	$connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)
	$connection.open()
	$command=$connection.CreateCommand()
	$command.CommandText=$queryAplGestor
	$reader=$command.ExecuteReader()

	$app_gestor = @()
	mostra "Codigo    : $CodApl"
	while ($reader.Read()) {
		if ( $app_gestor.length -eq 0 ) {
			mostra "Gestores  :"
		}
		$app_nome = $reader.GetString(1)
		$app_gestor += $reader.GetString(2).Trim()
		$app_chaveRT = $reader.GetString(3).Trim()
		mostra "$($reader.GetString(2))" 2		# $($reader.GetString(0)) / $($reader.GetString(1)) / 
	}
	mostra "RT        : $app_chaveRT"
	$connection.Close()
}
if ( $ConsultaCatBarramento ) {
	#region - Crediciais para SOAP ----------------------------------------------------------------------------------------------------------
	$usuarioSessao = $env:username
	$usernameConnect="SAPOWERBI"			# Deve ser maiúsculo
	$FilePass = "D:\Apps\cred\$($usuarioSessao)_user_$usernameConnect.txt" # Salvar a senha: (Get-Credential -Credential $usernameConnect ).Password | ConvertFrom-SecureString | Out-File $FilePass
	try {
		$MeuCredSoap = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usernameConnect, ( Get-Content $FilePass | ConvertTo-SecureString )
		#$cred 	 = New-Object -TypeName System.Management.Automation.PSCredential("esc0",$passSec)
	} catch {
		$Error[0]
	}
	$URI = "https://bs.petrobras.com.br/CATAPL/Services/ProxyServices/Aplicacao_1"
	$xmlRec = @"
	<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:ns='http://services.petrobras.biz/ti/aplicacao/1'>
		<soapenv:Header/>
		<soapenv:Body>
			<ns:buscaAplicacaoPorCodigo>	<ns:codigo>__CODIGO_APP__</ns:codigo>	</ns:buscaAplicacaoPorCodigo>
		</soapenv:Body>
	</soapenv:Envelope>
"@
	#endregion
	$xmlUsar = $xmlRec -replace "__CODIGO_APP__", $CodApl			#	CODAPL:		QACT	0G0Y	AARE	G2OZ	IG1J
	try {
		Clear-Variable result, xml -ErrorAction Ignore
	} catch {}
	$result = Invoke-WebRequest $URI -Method Post -ContentType 'application/xml' -Headers (@{SOAPAction='Read'}) -Body $xmlUsar -Credential $MeuCredSoap
	$xml = [xml]$result.Content
	$app_nome = $xml.Envelope.Body.buscaAplicacaoPorCodigoResponse.AplicacaoResponse.serviceData.aplicacao.geral.nome
	$app_gestor = $xml.Envelope.Body.buscaAplicacaoPorCodigoResponse.AplicacaoResponse.serviceData.aplicacao.Responsaveis.ResponsavelCliente.chave
	$app_chaveRT = $xml.Envelope.Body.BuscaAplicacaoPorCodigoResponse.AplicacaoResponse.ServiceData.Aplicacao.geral.LiderProduto.chave
	$paradigma = $xml.Envelope.Body.buscaAplicacaoPorCodigoResponse.AplicacaoResponse.serviceData.aplicacao.geral.paradigma
	if ( $app_nome ) { $app_nome = $app_nome.Trim() }
	#$app_gestor = $app_gestor
	if ( $app_chaveRT ) { $app_chaveRT = $app_chaveRT.Trim() }
	mostra "Codigo    : $CodApl"
	mostra "Gestores  : $app_gestor"
	mostra "RT        : $app_chaveRT"
}
if ( ($Oferta -eq 'AS-CS_POWERBI_WS_DADOS') ) {
	if ( $WSDadosSufixoNome ) {
		$WSDadosSufixoNome="_$WSDadosSufixoNome"
	}
	$NomeWS 		= "PBI_DADOS_$($CodApl)_$WSDadosNivelProt$WSDadosSufixoNome"
} else {
	if ( ! $app_nome ) {	# Cod App nao encontrada
		saiFalha 212
	}
	if ( $paradigma -ne 'Power BI' -and -not $debugOn ) {
		saiFalha 213
	}
	if ( $app_gestor.Count -eq 0 ) {
		saiFalha 214
	}
	$NomeWS 		= "PBI_$CodApl - $app_nome"
}
mostra "Workspace : $NomeWS"
#mostra "fim BD"; exit
#endregion
#region 12 - Credenciais -----------------------------------------------------------------------------------------------------------------------
mostra $et.12
# Bug: Conexao com PowerBI para executar Get-PowerBIWorkspace pela primeira vez e evitar o erro "Attempted to access an element as a type incompatible with the array"
# $usuarioSessao = $env:username
$usernamePowerBI1="SAPOWERBI"	# "brauner@petrobras.com.br"
#$usernamePowerBI2="SDPOWERBI@petrobras.com.br"	# "asesc0@petrobras.com.br"	# "brauner@petrobras.com.br"
#$username2 = "SDPOWERBI" # Igual a usernamePowerBI2
#$usernamePowerBI3="asesc0@petrobras.com.br"		# "asesc0@petrobras.com.br"	# "brauner@petrobras.com.br"
$FilePass1 = "D:\Apps\cred\$($usuarioSessao)_user_azuread_$usernamePowerBI1.txt" # Salvar a senha: (Get-Credential -Credential $usernamePowerBI1 ).Password | ConvertFrom-SecureString | Out-File $FilePass1
#$FilePass2 = "D:\Apps\cred\$($usuarioSessao)_user_azuread_$usernamePowerBI2.txt" # Salvar a senha: (Get-Credential -Credential $usernamePowerBI2 ).Password | ConvertFrom-SecureString | Out-File $FilePass2
#$FilePass3 = "D:\Apps\cred\$($usuarioSessao)_user_azuread_$usernamePowerBI3.txt" # Salvar a senha: (Get-Credential -Credential $usernamePowerBI3 ).Password | ConvertFrom-SecureString | Out-File $FilePass3
try {
	$MeuCred1 = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$usernamePowerBI1@petrobras.com.br", ( Get-Content $FilePass1 | ConvertTo-SecureString )
} catch {
	Mostra ( $Error[0] | out-string ) 2
	saierro 128
}
try {
	#$MeuCred2 = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usernamePowerBI2, ( Get-Content $FilePass2 | ConvertTo-SecureString )
} catch {
	Mostra ( $Error[0] | out-string ) 2
	saierro 128
}
<# try {
	$MeuCred3 = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usernamePowerBI3, ( Get-Content $FilePass3 | ConvertTo-SecureString )
} catch {
	Mostra ( $Error[0] | out-string ) 2
	saierro 128
}#>
#endregion
#region 12 - Conecta com Powerbi ---------------------------------------------------------------------------------------------------------------
mostra "Inicia Conexao PowerBI"
try {
	Get-PowerBIAccessToken | Out-Null # $gpbiat = 
} catch {
	try {
		$conPowerbi = Connect-PowerBIServiceAccount -Credential $MeuCred1 | Format-List | Out-String
		Mostra $conPowerbi 2
	} catch {
		Mostra ( $Error[0] | out-string ) 2
		saierro 121
	}
}
#endregion
if ( ($Oferta -eq "AS-CS_POWERBI_APL") -or ($Oferta -eq "AS-CS_POWERBI_GRUPO_SEGURANCA") -or ($Oferta -eq "AS-CS_POWERBI_WS_DADOS") ) {
	# Se a oferta for Criar WorkSpace ou criar GrupoSeg
	#region 13 - Verifica se Workspace existe --------------------------------------------------------------------------------------------------
	mostra $et.13
	#$ws = Get-PowerBIWorkspace -Name "$NomeWS" -Scope Organization
	#"P1  $($ws.name) / $($ws.State)"
	mostra "Nome WS: $NomeWS"
	try {
		$wsVerifica = Get-PowerBIWorkspace -Scope Individual -Name "$NomeWS"	# Organization ### alterado em 30/09/21
	} catch {
		Mostra ( $Error[0] | out-string ) 2
		saierro 151
	}
	$wsAtiva = $wsVerifica # | Where-Object { $_.State -eq 'Active' }  ### alterado em 30/09/21 - Juntamente com Scope Individual
		#$ws0 = Get-PowerBIWorkspace -Scope Organization -Name "$NomeWS" | ? { $_.State -eq "Active" }; $ws | select id, type
		# $ws diferente de Null sigfinica que ja existe Workspace com o nome informado (ja existe Active ou Deleted)
		# $ws com state Active significa que a WS está ativa (não deletada)
	if ( $WorkspaceCria -and ( ( $null -ne $wsAtiva ) ) ) { # -and ( $wsAtiva[$wsAtiva.Count-1].State -eq 'Active' ) ) ) {
		saiFalha 211
	}
	if ( ($Oferta -eq "AS-CS_POWERBI_GRUPO_SEGURANCA") -and ( ($null -eq $wsAtiva) ) ) { #  -or ($wsAtiva[$wsAtiva.Count-1].State -ne 'Active') ) ) {
		saiFalha 217		# WS nao existe
	}
	$IdWsAtiva = $wsAtiva.id
	#endregion
	#region 15 - Conecta com AzureAD -----------------------------------------------------------------------------------------------------------
	mostra $et.15
	Get-PowerBIWorkspace -Name "$NomeWS" -Scope Individual | Out-Null		# Organization		# $wsbug = # Para evitar o bug descrito acima
	# Credenciais temporárias do Azure / Necessario fazer o Connect-AzureAD ( para testes com a chave pessoal ). Depois sera utilizada o App com Client_id e Secrect_id.
	mostra "Inicia Conexao AzureAD"
	try {
		if ( $ConAzureManual ) {
			$conAzure = Connect-AzureAD | Format-List | Out-String
		} else {
			$conAzure = Connect-AzureAD -Credential $MeuCred1 | Format-List | Out-String		# Mudar MeuCred1 / MeuCred2
		}
		Mostra $conAzure 2
	} catch {
		Mostra ( $Error[0] | out-string ) 2
		saierro 122
	}
	#endregion
	#region 21 - Cria grupo de segurança -------------------------------------------------------------------------------------------------------
	if ( $GrupoSegPadraoCria -or ( $Oferta -eq 'AS-CS_POWERBI_GRUPO_SEGURANCA' ) ) {
		function CriaGrupoUsandoPool {
			param ( $NovoNomeGrupoSeg )
			$ogr = Get-AzureADGroup -SearchString 'GN_PBI_POOL' #-all $true
			if ( $ogr ) {
				$NovoDescGrupoSeg = $ogr[0].Description -replace ".* /Creator:", "$NovoNomeGrupoSeg /Creator:"
				Set-AzureADGroup  -ObjectId $ogr[0].ObjectId -DisplayName $NovoNomeGrupoSeg -Description $NovoDescGrupoSeg
				$oNovoGr = Get-AzureADGroup -Filter "DisplayName eq '$NovoNomeGrupoSeg'"
				return $oNovoGr
			}
		}
		mostra $et.21
		$ogrupoSeg = @{}
		$objUserSaPowerBI = Get-AzureADUser -filter "MailNickName eq 'sapowerbi'" 
		#$objUser2 = Get-AzureADUser -filter "MailNickName eq '$username2'" 
		foreach ( $tipoGrupo in $tiposGrupo ) {	#	'VISUALIZADOR','CONTRIBUIDOR','MEMBRO' / 'GERAL'
			$NomeGrupoSeg = "GN_PBI_$($CodApl)_$tipoGrupo$SufixoGrupoSegGeral"
			mostra "Grupo Seguranca $tipoGrupo :" # $NomeGrupoSeg"
			try {
				$grupoexiste = $false
				$oGrupoSeg[$tipoGrupo] = Get-AzureADGroup -Filter "DisplayName eq '$NomeGrupoSeg'"
				if ( $oGrupoSeg[$tipoGrupo] ) {
					$grupoexiste = $true
					mostra ( "Grupo existente" ) 2
				} else {
					#$oGrupoSeg[$tipoGrupo] = New-AzureADGroup -DisplayName $NomeGrupoSeg -Description $NomeGrupoSeg -MailEnabled $false -SecurityEnabled $true -MailNickName NotSet
					$oGrupoSeg[$tipoGrupo] = CriaGrupoUsandoPool("$NomeGrupoSeg")
				}
				mostra ( $oGrupoSeg[$tipoGrupo] | Format-List  ObjectId, DisplayName | out-string ) 2		# Description
				#mostra "Grupo de seguranca foi criado: $NomeGrupoSeg"
				try {		#### INCLUI as chaves dos Gestores e RT como Owner e Member de todos os grupos
					$chavesProp = 	[array]$(if($app_chaveRT){@($app_chaveRT)}) +
									[array]$(if($app_gestor){@($app_gestor)})
					if ( $debugOn ) {	$chavesProp = $debugChavesGestor	}
					if ( ! $grupoexiste ) {	# Grupo foi criado agora
						try {
							Add-AzureADGroupOwner -ObjectId $oGrupoSeg[$tipoGrupo].ObjectId -RefObjectId $objUserSaPowerBI.ObjectId
							#Remove-AzureADGroupOwner -ObjectId $oGrupoSeg[$tipoGrupo].ObjectId -OwnerID $objUser2.ObjectId
						} catch {}
						$chavesIncluidas = @()
						$chavesProp | ForEach-Object {
							if ( -not ($_ -in $chavesIncluidas) ) {	# para evitar duplicidade de inclusão de chave de RT que é a mesma de gestor
								$chaveg = $_		# .Trim()
								$objUser = Get-AzureADUser -filter "MailNickName eq '$chaveg'" #Get-AzureADUser -SearchString "$chaveg"
								mostra "chave: $chaveg" 2
								if ( $objUser ) {
									Add-AzureADGroupOwner  -ObjectId $oGrupoSeg[$tipoGrupo].ObjectId -RefObjectId $objUser.ObjectId
									Add-AzureADGroupMember -ObjectId $oGrupoSeg[$tipoGrupo].ObjectId -RefObjectId $objUser.ObjectId
								}
							}
							$chavesIncluidas += $_
						}
					}
				} catch {
					Mostra ( $Error[0] | out-string ) 2
					saierro 165
				}
			} catch {
				Mostra ( $Error[0] | out-string ) 2
				saierro 160
			}
			<# Em 03/02/2021 foi decidido incluir o RT também como proprietário
			if ( $tipoGrupo -eq 'CONTRIBUIDOR' -or $tipoGrupo -eq 'GERAL' ) {		#### INCLUI as chaves dos RT como Member dos grupos CONTRIBUIDOR e GERAL
				$chaveMember = $app_chaveRT
				if ( $debugOn ) {	$chaveMember = $debugChavesRT	}
				if ( $chaveMember -notin $chavesOwner ) { 
					$objUserC = Get-AzureADUser -SearchString "$chaveMember"
					try {
						Add-AzureADGroupMember -ObjectId $oGrupoSeg[$tipoGrupo].ObjectId -RefObjectId $objUserC.ObjectId
					} catch {
						Mostra ( $Error[0] | out-string ) 2
						saierro 167
					}
				}
			} #>
		}
		#exit
		<# comandos AzureADGroup
		New-AzureADGroup -DisplayName PBI_TRPB_VISUALIZADOR -Description 'Teste1' -MailEnabled $false -SecurityEnabled $true -MailNickName NotSet
		Set-AzureADGroup -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_VISUALIZADOR'" ).ObjectId -Description 'GN_PBI_AARE_VISUALIZADOR'

		Add-AzureADGroupMember -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_VISUALIZADOR'" ).ObjectId -RefObjectId (Get-AzureADUser -ObjectId "rafael.espiritosanto@petrobras.com.br")
		Add-AzureADGroupMember -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_VISUALIZADOR'" ).ObjectId -RefObjectId (Get-AzureADUser -ObjectId "joseanefreire@petrobras.com.br").ObjectId

		Add-AzureADGroupOwner -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_VISUALIZADOR'" ).ObjectId -RefObjectId (Get-AzureADUser -ObjectId "rafael.espiritosanto@petrobras.com.br").ObjectId

		Get-AzureADGroupMember -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_VISUALIZADOR'" ).ObjectId
		Get-AzureADGroupOwner  -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_VISUALIZADOR'" ).ObjectId

		Get-AzureADGroupMember -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_CONTRIBUIDOR'" ).ObjectId
		Get-AzureADGroupOwner  -ObjectId ( Get-AzureADGroup -Filter "DisplayName eq 'GN_PBI_AARE_CONTRIBUIDOR'" ).ObjectId

		#>
	}
	#endregion
	#region 25 - Cria grupo Microsoft 365 e equipe Teams ---------------------------------------------------------------------------------------
	if ( $GrupoMs365Cria ) {
		mostra $et.25

		$NomeGrupoMs365 = "MS365_$CodApl - $app_nome"
		mostra "Grupo Ms365: $NomeGrupoMs365"

		$MailNickName = ($NomeGrupoMs365 -Split ' |-')[0]
		try {						# Cria grupo Ms 365
			mostra "Cria grupo MS 365"
			$oGrupoMs365 = New-AzureADMSGroup -DisplayName $NomeGrupoMs365 -MailEnabled $false -MailNickName $MailNickName -SecurityEnabled $true -GroupTypes "Unified" -Visibility Private
			mostra ( $oGrupoMs365 | Format-List  Id, DisplayName, MailNickname, Mail | out-string ) 2
			mostra "Grupo MS 365 criado: $NomeGrupoMs365"
		} catch {
			#Mostra "Erro na criacao da grupo MS365"
			Mostra ( $Error[0] | out-string ) 2
			saierro 119
		}
		<#
		$oGrupoMS365 = Get-AzureADGroup -SearchString $NomeGrupoMs365
		if ( ! $oGrupoMS365 ) { 
			saierro 103
		}#>
	}
	if ( $EquipeTeamsCria ) {			# Cria equipe Teams
		mostra "Inicia Conexao Teams"
		try {
			$conTeams = Connect-MicrosoftTeams -Credential $MeuCred | Format-List  | Out-String
			Mostra $conTeams 2
		} catch {
			#Mostra "Erro Connect-MicrosoftTeams:"
			Mostra ( $Error[0] | out-string ) 2
			saierro 123
		}
		try {
			$oGrupoTeams = New-Team -GroupId $oGrupoMs365.Id #-DisplayName $NomeGrupoMs365 -MailNickName $MailNickName #"PBI_Teste9 - Governança Spotfire"
		} catch {
			#Mostra "Erro na criacao da equipe:"
			Mostra ( $Error[0] | out-string ) 2
			saierro 112
		}
		mostra "Equipe criada: $NomeGrupoMs365"
	}
	#endregion
	if ( $WorkspaceCria ) {
		#region 30 - Cria a WorkSpace ----------------------------------------------------------------------------------------------------------
		mostra $et.30
		#"Vai criar WS / ws -ne $null : $($ws -ne $null) / $ ws.Count: $($ws.Count) / ( ws[ws.Count-1].State -eq 'Active' ) $(( $ws[$ws.Count-1].State -eq 'Active' ))"
		try {
			$wsNova = New-PowerBIWorkspace -Name $NomeWS -ErrorAction Stop # Cria a WorkSpace
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 150
		}
		#exit
		mostra "WorkSpace criada / id: $($wsNova.Id)"

		$wsNovaUltima = $wsNova[$wsNova.Count-1]
		$IdWsNova = $wsNovaUltima.Id
		$IdWsAtiva = $IdWsNova
		mostra "Ultima Workspace / id: $($wsNova.Id)"
		#endregion
		#region 31 - Define WorkSpace: Descricao, Cod_App, Nome_ApP, chave Gestores ------------------------------------------------------------
		mostra $et.31
		$wsDescricao = "Nome app: $app_nome 
		Gestores : $app_gestor
		Resp.Tec.: $app_chaveRT
		"
		mostra $wsDescricao
		$wsNovaUltima.Description = $wsDescricao
		try {	# Modifica descrição
			Set-PowerBIWorkspace -Workspace $wsNovaUltima -Scope Organization
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 113
		}
		#endregion
	}
	#region 32 - Concede acessos de Admin, Membro e Contributor na WorkSpace -------------------------------------------------------------------
	mostra  $et.32
	<# antigo / inclui chaves como Membro da workspace
	$chavesIncluidas = @()
	$chaves1 = @($app_chaveRT) + $app_gestor
	if ( $debugOn ) {	$chaves1 = $debugChaves	}
	$chaves1 | % {
		$chave = $_.Trim()
		if ( -not ($chave -in $chavesIncluidas) ) {	# para evitar duplicidade de inclusão de chave de RT que é a mesma de gestor
			#mostra "$chave" 2
			try {
				$emailUser = (Get-AzureADUser -SearchString "$chave").UserPrincipalName
			} catch {
				Mostra ( $Error[0] | out-string ) 2
				saierro 114
			}		
			try {
				Add-PowerBIWorkspaceUser -AccessRight Member -Id $IdWsNova -PrincipalType User -Identifier $emailUser
				mostra "Membro da WorkSpace adicionado: $chave / $emailUser" 2
			} catch {
				Mostra ( $Error[0] | out-string ) 2
				saierro 115
			}
			if ( $EquipeTeamsCria ) {
				try {
					Add-TeamUser -GroupId $oGrupoTeams.GroupId -User $emailUser -Role Owner
					mostra "Membro da Equipe adicionado   : $chave / $emailUser" 2
				} catch {
					Mostra ( $Error[0] | out-string ) 2
					saierro 116
				}		
				# -Scope Organization 
			}
		}
		$chavesIncluidas += $chave
	}	#>
	if ( $WorkspaceCria ) {
		try {
			Add-PowerBIWorkspaceUser -AccessRight Admin -Id $IdWsAtiva -PrincipalType Group -Identifier $pbiAdminId -ErrorAction Stop
			mostra "acesso Admin concedido na WorkSpace para GN_PETROBRAS_POWER-BI_ADMINS"
		} catch {
			#Mostra "Erro adiconando grupo admin na WorkSpace"
			Mostra ( $Error[0] | out-string ) 2
			saierro 117
		}
	}
	if ( $GrupoSegPadraoCria ) {
		$ogrupoSegContrib = $oGrupoSeg['CONTRIBUIDOR']
		try {
			Add-PowerBIWorkspaceUser -AccessRight Contributor -Id $IdWsAtiva -PrincipalType Group -Identifier $ogrupoSegContrib.ObjectId -ErrorAction Stop
			mostra "acesso Contributor concedido na WorkSpace para $($ogrupoSegContrib.DisplayName)"
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 171
		}
		$ogrupoSegMembro = $oGrupoSeg['MEMBRO']
		try {
			Add-PowerBIWorkspaceUser -AccessRight Member -Id $IdWsAtiva -PrincipalType Group -Identifier $ogrupoSegMembro.ObjectId -ErrorAction Stop
			mostra "acesso Member concedido na WorkSpace para $($ogrupoSegMembro.DisplayName)"
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 172
		}
		try {
			Add-PowerBIWorkspaceUser -AccessRight Member -Id $IdWsAtiva -PrincipalType Group -Identifier $pbiGovernancaId -ErrorAction Stop
			mostra "acesso Member concedido na WorkSpace para GN_PETROBRAS_POWER-BI_GOV"
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 131
		}
		try {
			Add-PowerBIWorkspaceUser -AccessRight Contributor -Id $IdWsAtiva -PrincipalType Group -Identifier $pbiDesenvId -ErrorAction Stop
			mostra "acesso Contributor concedido na WorkSpace para GN_PETROBRAS_POWER-BI_DEV_TIC"
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 132
		}
	}
	if ( $Oferta -eq 'AS-CS_POWERBI_WS_DADOS' ) {
		try {
			Add-PowerBIWorkspaceUser -AccessRight Member -Id $IdWsAtiva -PrincipalType Group -Identifier $pbiGovernancaId -ErrorAction Stop
			mostra "acesso Member concedido na WorkSpace para GN_PETROBRAS_POWER-BI_GOV"
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 131
		}
		try {
			Add-PowerBIWorkspaceUser -AccessRight Member -Id $IdWsAtiva -PrincipalType Group -Identifier $pbiDesenvId -ErrorAction Stop
			mostra "acesso Member concedido na WorkSpace para GN_PETROBRAS_POWER-BI_DEV_TIC"
		} catch {
			Mostra ( $Error[0] | out-string ) 2
			saierro 132
		}
	}
	<# exemplos Add-PowerBIWorkspaceUser:
	Add-PowerBIWorkspaceUser  -Scope Organization -AccessRight Member -ID $ws.id -UserPrincipalName sizo.dn@petrobras.com.br
	Add-PowerBIWorkspaceUser -AccessRight Member -Id $ws.id -PrincipalType Group -Scope Individual 
	Add-PowerBIWorkspaceUser -AccessRight Member -Id $ws.id -PrincipalType Group -Identifier 3d45b5c3-4c2f-435d-b1f0-a21f1e52eebb
	Add-PowerBIWorkspaceUser -AccessRight Member -Id $ws.id -PrincipalType User -Identifier cassius.serra@petrobras.com.br
	#>
	function dbg {
		if ( $EquipeTeamsCria ) {
			mostra "Equipe Teams:"
			mostra ( $oGrupoTeams | Select-Object GroupID, DisplayName, MailNickName | out-string ) 2
		}
		mostra "Grupo MS365:"
		$Gr365T = Get-AzureADGroup -SearchString $NomeWS.substring(0,8)
		mostra ( $Gr365T |  Select-Object ObjectId, DisplayName, MailNickName | out-string ) 2
		$wsT = Get-PowerBIWorkspace -Scope Individual -Name "$NomeWS" | Where-Object { $_.State -eq "Active" }	# Organization
		mostra "WorkSpace:"
		mostra ( $wsT | Select-Object id, name, type | out-string ) 2
	}
	#endregion
	#region 33 - Concede acesso de Visualizador na WorkSpace -----------------------------------------------------------------------------------
	if ( $GrupoSegPadraoCria ) {
		mostra $et.33
		#Add-PowerBIWorkspaceUser -AccessRight Viewer -Id $IdWsAtiva -PrincipalType Group -Identifier $oGrupoMS365.ObjectId
		$ogrupoSegVisualizador = $oGrupoSeg['VISUALIZADOR']
		try { #IdWsAtiva		IdWsNova				?workspaceV2=true
			Invoke-PowerBIRestMethod -Url "groups/$IdWsAtiva/users?workspaceV2=true" -Method POST -Body ( @{identifier="$($ogrupoSegVisualizador.ObjectId)"; groupUserAccessRight='Viewer'; principalType='Group'} | ConvertTo-Json -Depth 2 -Compress )
			mostra "Viewer da WorkSpace para $($oGrupoSeg['VISUALIZADOR'].DisplayName)" 					#	$($oGrupoMS365.DisplayName)"
		} catch {
			#Mostra "Erro adiconando grupo de Viewer na WorkSpace"
			Mostra ( $Error[0] | out-string ) 2
			saierro 118
		}
	}
	#endregion
	#region 34 - Associa Cacacidade Premium a Workspace recem criada ---------------------------------------------------------------------------
	if ( ($Premium -eq "CL_001") -or ($Oferta -eq 'AS-CS_POWERBI_WS_DADOS') ) {
		mostra $et.34
		try {
			Set-PowerBIWorkspace -id $IdWsAtiva -Scope Organization -CapacityId $idCapacidade
		} catch {
			#Mostra "Erro na associacao da capacidade"
			Mostra ( $Error[0] | out-string ) 2
			saierro 143
		}
	}
	#endregion
	#region 9 - email Workspace criada ou GRUPOSeg Criado
	if ( $EmailEnvia ) {
		mostra $et.90
		#$chavesEmail = @($app_chaveRT) + $app_gestor
		$chavesEmail = 	[array]$(if($ChaveRegistro){@($ChaveRegistro)}) +
						[array]$(if($app_chaveRT){@($app_chaveRT)}) +
						[array]$(if($app_gestor){@($app_gestor)})
		#"ChaveRegistro: $ChaveRegistro"; "ChavesEmail: $chavesEmail"
		if ( $debugOn ) {	$chavesEmail = $debugChaves	}
		$idWorkSpace = $IdWsAtiva			#	$wsExistenteId
		$chavesIncluidas = @()
		$chavesEmail | ForEach-Object {
			$chave = $_
			if ( $chave ) { 
				if ( -not ($chave -in $chavesIncluidas) ) {	# para evitar duplicidade de inclusão de chave de RT que é a mesma de gestor
					Mostra "Enviando email para chave: $chave" 2
					if ( $Oferta -eq "AS-CS_POWERBI_APL" ) {
						./enviaEmail.ps1 -Oferta $Oferta -CodApl $CodApl -Registro $Registro -Premium $Premium -NomeAplic $app_nome -idWorkSpace $idWorkSpace -ChaveUsuario $chave -TipoEmail 'ok'
					} elseif ( $Oferta -eq "AS-CS_POWERBI_GRUPO_SEGURANCA" ) {
						./enviaEmail.ps1 -Oferta $Oferta -CodApl $CodApl -Registro $Registro                   -NomeAplic $app_nome -idWorkSpace $idWorkSpace -ChaveUsuario $chave -TipoEmail 'ok' -NomeGrupoSeg $NomeGrupoSeg
					}
				}
			}
			$chavesIncluidas += $chave
		}
	}
	#endregion
	#region fim
	mostra "$($et.99) - WS: https://app.powerbi.com/groups/$IdWsAtiva"
	#endregion
} elseif ( $Oferta -eq "AS-CS_POWERBI_PREMIUM" )  {	# Oferta de Associar/Desassociar Capacidade Premium
	#region 40 - Associar/Desassociar Capacidade Premium ---------------------------------------------------------------------------------------
	mostra $et.40
	$wsExistente = Get-PowerBIWorkspace -Scope Individual -Name "$NomeWS"		# Organization
	$wsExistenteAtiva = $wsExistente #| Where-Object { $_.State -eq 'Active' }				# alterado em 04/10
	if ( ( $null -eq $wsExistenteAtiva ) ) { # -or ( $wsExistenteAtiva[$wsExistenteAtiva.Count-1].State -ne 'Active' ) ) {		# alterado em 04/10
		saiFalha 217		# WS nao existe
	}
	$wsExistenteUltima = $wsExistenteAtiva[$wsExistenteAtiva.Count-1]
	$wsExistenteId = $wsExistenteUltima.Id
	if ( $Premium -eq "CL_001"  ) {	# "Sim","Associar"
		try {
			Set-PowerBIWorkspace -id $wsExistenteId -Scope Organization -CapacityId $idCapacidade
		} catch {
			#Mostra "Erro na associacao da capacidade"
			Mostra ( $Error[0] | out-string ) 2
			saierro 143
		}
		
	} else {	# Desassociar
		try {
			Set-PowerBIWorkspace -id $wsExistenteId -Scope Organization -CapacityId $idCapacidade000
		} catch {
			#Mostra "Erro na desassociacao da capacidade"
			Mostra ( $Error[0] | out-string ) 2
			saierro 145
		}
	}

	#endregion
	#region 9 - email Associar/Desassociar -----------------------------------------------------------------------------------------------------
	if ( $EmailEnvia ) {
		mostra $et.90
		$chavesEmail = @($app_chaveRT) + $app_gestor
		if ( $debugOn ) {	$chavesEmail = $debugChaves	}
		$chavesIncluidas = @()
		$idWorkSpace = $wsExistenteId		# $IdWsAtiva	$IdWsNova
		$chavesEmail | ForEach-Object {
			$chave = $_
			if ( -not ($chave -in $chavesIncluidas) ) {	# para evitar duplicidade de inclusão de chave de RT que é a mesma de gestor
				./enviaEmail.ps1 -Oferta $Oferta -CodApl $CodApl -Registro $Registro -Premium $Premium -NomeAplic $app_nome -idWorkSpace $idWorkSpace -ChaveUsuario $chave -TipoEmail 'ok'
			}
			$chavesIncluidas += $chave
		}
	}
	#endregion
	#region fim --------------------------------------------------------------------------------------------------------------------------------
	mostra "$($et.99) - https://app.powerbi.com/groups/$wsExistenteId"
	#endregion
}
#region 99 - Fim -------------------------------------------------------------------------------------------------------------------------------
<# Diversos:
$NomeGrupoMs365 = "PBI_TRPB Grupo criado separadamente"
$oGrupoMS365 = Get-AzureADGroup -SearchString $NomeGrupoMs365
New-AzureADMSGroup -DisplayName $NomeGrupoMs365 -MailEnabled $false -MailNickName "$MailNickName-manual" -SecurityEnabled $true -GroupTypes "Unified" -Visibility Private
$authheader = Get-PowerBIAccessToken
  $grpid = "45295b08-1f1c-42a0-8a81-34e9220680f6"
  #"39e07cba-85eb-4b73-a753-9684e9a69b73"
  #"07187aa1-f87e-4e94-8ed7-1724d7f50c66"		#"70626bb6-e6e8-4cbe-b216-eb60445a7955"
  $uri   = "https://api.powerbi.com/v1.0/myorg/groups/$grpid"
  Invoke-RestMethod -Uri $uri -Method Delete -Headers $authheader 
  $uri = "https://api.powerbi.com/v1.0/myorg/groups"
  $uri = "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=20&%24skip=01"
# $uri = "https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.PowerBI/workspaceCollections/{workspaceCollectionName}/workspaces?api-version=2016-01-29"
  Invoke-RestMethod -Uri $uri -Method Get -Headers $authheader | select -ExpandProperty value | select id, state, type, name
# Invoke-RestMethod -Method GET -Uri $uri -Headers $authheader
# -Body ( @{identifier="$($oGrupoMS365.ObjectId)"; groupUserAccessRight='Viewer'; principalType='Group'} | ConvertTo-Json -Depth 2 -Compress )
#>
#endregion
