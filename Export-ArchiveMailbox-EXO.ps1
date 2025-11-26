# Script para criar pesquisas de eDiscovery para exportação de arquivo morto
# Método oficial: SearchExport (Compliance Search)

<#
.SYNOPSIS
    Cria pesquisa de eDiscovery para exportar arquivo morto de caixa de correio do M365.

.DESCRIPTION
    Este script cria automaticamente pesquisas de Compliance (eDiscovery) no Microsoft Purview
    com filtros de data para exportação de mensagens antigas do arquivo morto.
    A exportação final é feita manualmente pelo portal.
    
    O script oferece duas formas de uso:
    - Modo interativo: Exibe TOP N mailboxes com maior uso ou permite digitar email
    - Modo direto: Especifica a mailbox via parâmetro

.PARAMETER Mailbox
    Email da caixa de correio que possui arquivo morto (opcional se usar -ShowTop10)

.PARAMETER ShowTop10
    Exibe as TOP N mailboxes com maior percentual de uso do arquivo morto

.PARAMETER TopN
    Número de mailboxes a exibir no ranking (padrão: 10, máximo: 50)

.PARAMETER OlderThanDays
    Filtrar apenas mensagens mais antigas que X dias (exemplo: 730 para mais de 2 anos)

.PARAMETER StartDate
    Data inicial para filtrar mensagens (formato: yyyy-MM-dd)

.PARAMETER EndDate
    Data final para filtrar mensagens (formato: yyyy-MM-dd)

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -ShowTop10
    Exibe TOP 10 mailboxes com maior uso e permite selecionar

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -ShowTop10 -TopN 5
    Exibe TOP 5 mailboxes com maior uso

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -ShowTop10 -TopN 15
    Exibe TOP 15 mailboxes com maior uso

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1
    Modo interativo: oferece opções de visualizar TOP N ou digitar email

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -Mailbox "usuario@contoso.com" -OlderThanDays 730
    Cria pesquisa para mensagens com mais de 2 anos

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -Mailbox "usuario@contoso.com" -StartDate "2020-01-01" -EndDate "2022-12-31"
    Cria pesquisa para período específico

.NOTES
    Requisitos:
    - Módulo ExchangeOnlineManagement
    - Permissões: eDiscovery Manager ou Compliance Administrator
    - Exportação manual pelo portal: https://purview.microsoft.com/contentsearch
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$Mailbox,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowTop10,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 50)]
    [int]$TopN = 10,
    
    [Parameter(Mandatory=$false)]
    [int]$OlderThanDays = 0,
    
    [Parameter(Mandatory=$false)]
    [string]$StartDate,
    
    [Parameter(Mandatory=$false)]
    [string]$EndDate
)

# Função para instalar módulo Exchange Online
function Install-ExchangeOnlineModule {
    Write-Host "Verificando módulo ExchangeOnlineManagement..." -ForegroundColor Cyan
    
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "Módulo não encontrado. Instalando..." -ForegroundColor Yellow
        try {
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
            Write-Host "✓ Módulo instalado com sucesso!" -ForegroundColor Green
        }
        catch {
            Write-Error "Erro ao instalar módulo: $_"
            return $false
        }
    }
    else {
        Write-Host "✓ Módulo já está instalado" -ForegroundColor Green
    }
    
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
    return $true
}

# Função para conectar ao Exchange Online
function Connect-ToExchangeOnline {
    Write-Host "Conectando ao Exchange Online..." -ForegroundColor Cyan
    
    try {
        # Tenta conectar
        Connect-ExchangeOnline -ShowBanner:$false
        
        Write-Host "✓ Conectado ao Exchange Online" -ForegroundColor Green
        
        # Verifica o usuário conectado
        $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        if ($orgConfig) {
            Write-Host "Organização: $($orgConfig.DisplayName)" -ForegroundColor Gray
        }
        
        return $true
    }
    catch {
        Write-Error "Erro ao conectar: $_"
        return $false
    }
}

# Função para listar TOP N mailboxes com maior uso de espaço
function Get-TopMailboxesByUsage {
    param([int]$Top = 10)
    
    Write-Host "`nBuscando mailboxes com arquivo morto ativo..." -ForegroundColor Cyan
    Write-Host "Isso pode levar alguns minutos..." -ForegroundColor Yellow
    
    try {
        # Busca todas as mailboxes com arquivo morto ativo
        $mailboxes = Get-Mailbox -ResultSize Unlimited -Filter {ArchiveStatus -eq 'Active'} | 
            Select-Object DisplayName, PrimarySmtpAddress, ArchiveQuota, ArchiveWarningQuota, ProhibitSendReceiveQuota, IssueWarningQuota
        
        Write-Host "Encontradas $($mailboxes.Count) mailboxes com arquivo morto ativo" -ForegroundColor Green
        Write-Host "Obtendo estatísticas de uso (arquivo morto + caixa principal)..." -ForegroundColor Cyan
        
        $mailboxStats = @()
        $counter = 0
        
        foreach ($mbx in $mailboxes) {
            $counter++
            Write-Progress -Activity "Coletando estatísticas" -Status "Processando $($mbx.DisplayName)" -PercentComplete (($counter / $mailboxes.Count) * 100)
            
            try {
                # Estatísticas do arquivo morto
                $archiveStats = Get-MailboxStatistics -Identity $mbx.PrimarySmtpAddress -Archive -ErrorAction SilentlyContinue
                
                # Estatísticas da caixa principal
                $primaryStats = Get-MailboxStatistics -Identity $mbx.PrimarySmtpAddress -ErrorAction SilentlyContinue
                
                if ($archiveStats) {
                    # Converte tamanho do ARQUIVO MORTO para MB
                    $archiveSizeMB = 0
                    if ($archiveStats.TotalItemSize -match '(\d+\.?\d*)\s*([KMGT]?B)') {
                        $size = [double]$matches[1]
                        $unit = $matches[2]
                        
                        switch ($unit) {
                            'KB' { $archiveSizeMB = $size / 1024 }
                            'MB' { $archiveSizeMB = $size }
                            'GB' { $archiveSizeMB = $size * 1024 }
                            'TB' { $archiveSizeMB = $size * 1024 * 1024 }
                            default { $archiveSizeMB = $size / (1024 * 1024) }
                        }
                    }
                    
                    # Converte tamanho da CAIXA PRINCIPAL para MB
                    $primarySizeMB = 0
                    if ($primaryStats -and $primaryStats.TotalItemSize -match '(\d+\.?\d*)\s*([KMGT]?B)') {
                        $size = [double]$matches[1]
                        $unit = $matches[2]
                        
                        switch ($unit) {
                            'KB' { $primarySizeMB = $size / 1024 }
                            'MB' { $primarySizeMB = $size }
                            'GB' { $primarySizeMB = $size * 1024 }
                            'TB' { $primarySizeMB = $size * 1024 * 1024 }
                            default { $primarySizeMB = $size / (1024 * 1024) }
                        }
                    }
                    
                    # Extrai quota do ARQUIVO MORTO em GB
                    # Se for "Unlimited", considera 1.5TB (1536 GB) para Exchange Plan 2
                    $archiveQuotaGB = 0
                    $isAutoExpanding = $false
                    
                    if ($mbx.ArchiveQuota -match 'Unlimited' -or $mbx.ArchiveQuota -match '1\.5 TB') {
                        $archiveQuotaGB = 1536  # 1.5 TB para auto-expanding archives (Exchange Plan 2)
                        $isAutoExpanding = $true
                    }
                    elseif ($mbx.ArchiveQuota -match '(\d+\.?\d*)\s*([KMGT]?B)') {
                        $qSize = [double]$matches[1]
                        $qUnit = $matches[2]
                        
                        switch ($qUnit) {
                            'KB' { $archiveQuotaGB = $qSize / (1024 * 1024) }
                            'MB' { $archiveQuotaGB = $qSize / 1024 }
                            'GB' { $archiveQuotaGB = $qSize }
                            'TB' { $archiveQuotaGB = $qSize * 1024 }
                            default { $archiveQuotaGB = 100 }
                        }
                    }
                    else {
                        $archiveQuotaGB = 100  # default quota
                    }
                    
                    # Extrai quota da CAIXA PRINCIPAL em GB
                    $primaryQuotaGB = 0
                    if ($mbx.ProhibitSendReceiveQuota -match '(\d+\.?\d*)\s*([KMGT]?B)') {
                        $qSize = [double]$matches[1]
                        $qUnit = $matches[2]
                        
                        switch ($qUnit) {
                            'KB' { $primaryQuotaGB = $qSize / (1024 * 1024) }
                            'MB' { $primaryQuotaGB = $qSize / 1024 }
                            'GB' { $primaryQuotaGB = $qSize }
                            'TB' { $primaryQuotaGB = $qSize * 1024 }
                            default { $primaryQuotaGB = 50 }
                        }
                    }
                    else {
                        $primaryQuotaGB = 50  # default quota
                    }
                    
                    $archivePercentUsed = if ($archiveQuotaGB -gt 0) { ($archiveSizeMB / 1024) / $archiveQuotaGB * 100 } else { 0 }
                    $primaryPercentUsed = if ($primaryQuotaGB -gt 0) { ($primarySizeMB / 1024) / $primaryQuotaGB * 100 } else { 0 }
                    
                    $mailboxStats += [PSCustomObject]@{
                        DisplayName = $mbx.DisplayName
                        Email = $mbx.PrimarySmtpAddress
                        ArchiveItemCount = $archiveStats.ItemCount
                        ArchiveSizeMB = [math]::Round($archiveSizeMB, 2)
                        ArchiveSizeGB = [math]::Round($archiveSizeMB / 1024, 2)
                        ArchiveQuotaGB = [math]::Round($archiveQuotaGB, 2)
                        ArchivePercentUsed = [math]::Round($archivePercentUsed, 2)
                        IsAutoExpanding = $isAutoExpanding
                        PrimaryItemCount = if ($primaryStats) { $primaryStats.ItemCount } else { 0 }
                        PrimarySizeGB = [math]::Round($primarySizeMB / 1024, 2)
                        PrimaryQuotaGB = [math]::Round($primaryQuotaGB, 2)
                        PrimaryPercentUsed = [math]::Round($primaryPercentUsed, 2)
                    }
                }
            }
            catch {
                # Ignora erros de mailboxes individuais
            }
        }
        
        Write-Progress -Activity "Coletando estatísticas" -Completed
        
        # Ordena por percentual de uso do arquivo e pega TOP N
        $topMailboxes = $mailboxStats | Sort-Object ArchivePercentUsed -Descending | Select-Object -First $Top
        
        return $topMailboxes
    }
    catch {
        Write-Error "Erro ao buscar mailboxes: $_"
        return @()
    }
}

# Função para exibir e selecionar mailbox
function Show-MailboxSelectionMenu {
    param([array]$Mailboxes)
    
    if ($Mailboxes.Count -eq 0) {
        Write-Host "Nenhuma mailbox encontrada." -ForegroundColor Yellow
        return $null
    }
    
    $title = "TOP $($Mailboxes.Count) MAILBOXES COM MAIOR USO DE ARQUIVO MORTO"
    $titlePadding = [Math]::Max(0, (127 - $title.Length) / 2)
    $paddedTitle = (" " * $titlePadding) + $title
    
    Write-Host "`n╔═══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║$($paddedTitle.PadRight(127))║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "ARQUIVO MORTO:" -ForegroundColor Yellow
    Write-Host ("{0,-3} {1,-30} {2,8} {3,9} {4,9} {5,8} | {6,9} {7,9} {8,8}" -f "#", "Nome", "Itens", "Tamanho", "Quota", "% Uso", "Tam.Prin", "Quota", "% Uso") -ForegroundColor White
    Write-Host ("{0,-3} {1,-30} {2,8} {3,9} {4,9} {5,8} | {6,9} {7,9} {8,8}" -f "---", "------------------------------", "--------", "---------", "---------", "--------", "---------", "---------", "--------") -ForegroundColor Gray
    
    for ($i = 0; $i -lt $Mailboxes.Count; $i++) {
        $mbx = $Mailboxes[$i]
        
        # Cor baseada no % de uso do arquivo
        $archiveColor = if ($mbx.ArchivePercentUsed -ge 90) { 'Red' } 
                        elseif ($mbx.ArchivePercentUsed -ge 75) { 'Yellow' } 
                        else { 'Green' }
        
        # Cor baseada no % de uso da caixa principal
        $primaryColor = if ($mbx.PrimaryPercentUsed -ge 90) { 'Red' } 
                        elseif ($mbx.PrimaryPercentUsed -ge 75) { 'Yellow' } 
                        else { 'Green' }
        
        # Indicador de auto-expanding
        $planType = if ($mbx.IsAutoExpanding) { " [Plan2]" } else { "" }
        
        Write-Host ("{0,-3}" -f ($i + 1)) -NoNewline -ForegroundColor Cyan
        Write-Host ("{0,-30}" -f ($mbx.DisplayName.Substring(0, [Math]::Min(28, $mbx.DisplayName.Length)) + $planType)) -NoNewline -ForegroundColor White
        Write-Host ("{0,8}" -f $mbx.ArchiveItemCount) -NoNewline -ForegroundColor Gray
        Write-Host ("{0,8} GB" -f $mbx.ArchiveSizeGB) -NoNewline -ForegroundColor Gray
        
        # Mostra quota do arquivo com indicador se for auto-expanding
        if ($mbx.IsAutoExpanding) {
            Write-Host ("{0,8} TB" -f ([math]::Round($mbx.ArchiveQuotaGB / 1024, 1))) -NoNewline -ForegroundColor DarkCyan
        }
        else {
            Write-Host ("{0,8} GB" -f $mbx.ArchiveQuotaGB) -NoNewline -ForegroundColor Gray
        }
        
        Write-Host ("{0,7}%" -f $mbx.ArchivePercentUsed) -NoNewline -ForegroundColor $archiveColor
        Write-Host " | " -NoNewline -ForegroundColor DarkGray
        Write-Host ("{0,8} GB" -f $mbx.PrimarySizeGB) -NoNewline -ForegroundColor Gray
        Write-Host ("{0,8} GB" -f $mbx.PrimaryQuotaGB) -NoNewline -ForegroundColor Gray
        Write-Host ("{0,7}%" -f $mbx.PrimaryPercentUsed) -ForegroundColor $primaryColor
        
        Write-Host ("    {0}" -f $mbx.Email) -ForegroundColor DarkGray
    }
    
    Write-Host ""
    Write-Host "Legenda: [Plan2] = Exchange Plan 2 com auto-expanding archive (até 1.5 TB)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Digite o número da mailbox (1-$($Mailboxes.Count)) ou 0 para digitar email manualmente: " -NoNewline -ForegroundColor Yellow
    
    $selection = Read-Host
    
    if ($selection -match '^\d+$') {
        $index = [int]$selection
        
        if ($index -eq 0) {
            Write-Host "Digite o email da mailbox: " -NoNewline -ForegroundColor Yellow
            $email = Read-Host
            return $email
        }
        elseif ($index -ge 1 -and $index -le $Mailboxes.Count) {
            return $Mailboxes[$index - 1].Email
        }
        else {
            Write-Host "Seleção inválida!" -ForegroundColor Red
            return $null
        }
    }
    else {
        Write-Host "Entrada inválida!" -ForegroundColor Red
        return $null
    }
}

# Função para verificar se o arquivo morto existe
function Test-ArchiveMailboxExists {
    param([string]$UserEmail)
    
    Write-Host "`nVerificando arquivo morto para: $UserEmail" -ForegroundColor Cyan
    
    try {
        $mailbox = Get-Mailbox -Identity $UserEmail -ErrorAction Stop
        
        Write-Host "Caixa de correio encontrada:" -ForegroundColor Green
        Write-Host "  Nome: $($mailbox.DisplayName)" -ForegroundColor Gray
        Write-Host "  Email: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Gray
        Write-Host "  Arquivo morto habilitado: $($mailbox.ArchiveStatus)" -ForegroundColor $(if($mailbox.ArchiveStatus -eq 'Active'){'Green'}else{'Yellow'})
        
        if ($mailbox.ArchiveStatus -eq 'Active') {
            Write-Host "  ✓ Arquivo morto está ATIVO e acessível" -ForegroundColor Green
            
            # Tenta obter estatísticas do arquivo
            try {
                $archiveStats = Get-MailboxFolderStatistics -Identity $UserEmail -Archive -ErrorAction SilentlyContinue
                $itemCount = ($archiveStats | Measure-Object -Property ItemsInFolder -Sum).Sum
                Write-Host "  Total de itens no arquivo: $itemCount" -ForegroundColor Cyan
            }
            catch {
                Write-Host "  Não foi possível obter estatísticas detalhadas" -ForegroundColor Yellow
            }
            
            return $true
        }
        else {
            Write-Warning "Arquivo morto não está ativo para este usuário"
            Write-Host "`nPara habilitar o arquivo morto:" -ForegroundColor Yellow
            Write-Host "  Enable-Mailbox -Identity '$UserEmail' -Archive" -ForegroundColor Gray
            return $false
        }
    }
    catch {
        Write-Error "Erro ao verificar caixa de correio: $_"
        return $false
    }
}

# Função para listar pastas do arquivo morto
function Get-ArchiveMailboxFolders {
    param([string]$UserEmail)
    
    Write-Host "`nListando pastas da caixa principal..." -ForegroundColor Cyan
    
    try {
        $mainFolders = Get-MailboxFolderStatistics -Identity $UserEmail | 
            Select-Object Name, FolderPath, ItemsInFolder, FolderSize |
            Sort-Object ItemsInFolder -Descending |
            Where-Object { $_.ItemsInFolder -gt 0 -or $_.Name -like "Caixa de Entrada" -or $_.Name -like "Inbox" -or $_.Name -like "Sent*" -or $_.Name -like "Itens Enviados" }
        
        Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║  Caixa de Correio Principal                                    ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        
        foreach ($folder in $mainFolders) {
            Write-Host ""
            Write-Host "$($folder.Name)" -ForegroundColor Yellow
            Write-Host "   Caminho: $($folder.FolderPath)" -ForegroundColor Gray
            Write-Host "   Itens: $($folder.ItemsInFolder)" -ForegroundColor Gray
            Write-Host "   Tamanho: $($folder.FolderSize)" -ForegroundColor Gray
        }
    }
    catch {
        Write-Warning "Não foi possível listar pastas da caixa principal: $_"
    }
    
    Write-Host "`n" -NoNewline
    Write-Host "Listando pastas do arquivo morto..." -ForegroundColor Cyan
    
    try {
        $folders = Get-MailboxFolderStatistics -Identity $UserEmail -Archive | 
            Select-Object Name, FolderPath, ItemsInFolder, FolderSize |
            Sort-Object ItemsInFolder -Descending
        
        Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║  In-Place Archive (Arquivo Morto)                              ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        
        foreach ($folder in $folders) {
            Write-Host ""
            Write-Host "$($folder.Name)" -ForegroundColor White
            Write-Host "   Caminho: $($folder.FolderPath)" -ForegroundColor Gray
            Write-Host "   Itens: $($folder.ItemsInFolder)" -ForegroundColor Gray
            Write-Host "   Tamanho: $($folder.FolderSize)" -ForegroundColor Gray
        }
        
        Write-Host ""
        return $folders
    }
    catch {
        Write-Error "Erro ao listar pastas: $_"
        return @()
    }
}

# Função para exportar usando Compliance Search (único método disponível)
function Export-ArchiveUsingComplianceSearch {
    param(
        [string]$UserEmail,
        [string]$OutputPath,
        [int]$OlderThanDays = 0,
        [string]$StartDate = "",
        [string]$EndDate = ""
    )
    
    Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  Compliance Search - Criação de Pesquisa Filtrada             ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    
    Write-Host "`nEste método:" -ForegroundColor Yellow
    Write-Host "   • Cria a pesquisa automaticamente com filtros de data" -ForegroundColor Gray
    Write-Host "   • Exportação deve ser feita manualmente no portal" -ForegroundColor Gray
    Write-Host "   • Requer: eDiscovery Manager ou Compliance Administrator" -ForegroundColor Gray
    Write-Host ""
    
    try {
        # Conecta ao Compliance Center
        Write-Host "Conectando ao Security & Compliance Center..." -ForegroundColor Cyan
        Connect-IPPSSession -ShowBanner:$false
        
        Write-Host "✓ Conectado" -ForegroundColor Green
        
        # Nome da pesquisa
        $searchName = "ArchiveExport_$($UserEmail -replace '@|\.','_')_$(Get-Date -Format 'yyyyMMddHHmmss')"
        
        Write-Host "`nCriando pesquisa de compliance..." -ForegroundColor Cyan
        Write-Host "Nome: $searchName" -ForegroundColor Gray
        
        # Cria pesquisa incluindo APENAS o arquivo morto (não a caixa principal)
        # Usar o formato especial: usuario@dominio.onmicrosoft.com (Archive)
        Write-Host "Configurando pesquisa com filtros..." -ForegroundColor Yellow
        
        # Monta o filtro de busca
        # IMPORTANTE: Busca em toda caixa (principal + arquivo morto juntos)
        # O Compliance Search não consegue separar apenas arquivo morto
        $searchQuery = "kind:email"
        
        # Adiciona filtro de data se especificado
        if ($OlderThanDays -gt 0) {
            $dateLimit = (Get-Date).AddDays(-$OlderThanDays).ToString("yyyy-MM-dd")
            $searchQuery += " AND received<$dateLimit"
            Write-Host "  Filtrando mensagens mais antigas que $OlderThanDays dias (antes de $dateLimit)" -ForegroundColor Cyan
        }
        elseif ($StartDate -and $EndDate) {
            $searchQuery += " AND received>=$StartDate AND received<=$EndDate"
            Write-Host "  Filtrando mensagens entre $StartDate e $EndDate" -ForegroundColor Cyan
        }
        
        Write-Host "  Query de busca: $searchQuery" -ForegroundColor Gray
        Write-Host "  AVISO: Compliance Search busca em TODA caixa (principal + arquivo)" -ForegroundColor Yellow
        Write-Host "             Para exportar SOMENTE arquivo morto, use -Method PST" -ForegroundColor Yellow
        
        # Cria pesquisa em toda a caixa de correio (não há como separar apenas arquivo morto)
        New-ComplianceSearch `
            -Name $searchName `
            -ExchangeLocation "$UserEmail" `
            -AllowNotFoundExchangeLocationsEnabled $true `
            -ContentMatchQuery $searchQuery | Out-Null
        
        Write-Host "✓ Pesquisa criada" -ForegroundColor Green
        
        # Inicia a pesquisa
        Write-Host "Iniciando pesquisa..." -ForegroundColor Cyan
        Start-ComplianceSearch -Identity $searchName
        
        # Monitora progresso
        Write-Host "Aguardando conclusão da pesquisa..." -ForegroundColor Yellow
        do {
            Start-Sleep -Seconds 10
            $search = Get-ComplianceSearch -Identity $searchName
            Write-Host "Status: $($search.Status) - Itens: $($search.Items)" -ForegroundColor Gray
        } while ($search.Status -ne "Completed")
        
        Write-Host "✓ Pesquisa concluída!" -ForegroundColor Green
        Write-Host "Total de itens encontrados: $($search.Items)" -ForegroundColor Cyan
        Write-Host "Tamanho total: $($search.Size)" -ForegroundColor Cyan
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║  PESQUISA CRIADA COM SUCESSO!                                  ║" -ForegroundColor Green
        Write-Host "═══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "Nome da pesquisa: " -NoNewline -ForegroundColor Cyan
        Write-Host "$searchName" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║  PRÓXIMOS PASSOS - EXPORTAÇÃO MANUAL                          ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "1. Acesse: " -NoNewline -ForegroundColor White
        Write-Host "https://compliance.microsoft.com/contentsearch" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "2. Localize a pesquisa: " -NoNewline -ForegroundColor White
        Write-Host "$searchName" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "3. Clique na pesquisa para abrir os detalhes" -ForegroundColor White
        Write-Host ""
        Write-Host "4. Clique no botão " -NoNewline -ForegroundColor White
        Write-Host "'Export results'" -ForegroundColor Green -NoNewline
        Write-Host " (na barra superior)" -ForegroundColor White
        Write-Host ""
        Write-Host "5. Configure as opções de exportação:" -ForegroundColor White
        Write-Host "   • Output options: escolha o formato desejado" -ForegroundColor Gray
        Write-Host "   • Export exchange content as: PST ou Individual messages" -ForegroundColor Gray
        Write-Host ""
        Write-Host "6. Após preparar a exportação, baixe usando o " -NoNewline -ForegroundColor White
        Write-Host "'eDiscovery Export Tool'" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Nota: A exportação via PowerShell foi descontinuada pela Microsoft" -ForegroundColor DarkGray
        Write-Host "   em maio de 2025. Agora é necessário exportar pelo portal." -ForegroundColor DarkGray
        Write-Host ""
        
        return $true
    }
    catch {
        Write-Error "Erro: $_"
        return $false
    }
}

# Função para mostrar informações
function Show-ArchiveExportInfo {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          MÉTODOS DE EXPORTAÇÃO DE ARQUIVO MORTO                ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "MÉTODOS DISPONÍVEIS:" -ForegroundColor White
    Write-Host ""
    Write-Host "1. NEW-MAILBOXEXPORTREQUEST (PST) - Método Oficial" -ForegroundColor Cyan
    Write-Host "   ✓ Exporta diretamente para PST" -ForegroundColor Green
    Write-Host "   ✓ Preserva estrutura de pastas" -ForegroundColor Green
    Write-Host "   X Requer permissão 'Mailbox Import Export'" -ForegroundColor Red
    Write-Host "   X Requer caminho UNC (compartilhamento de rede)" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Comando:" -ForegroundColor White
    Write-Host "   .\Export-ArchiveMailbox-EXO.ps1 -Mailbox 'user@contoso.com' -ExportPath '\\\\servidor\\share' -Method PST" -ForegroundColor Gray
    Write-Host ""
    Write-Host "2. COMPLIANCE SEARCH (eDiscovery) - Criação de Pesquisa" -ForegroundColor Cyan
    Write-Host "   ✓ Cria pesquisa automaticamente com filtros" -ForegroundColor Green
    Write-Host "   ✓ Suporta filtros de data avançados" -ForegroundColor Green
    Write-Host "   ! Exportação manual pelo portal (Microsoft Purview)" -ForegroundColor Yellow
    Write-Host "   ! Busca em TODA caixa (principal + arquivo morto junto)" -ForegroundColor Yellow
    Write-Host "   X Requer permissão eDiscovery Manager" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Comando:" -ForegroundColor White
    Write-Host "   .\Export-ArchiveMailbox-EXO.ps1 -Mailbox 'user@contoso.com' -ExportPath 'C:\Export' -Method SearchExport -OlderThanDays 365" -ForegroundColor Gray
    Write-Host "   (Cria a pesquisa, exportação manual no portal)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "3. GRAPH API (EML) - Via outro script" -ForegroundColor Cyan
    Write-Host "   ✓ Não requer permissões especiais" -ForegroundColor Green
    Write-Host "   ✓ Autenticação interativa" -ForegroundColor Green
    Write-Host "   ! Limitado a 1000 itens por vez" -ForegroundColor Yellow
    Write-Host "   ! Não acessa arquivo morto diretamente" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Use o script: Export-ArchiveMailbox.ps1" -ForegroundColor Gray
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "RECOMENDAÇÃO:" -ForegroundColor Yellow
    Write-Host "   Para exportar APENAS o arquivo morto:" -ForegroundColor White
    Write-Host "   → Use o Método 1 (PST) ou Método 2 (Compliance Search)" -ForegroundColor White
    Write-Host ""
    Write-Host "   O Graph API (Método 3) acessa a caixa principal," -ForegroundColor White
    Write-Host "   não o arquivo morto especificamente." -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
}

# ==== SCRIPT PRINCIPAL ====

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║    Criação de Pesquisa eDiscovery - Arquivo Morto             ║" -ForegroundColor Cyan
Write-Host "║    Microsoft Purview Compliance                                ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Instala módulo
if (-not (Install-ExchangeOnlineModule)) {
    exit 1
}

# Conecta ao Exchange Online
if (-not (Connect-ToExchangeOnline)) {
    exit 1
}

# Se não foi fornecido mailbox e não foi solicitado ShowTop10, pergunta ao usuário
if ([string]::IsNullOrWhiteSpace($Mailbox) -and -not $ShowTop10) {
    Write-Host "Escolha uma opção:" -ForegroundColor Yellow
    Write-Host "  1. Mostrar TOP $TopN mailboxes com maior uso" -ForegroundColor White
    Write-Host "  2. Digitar email da mailbox manualmente" -ForegroundColor White
    Write-Host ""
    Write-Host "Opção (1 ou 2): " -NoNewline -ForegroundColor Yellow
    $option = Read-Host
    
    if ($option -eq "1") {
        $ShowTop10 = $true
    }
    elseif ($option -eq "2") {
        Write-Host "Digite o email da mailbox: " -NoNewline -ForegroundColor Yellow
        $Mailbox = Read-Host
    }
    else {
        Write-Host "Opção inválida!" -ForegroundColor Red
        exit 1
    }
}

# Se ShowTop10 foi solicitado, busca e exibe as mailboxes
if ($ShowTop10) {
    $topMailboxes = Get-TopMailboxesByUsage -Top $TopN
    
    if ($topMailboxes.Count -eq 0) {
        Write-Host "Nenhuma mailbox com arquivo morto encontrada." -ForegroundColor Yellow
        exit 1
    }
    
    $Mailbox = Show-MailboxSelectionMenu -Mailboxes $topMailboxes
    
    if ([string]::IsNullOrWhiteSpace($Mailbox)) {
        Write-Host "Nenhuma mailbox selecionada." -ForegroundColor Yellow
        exit 1
    }
}

# Valida que temos uma mailbox para processar
if ([string]::IsNullOrWhiteSpace($Mailbox)) {
    Write-Host "Erro: Nenhuma mailbox especificada!" -ForegroundColor Red
    Write-Host "Use: .\Export-ArchiveMailbox-EXO.ps1 -Mailbox 'email@dominio.com'" -ForegroundColor Yellow
    Write-Host "Ou:  .\Export-ArchiveMailbox-EXO.ps1 -ShowTop10 -TopN 15" -ForegroundColor Yellow
    exit 1
}

Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  Mailbox selecionada: $($Mailbox.PadRight(45))" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green

# Verifica se o arquivo morto existe
if (-not (Test-ArchiveMailboxExists -UserEmail $Mailbox)) {
    Write-Host "`nNão é possível continuar sem arquivo morto ativo." -ForegroundColor Red
    exit 1
}

# Lista pastas do arquivo
Get-ArchiveMailboxFolders -UserEmail $Mailbox

# Executa SearchExport (único método disponível)
Write-Host ""
$success = Export-ArchiveUsingComplianceSearch -UserEmail $Mailbox `
                                               -OlderThanDays $OlderThanDays `
                                               -StartDate $StartDate `
                                               -EndDate $EndDate

# Desconecta
Write-Host "`nDesconectando..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "✓ Concluído!" -ForegroundColor Green

exit $(if($success){0}else{1})
