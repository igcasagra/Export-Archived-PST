# Script para criar pesquisas de eDiscovery para exportação de arquivo morto
# Método oficial: SearchExport (Compliance Search)

<#
.SYNOPSIS
    Cria pesquisa de eDiscovery para exportar arquivo morto de caixa de correio do M365.

.DESCRIPTION
    Este script cria automaticamente pesquisas de Compliance (eDiscovery) no Microsoft Purview
    com filtros de data para exportação de mensagens antigas do arquivo morto.
    A exportação final é feita manualmente pelo portal.

.PARAMETER Mailbox
    Email da caixa de correio que possui arquivo morto

.PARAMETER OlderThanDays
    Filtrar apenas mensagens mais antigas que X dias (exemplo: 730 para mais de 2 anos)

.PARAMETER StartDate
    Data inicial para filtrar mensagens (formato: yyyy-MM-dd)

.PARAMETER EndDate
    Data final para filtrar mensagens (formato: yyyy-MM-dd)

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -Mailbox "usuario@contoso.com" -OlderThanDays 730

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -Mailbox "usuario@contoso.com" -OlderThanDays 365

.EXAMPLE
    .\Export-ArchiveMailbox-EXO.ps1 -Mailbox "usuario@contoso.com" -StartDate "2020-01-01" -EndDate "2022-12-31"

.NOTES
    Requisitos:
    - Módulo ExchangeOnlineManagement
    - Permissões: eDiscovery Manager ou Compliance Administrator
    - Exportação manual pelo portal: https://purview.microsoft.com/contentsearch
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Mailbox,
    
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
    param(
        [string]$UserEmail,
        [int]$DaysFilter,
        [string]$Start,
        [string]$End
    )
    
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
                Write-Host "  ⚠️  Não foi possível obter estatísticas detalhadas" -ForegroundColor Yellow
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
            Write-Host "📧 $($folder.Name)" -ForegroundColor Yellow
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
            Write-Host "📁 $($folder.Name)" -ForegroundColor White
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
    
    Write-Host "`n⚠️  Este método:" -ForegroundColor Yellow
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
        Write-Host "⚠️  Configurando pesquisa com filtros..." -ForegroundColor Yellow
        
        # Monta o filtro de busca
        # IMPORTANTE: Busca em toda caixa (principal + arquivo morto juntos)
        # O Compliance Search não consegue separar apenas arquivo morto
        $searchQuery = "kind:email"
        
        # Adiciona filtro de data se especificado
        if ($OlderThanDays -gt 0) {
            $dateLimit = (Get-Date).AddDays(-$OlderThanDays).ToString("yyyy-MM-dd")
            $searchQuery += " AND received<$dateLimit"
            Write-Host "  📅 Filtrando mensagens mais antigas que $OlderThanDays dias (antes de $dateLimit)" -ForegroundColor Cyan
        }
        elseif ($StartDate -and $EndDate) {
            $searchQuery += " AND received>=$StartDate AND received<=$EndDate"
            Write-Host "  📅 Filtrando mensagens entre $StartDate e $EndDate" -ForegroundColor Cyan
        }
        
        Write-Host "  🔍 Query de busca: $searchQuery" -ForegroundColor Gray
        Write-Host "  ⚠️  AVISO: Compliance Search busca em TODA caixa (principal + arquivo)" -ForegroundColor Yellow
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
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "📋 Nome da pesquisa: " -NoNewline -ForegroundColor Cyan
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
        Write-Host "ℹ️  Nota: A exportação via PowerShell foi descontinuada pela Microsoft" -ForegroundColor DarkGray
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
    Write-Host "📋 MÉTODOS DISPONÍVEIS:" -ForegroundColor White
    Write-Host ""
    Write-Host "1️⃣  NEW-MAILBOXEXPORTREQUEST (PST) - Método Oficial" -ForegroundColor Cyan
    Write-Host "   ✅ Exporta diretamente para PST" -ForegroundColor Green
    Write-Host "   ✅ Preserva estrutura de pastas" -ForegroundColor Green
    Write-Host "   ❌ Requer permissão 'Mailbox Import Export'" -ForegroundColor Red
    Write-Host "   ❌ Requer caminho UNC (compartilhamento de rede)" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Comando:" -ForegroundColor White
    Write-Host "   .\Export-ArchiveMailbox-EXO.ps1 -Mailbox 'user@contoso.com' -ExportPath '\\\\servidor\\share' -Method PST" -ForegroundColor Gray
    Write-Host ""
    Write-Host "2️⃣  COMPLIANCE SEARCH (eDiscovery) - Criação de Pesquisa" -ForegroundColor Cyan
    Write-Host "   ✅ Cria pesquisa automaticamente com filtros" -ForegroundColor Green
    Write-Host "   ✅ Suporta filtros de data avançados" -ForegroundColor Green
    Write-Host "   ⚠️  Exportação manual pelo portal (Microsoft Purview)" -ForegroundColor Yellow
    Write-Host "   ⚠️  Busca em TODA caixa (principal + arquivo morto junto)" -ForegroundColor Yellow
    Write-Host "   ❌ Requer permissão eDiscovery Manager" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Comando:" -ForegroundColor White
    Write-Host "   .\Export-ArchiveMailbox-EXO.ps1 -Mailbox 'user@contoso.com' -ExportPath 'C:\Export' -Method SearchExport -OlderThanDays 365" -ForegroundColor Gray
    Write-Host "   (Cria a pesquisa, exportação manual no portal)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "3️⃣  GRAPH API (EML) - Via outro script" -ForegroundColor Cyan
    Write-Host "   ✅ Não requer permissões especiais" -ForegroundColor Green
    Write-Host "   ✅ Autenticação interativa" -ForegroundColor Green
    Write-Host "   ⚠️  Limitado a 1000 itens por vez" -ForegroundColor Yellow
    Write-Host "   ⚠️  Não acessa arquivo morto diretamente" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Use o script: Export-ArchiveMailbox.ps1" -ForegroundColor Gray
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "💡 RECOMENDAÇÃO:" -ForegroundColor Yellow
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
