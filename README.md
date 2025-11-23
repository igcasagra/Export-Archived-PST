# 📧 Exportação de E-mails via eDiscovery

Script PowerShell para criar pesquisas de eDiscovery e exportar e-mails de caixas de correio do Microsoft 365 com filtros por idade de mensagens.

## 📋 Descrição

O script `Export-ArchiveMailbox-EXO.ps1` permite criar pesquisas de Compliance (eDiscovery) automaticamente no Microsoft 365 para exportar e-mails com mais de X dias (por exemplo, mais de 2 anos).

**Método:** SearchExport (eDiscovery)
- Cria a pesquisa automaticamente no portal
- Aplica filtros de data (mensagens antigas)
- Exportação final é manual pelo portal do Microsoft Purview

## 🔧 Requisitos

### Módulo PowerShell
- **ExchangeOnlineManagement** (instalado automaticamente pelo script se necessário)

### Permissões Microsoft 365
- **eDiscovery Manager** OU
- **Compliance Administrator**

### Sistema
- Windows 10/11 ou Windows Server
- PowerShell 5.1 ou superior
- Conexão com internet

## 🎯 Como Atribuir Permissões

### Opção 1: Usar Script Automatizado (Recomendado)

Execute o script de configuração de permissões:

```powershell
.\Configure-eDiscoveryPermissions.ps1 -UserEmail "admin@contoso.com"
```

**Parâmetros:**
- `-UserEmail`: Email do usuário que receberá as permissões
- `-RoleGroup`: (Opcional) `eDiscoveryManager` (padrão) ou `eDiscoveryAdministrator`

**Exemplo com Administrator:**
```powershell
.\Configure-eDiscoveryPermissions.ps1 -UserEmail "admin@contoso.com" -RoleGroup "eDiscoveryAdministrator"
```

**Requisitos para executar o script:**
- Permissões de Administrador Global ou Compliance Administrator
- Módulo ExchangeOnlineManagement (instalado automaticamente)

### Opção 2: Configuração Manual pelo Portal

1. Acesse: https://purview.microsoft.com
2. Vá em **Permissions** → **Roles**
3. Selecione **eDiscovery Manager**
4. Clique em **Edit** na seção de membros
5. Adicione o usuário desejado
6. Salve as alterações

> ⚠️ **Importante:** Aguarde ~15 minutos para propagação das permissões após a configuração

## 🚀 Como Usar

### Sintaxe Básica

```powershell
.\Export-ArchiveMailbox-EXO.exe -Mailbox <email> -Method SearchExport -OlderThanDays <dias>
```

### Parâmetros

| Parâmetro | Obrigatório | Descrição | Exemplo |
|-----------|-------------|-----------|---------|
| `-Mailbox` | ✅ Sim | E-mail da caixa de correio | `"usuario@dominio.com"` |
| `-Method` | ✅ Sim | Método de exportação | `SearchExport` |
| `-OlderThanDays` | ❌ Não | Mensagens com mais de X dias | `730` (2 anos), `365` (1 ano) |
| `-StartDate` | ❌ Não | Data inicial do filtro | `"2020-01-01"` |
| `-EndDate` | ❌ Não | Data final do filtro | `"2023-12-31"` |

## 📝 Exemplos Práticos

### 1. E-mails com Mais de 2 Anos (730 dias)
```powershell
.\Export-ArchiveMailbox-EXO.exe -Mailbox "liliane.maus@leaderlog.com.br" -Method SearchExport -OlderThanDays 730
```

### 2. E-mails com Mais de 1 Ano (365 dias)
```powershell
.\Export-ArchiveMailbox-EXO.exe -Mailbox "usuario@dominio.com" -Method SearchExport -OlderThanDays 365
```

### 3. E-mails com Mais de 3 Anos (1095 dias)
```powershell
.\Export-ArchiveMailbox-EXO.exe -Mailbox "usuario@dominio.com" -Method SearchExport -OlderThanDays 1095
```

### 4. Período Específico (Entre Datas)
```powershell
.\Export-ArchiveMailbox-EXO.exe -Mailbox "usuario@dominio.com" -Method SearchExport -StartDate "2020-01-01" -EndDate "2022-12-31"
```

## 🔄 Fluxo de Trabalho Completo

### Passo 1: Executar o Script
```powershell
.\Export-ArchiveMailbox-EXO.exe -Mailbox "usuario@dominio.com" -Method SearchExport -OlderThanDays 730
```

### Passo 2: O Script Irá
1. ✅ Conectar ao Exchange Online (solicitará login)
2. ✅ Verificar se o arquivo morto está ativo
3. ✅ Listar todas as pastas com itens
4. ✅ Criar pesquisa de Compliance com filtros
5. ✅ Executar a pesquisa automaticamente
6. ✅ Exibir resultado (total de itens e tamanho)
7. ✅ Fornecer instruções para exportação

### Passo 3: Exportar pelo Portal
Após o script criar a pesquisa, você precisa exportar manualmente:

1. Acesse: https://purview.microsoft.com/contentsearch
2. Localize a pesquisa criada (nome: `ArchiveExport_...`)
3. Clique na pesquisa para ver detalhes
4. Clique no botão **"Export results"** (barra superior)
5. Configure as opções:
   - **Export exchange content as:** PST ou Individual messages
   - Marque as opções desejadas
6. Clique em **Export**
7. Aguarde a preparação (pode levar minutos/horas dependendo do tamanho)
8. Baixe usando o **eDiscovery Export Tool**

## 📊 O Que o Script Faz

### ✅ Automático (Pelo Script)
- Conexão ao Exchange Online
- Verificação do arquivo morto
- Listagem de pastas e itens
- Criação da pesquisa de Compliance
- Aplicação de filtros de data
- Execução da pesquisa
- Exibição de resultados

### ⚠️ Manual (Pelo Portal)
- Exportação dos resultados
- Download dos arquivos

## 🎯 Exemplo de Saída do Script

```
╔════════════════════════════════════════════════════════════════╗
║  Exportação de Arquivo Morto (Archive)                         ║
║  Exchange Online PowerShell                                    ║
╚════════════════════════════════════════════════════════════════╝

Conectando ao Exchange Online...
✓ Conectado

Verificando arquivo morto...
✓ Arquivo morto ativo

Listando pastas do arquivo morto...
📁 Clientes 2022 até 2025 - 6316 itens (10.16 GB)
📁 Beira Rio - 3836 itens (10.03 GB)
📁 Armadores - 3221 itens (5.45 GB)
...

╔════════════════════════════════════════════════════════════════╗
║  Compliance Search - Criação de Pesquisa Filtrada             ║
╚════════════════════════════════════════════════════════════════╝

Criando pesquisa de compliance...
Nome: ArchiveExport_usuario_dominio_com_20251122123456
  📅 Filtrando mensagens mais antigas que 730 dias
  🔍 Query de busca: kind:email AND received<2023-11-23
✓ Pesquisa criada

Iniciando pesquisa...
Status: Completed - Itens: 25002
✓ Pesquisa concluída!
Total de itens encontrados: 25002
Tamanho total: ~22 GB

╔════════════════════════════════════════════════════════════════╗
║  PESQUISA CRIADA COM SUCESSO!                                  ║
╚════════════════════════════════════════════════════════════════╝

📋 Nome da pesquisa: ArchiveExport_usuario_dominio_com_20251122123456

╔════════════════════════════════════════════════════════════════╗
║  PRÓXIMOS PASSOS - EXPORTAÇÃO MANUAL                          ║
╚════════════════════════════════════════════════════════════════╝

1. Acesse: https://purview.microsoft.com/contentsearch
2. Localize a pesquisa criada
3. Clique em "Export results"
...
```

## ⚠️ Avisos Importantes

### Sobre o Método SearchExport
- ✅ Cria pesquisa automaticamente
- ✅ Aplica filtros de data
- ⚠️ Busca em **TODA a caixa** (principal + arquivo morto)
- ⚠️ Exportação final é **manual** pelo portal
- ℹ️ Microsoft descontinuou exportação via PowerShell em maio/2025

### Sobre Permissões
- Sem permissão **eDiscovery Manager**, o script falhará
- A permissão pode levar alguns minutos para ser efetivada
- Requer autenticação MFA (multifator)

### Sobre Arquivo Morto
- O arquivo morto precisa estar **ativo**
- Se não estiver ativo, o script informará
- Para ativar: `Enable-Mailbox -Identity "usuario@dominio.com" -Archive`

## 🛠️ Troubleshooting

### Erro: "Access denied to compliance search"
**Causa:** Usuário não tem permissão eDiscovery Manager

**Solução:**
1. Acesse: https://purview.microsoft.com/permissions
2. Adicione o usuário ao grupo **eDiscovery Manager**
3. Aguarde 5-10 minutos
4. Tente novamente

### Erro: "Arquivo morto não está ativo"
**Causa:** A caixa de correio não tem arquivo morto habilitado

**Solução:**
```powershell
Connect-ExchangeOnline
Enable-Mailbox -Identity "usuario@dominio.com" -Archive
```

### Erro: "User canceled authentication"
**Causa:** Login foi cancelado ou credenciais incorretas

**Solução:**
- Complete o processo de login
- Verifique suas credenciais
- Certifique-se de ter acesso ao tenant

### Erro: "Module ExchangeOnlineManagement not found"
**Causa:** Módulo não está instalado

**Solução:** O script instala automaticamente. Se falhar:
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

## 📚 Recursos Adicionais

### Portais Microsoft
- **Microsoft Purview:** https://purview.microsoft.com
- **Content Search:** https://purview.microsoft.com/contentsearch
- **Exchange Admin Center:** https://admin.exchange.microsoft.com

### Documentação Oficial Microsoft
- [eDiscovery no Microsoft 365](https://learn.microsoft.com/microsoft-365/compliance/ediscovery)
- [Content Search](https://learn.microsoft.com/microsoft-365/compliance/content-search)
- [Exportar resultados de pesquisa](https://learn.microsoft.com/microsoft-365/compliance/export-search-results)

## 🔐 Segurança e Conformidade

- ✅ Usa autenticação moderna do Microsoft 365
- ✅ Requer MFA (autenticação multifator)
- ✅ Todas as operações são registradas no audit log
- ✅ Requer permissões específicas (princípio do menor privilégio)
- ✅ Não armazena credenciais

## 📄 Versões

### Executável vs Script PowerShell

**`Export-ArchiveMailbox-EXO.exe`** (Recomendado)
- ✅ Pode ser executado diretamente
- ✅ Não requer permissões de execução de script
- ✅ Mais fácil para usuários finais

**`Export-ArchiveMailbox-EXO.ps1`**
- ✅ Código-fonte aberto
- ✅ Pode ser modificado
- ⚠️ Requer `Set-ExecutionPolicy` adequado

## 💡 Dicas Práticas

### Para Múltiplos Usuários
Execute o script para cada usuário separadamente:
```powershell
$usuarios = @("user1@dominio.com", "user2@dominio.com", "user3@dominio.com")
foreach ($user in $usuarios) {
    .\Export-ArchiveMailbox-EXO.exe -Mailbox $user -Method SearchExport -OlderThanDays 730
}
```

### Para Diferentes Períodos
```powershell
# Mais de 1 ano
.\Export-ArchiveMailbox-EXO.exe -Mailbox "usuario@dominio.com" -Method SearchExport -OlderThanDays 365

# Mais de 2 anos
.\Export-ArchiveMailbox-EXO.exe -Mailbox "usuario@dominio.com" -Method SearchExport -OlderThanDays 730

# Mais de 5 anos
.\Export-ArchiveMailbox-EXO.exe -Mailbox "usuario@dominio.com" -Method SearchExport -OlderThanDays 1825
```

### Verificar Pesquisas Criadas
```powershell
Connect-IPPSSession
Get-ComplianceSearch | Where-Object {$_.Name -like "ArchiveExport*"} | Select-Object Name, Items, Status
```

## ❓ Perguntas Frequentes

**P: Por que a exportação não é automática?**
R: A Microsoft descontinuou a exportação automática via PowerShell em maio de 2025. Agora é obrigatório usar o portal.

**P: Posso exportar apenas o arquivo morto?**
R: O SearchExport busca em toda a caixa (principal + arquivo). Para filtrar, use os parâmetros de data.

**P: Quanto tempo leva para criar a pesquisa?**
R: Geralmente de segundos a poucos minutos, dependendo do número de mensagens.

**P: Quanto tempo leva para preparar a exportação?**
R: Pode variar de minutos a horas, dependendo do tamanho total dos dados.

**P: Preciso deixar o PowerShell aberto durante a exportação?**
R: Não. Após criar a pesquisa, você pode fechar. A exportação pelo portal é independente.


---

**Última Atualização:** Novembro 2025  
**Versão:** 1.0  
**Compatível com:** Exchange Online, Microsoft 365  
**Método:** SearchExport (eDiscovery)
