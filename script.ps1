Clear-Host
$WarningPreference = "SilentlyContinue"

Write-Host "Realize o login para continuar"
$usuarioLogado = Connect-MicrosoftTeams

function criarEquipe {
    $nomeEquipe = Read-Host "Digite o nome da equipe: "
    try {
        New-Team -DisplayName $nomeEquipe
        Write-Host "Equipe criada com sucesso"
    } catch {
        Write-Host "Erro ao criar a equipe"
    }
}
function adicionarMembrosExcel {
    $excel = ".\equipe.xlsx" #coloque o caminho do excel
    $dados = Import-Excel -Path $excel

    foreach ($row in $dados) {
        if (![string]::IsNullOrWhiteSpace($row.'E-mail')){
            $emailUsuario = $row.'E-mail'
            $grupo = Get-team -User $usuarioLogado.account -DisplayName $row.Time -ErrorAction SilentlyContinue
            $cargoUsuario = $row.Cargo
            try { 
                Add-TeamUser -GroupId $grupo.groupId -User $emailUsuario -Role $cargoUsuario
                Write-Host "Sucesso! $emailUsuario foi adicionado no time!"
            } catch { 
                Write-Host "Erro ao tentar adicionar usuário: $emailUsuario"
            }
        } else {
            break
        }
    }
}


function adicionarMembroManual {
    $emailMembro = Read-Host "Informe o email do membro: "
    $nomeEquipe = Read-Host "Informe o nome da equipe: "
    $cargoUsuario = Read-Host 'Digite "Owner" para adicionar o cargo propritário ou "Member" para adicionar o cargo membro '

    try {
        $idEquipe = Get-team -User $usuarioLogado.account -DisplayName $nomeEquipe -ErrorAction SilentlyContinue
        Add-TeamUSer -GroupId $idEquipe.groupId -User $emailMembro -Role $cargoUsuario
        Write-Host "Usuario $emailMembro foi adicionado com sucesso!"
    } catch {
        Write-Host "Erro ao tentar adicionar membro na equipe"
    }
}

function removerEquipe {
    $nomeEquipe = Read-Host "Digite o nome da equipe: "
    $idEquipe = Get-team -User $usuarioLogado.account -DisplayName $nomeEquipe -ErrorAction SilentlyContinue

    try {
        Remove-Team -GroupId $idEquipe.groupId
        Write-Host "Equipe deletada com sucesso"
    } catch {
        Write-Host "Erro ao deletar a equipe"
    } 
}


$opt = 1
while ($opt -ge 0) {
    Write-Host "Gerenciando o Microsoft Teams!"
    Write-Host "1 - Listar suas equipes"
    Write-Host "2 - Criar equipe"
    Write-Host "3 - Adicionar membros a partir de um arquivo excel"
    Write-Host "4 - Adicionar um membro na equipe"
    Write-Host "5 - Excluir uma equipe:"
    Write-Host "6 - Sair"
    #----------------------------------
    $opt = Read-Host "Digite uma opção: "

    switch ($opt) {
        "1" { Get-team -user $usuarioLogado.account }
        "2" { criarEquipe }
        "3" { adicionarMembrosExcel }
        "4" { adicionarMembroManual } 
        "5" { removerEquipe}
        "6" { 
            Write-Host "Adeus!"
            Exit
        }
        Default {
            Write-Host "Opção inválida, tente novamente!"
        }
    }
}



