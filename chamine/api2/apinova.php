<?php
// Definir cabeçalhos para aceitar requisições JSON
header("Access-Control-Allow-Origin: *");
header("Content-Type: application/json; charset=UTF-8");
header("Access-Control-Allow-Methods: POST");

// Capturar os dados do POST
$jsonData = file_get_contents("php://input");

// Verificar se recebeu algum dado
if (!$jsonData) {
    echo json_encode(["status" => "erro", "mensagem" => "Nenhum dado recebido"]);
    exit;
}

// Converter JSON para array PHP
$data = json_decode($jsonData, true);

// Verificar se o JSON é válido
if (!$data) {
    echo json_encode(["status" => "erro", "mensagem" => "JSON inválido"]);
    exit;
}

// Configuração do banco de dados (ajuste conforme necessário)
$host = "localhost";
$dbname = "chamine";
$username = "root";
$password = "195575";

// Conectar ao banco de dados
try {
    $pdo = new PDO("mysql:host=$host;dbname=$dbname;charset=utf8", $username, $password);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
} catch (PDOException $e) {
    echo json_encode(["status" => "erro", "mensagem" => "Erro ao conectar ao banco: " . $e->getMessage()]);
    exit;
}

// Função para conversão e validação de valores decimais
function parseDecimal($value) {
    $value = str_replace(',', '.', $value);
    return is_numeric($value) ? $value : null;
}

// Inserir o perfil na tabela `tabela_perfil`
$stmt = $pdo->prepare("
INSERT INTO tabela_perfil (
    projeto, fonte, data_execucao, equipe, empresa_nome, cidade, processo, parametros_medidos, 
    montante, jusante, diametro_medio, jdc, mdc, area_chamine, comprimento, 
    diametro_equivalente, numero_de_pontos, numero_de_eixos, numero_de_pontos_por_eixo, 
    tempo_total_coleta, tempo_por_ponto
) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
");

$stmt->execute([
    $data["perfil"]["Projeto"],
    $data["perfil"]["Fonte"],
    date("Y-m-d H:i:s", strtotime($data["perfil"]["Data"])),
    $data["perfil"]["Equipe"],
    $data["perfil"]["Empresa"],
    $data["perfil"]["Cidade"],
    $data["perfil"]["Processo"],
    $data["perfil"]["Parametros"],

    parseDecimal($data["perfil"]["Montante"]),
    parseDecimal($data["perfil"]["Jusante"]),
    parseDecimal($data["perfil"]["DiametroMedio"]),
    parseDecimal($data["perfil"]["Jdc"]),
    parseDecimal($data["perfil"]["Mdc"]),
    parseDecimal($data["perfil"]["AreaChamine"]),
    parseDecimal($data["perfil"]["Comprimento"]),
    parseDecimal($data["perfil"]["Diametroequivalente"]),

    (int) str_replace(',', '.', $data["perfil"]["NumeroDePontos"]),
    (int) str_replace(',', '.', $data["perfil"]["NumeroDeEixos"]),
    (int) str_replace(',', '.', $data["perfil"]["NumeroDePontosPorEixo"]),
    (int) str_replace(',', '.', $data["perfil"]["TempoTotalColeta"]),
    (int) str_replace(',', '.', $data["perfil"]["TempoporPonto"])
]);

// Obter o ID do perfil inserido
$id_perfil = $pdo->lastInsertId();

// Preparar a inserção de parâmetros
$stmt_parametro = $pdo->prepare("INSERT IGNORE INTO tabela_parametros (nome, unidade) VALUES (?, ?)");
$stmt_valor = $pdo->prepare("INSERT INTO tabela_valores_coletas (id_perfil, id_parametro, numero_coleta, valor) VALUES (?, ?, ?, ?)");

// Inserir os parâmetros e valores das coletas
foreach ($data["coletas"] as $coleta) {
    $nome_parametro = trim($coleta["parametro"]);
    $unidade = trim($coleta["unidade"]);

    // Inserir o parâmetro (caso não exista)
    $stmt_parametro->execute([$nome_parametro, $unidade]);

    // Obter o ID do parâmetro inserido ou existente
    $id_parametro = $pdo->lastInsertId();
    if ($id_parametro == 0) {
        $stmt_busca = $pdo->prepare("SELECT id_parametro FROM tabela_parametros WHERE nome = ?");
        $stmt_busca->execute([$nome_parametro]);
        $id_parametro = $stmt_busca->fetchColumn();
    }

    // Inserir os valores das coletas (com número da coleta)
    for ($i = 1; $i <= 3; $i++) {
        $valor = str_replace(",", ".", $coleta["coleta_$i"]); // Troca vírgula por ponto para evitar erro no decimal
        if ($valor !== "" && is_numeric($valor)) {
            $stmt_valor->execute([$id_perfil, $id_parametro, $i, $valor]); // Agora armazenamos também o número da coleta!
        }
    }
}

// Responder com sucesso
echo json_encode(["status" => "sucesso", "mensagem" => "Dados inseridos com sucesso", "id_perfil" => $id_perfil]);


