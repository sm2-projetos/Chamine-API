<?php
// Configuração do banco de dados
$host = "localhost";
$user = "root"; // Altere conforme sua configuração
$password = "Z!X@c3v4"; // Altere conforme sua configuração
$database = "sistemachamine"; // Nome do banco de dados

// Conectar ao banco
$conn = new mysqli($host, $user, $password, $database);
if ($conn->connect_error) {
    die(json_encode([
        'status' => 'error',
        'message' => 'Database connection failed: ' . $conn->connect_error
    ]));
}

// Configurar cabeçalho de resposta
header('Content-Type: application/json');

// Se for uma requisição GET, apenas retorna uma mensagem
if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    echo json_encode([
        'status' => 'success',
        'message' => 'Welcome to the API'
    ]);
    exit;
}

// Se for uma requisição POST, processa os dados
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $postData = file_get_contents('php://input');
    $data = json_decode($postData, true);
    var_dump($data);
    if (json_last_error() === JSON_ERROR_NONE) {
        $conn->begin_transaction(); // Inicia transação para evitar inserções parciais
        
        try {
            // Inserir os dados do PerfilChamine
            if (isset($data['Perfil']) && is_array($data['Perfil'])) {
                $perfil = $data['Perfil'];
                
                $stmt = $conn->prepare("
                    INSERT INTO perfil_chamine 
                    (projeto, fonte, data, equipe, empresa, cidade, processo, parametros, montante, jusante, diametro1, diametro2, diametro3, diametro_medio, jdc, mdc, area_chamine, comprimento, diametro_equivalente, numero_de_pontos, numero_de_eixos, numero_de_pontos_por_eixo, tempo_total_coleta, tempo_por_ponto) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ");
                $stmt->bind_param("ssssssssssssssssssssssss",
                    $perfil['Projeto'], $perfil['Fonte'], $perfil['Data'], $perfil['Equipe'], $perfil['Empresa'], $perfil['Cidade'], $perfil['Processo'], $perfil['Parametros'], 
                    $perfil['Montante'], $perfil['Jusante'], $perfil['Diametro1'], $perfil['Diametro2'], $perfil['Diametro3'], $perfil['DiametroMedio'], 
                    $perfil['Jdc'], $perfil['Mdc'], $perfil['AreaChamine'], $perfil['Comprimento'], $perfil['Diametroequivalente'], 
                    $perfil['NumeroDePontos'], $perfil['NumeroDeEixos'], $perfil['NumeroDePontosPorEixo'], $perfil['TempoTotalColeta'], $perfil['TempoporPonto']
                );
                $stmt->execute();
                $perfilChamineId = $stmt->insert_id; // Pega o ID do perfil inserido
                $stmt->close();
            } else {
                throw new Exception("Perfil inválido");
            }

            // Inserir os dados de Nível
            if (isset($data['Nivel']) && is_array($data['Nivel'])) {
                $nivel = $data['Nivel'];

                $stmt = $conn->prepare("
                    INSERT INTO nivel 
                    (perfil_chamine_id, data_da_coleta_media, data_da_coleta_limite, horario_inicial_media, horario_inicial_limite, 
                    horario_final_media, horario_final_limite, temperatura_media, temperatura_limite, umidade_media, umidade_limite, 
                    o2_media, o2_limite, taxa_de_emissao_co2_media, taxa_de_emissao_co2_limite, taxa_de_emissao_o2_media, taxa_de_emissao_o2_limite, 
                    taxa_de_emissao_co_media, taxa_de_emissao_co_limite, taxa_de_emissao_n2_media, taxa_de_emissao_n2_limite, velocidade_media, 
                    velocidade_limite, vazao1_media, vazao1_limite, vazao2_media, vazao2_limite, concentracao_de_material_particulado1_media, 
                    concentracao_de_material_particulado1_limite, concentracao_de_material_particulado2_media, concentracao_de_material_particulado2_limite, 
                    taxa_de_emissao_de_material_particulado_media, taxa_de_emissao_de_material_particulado_limite, concentracao_de_monoxido_de_carbono1_media, 
                    concentracao_de_monoxido_de_carbono1_limite, concentracao_de_monoxido_de_carbono2_media, concentracao_de_monoxido_de_carbono2_limite, 
                    concentracao_de_dioxido_de_nitrogenio1_media, concentracao_de_dioxido_de_nitrogenio1_limite, concentracao_de_dioxido_de_nitrogenio2_media, 
                    concentracao_de_dioxido_de_nitrogenio2_limite, taxa_de_emissao_de_dioxido_de_nitrogenio_media, taxa_de_emissao_de_dioxido_de_nitrogenio_limite, 
                    isocinetica_media, isocinetica_limite, numero_do_bequer_branco, numero_do_filtro_branco, massa_do_bequer_branco, massa_do_filtro_branco, 
                    massa_total_branco, volume_branco, concentracao_de_material_particulado3_branco, incerteza_de_medicao_de_material_particulado_branco) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ");
                $stmt->bind_param("issssssssssssssssssssssssssssssssssssssssssssssssssss",
                    $perfilChamineId,
                    $nivel['DataDaColeta']['Media'], $nivel['DataDaColeta']['Limite'],
                    $nivel['HorarioInicial']['Media'], $nivel['HorarioInicial']['Limite'],
                    $nivel['Horariofinal']['Media'], $nivel['Horariofinal']['Limite'],
                    $nivel['Temperatura']['Media'], $nivel['Temperatura']['Limite'],
                    $nivel['Umidade']['Media'], $nivel['Umidade']['Limite'],
                    $nivel['O2']['Media'], $nivel['O2']['Limite'],
                    $nivel['TaxaDeEmissaoCO2']['Media'], $nivel['TaxaDeEmissaoCO2']['Limite'],
                    $nivel['TaxaDeEmissaoO2']['Media'], $nivel['TaxaDeEmissaoO2']['Limite'],
                    $nivel['TaxaDeEmissaoCO']['Media'], $nivel['TaxaDeEmissaoCO']['Limite'],
                    $nivel['TaxaDeEmissaoN2']['Media'], $nivel['TaxaDeEmissaoN2']['Limite'],
                    $nivel['Velocidade']['Media'], $nivel['Velocidade']['Limite'],
                    $nivel['Vazao1']['Media'], $nivel['Vazao1']['Limite'],
                    $nivel['Vazao2']['Media'], $nivel['Vazao2']['Limite'],
                    $nivel['ConcentracaoDeMaterialParticulado1']['Media'], $nivel['ConcentracaoDeMaterialParticulado1']['Limite'],
                    $nivel['ConcentracaoDeMaterialParticulado2']['Media'], $nivel['ConcentracaoDeMaterialParticulado2']['Limite'],
                    $nivel['TaxaDeEmissaoDeMaterialParticulado']['Media'], $nivel['TaxaDeEmissaoDeMaterialParticulado']['Limite'],
                    $nivel['ConcentracaoDeMonoxidoDeCarbono1']['Media'], $nivel['ConcentracaoDeMonoxidoDeCarbono1']['Limite'],
                    $nivel['ConcentracaoDeMonoxidoDeCarbono2']['Media'], $nivel['ConcentracaoDeMonoxidoDeCarbono2']['Limite'],
                    $nivel['ConcentracaoDeDioxidoDeNitrogenio1']['Media'], $nivel['ConcentracaoDeDioxidoDeNitrogenio1']['Limite'],
                    $nivel['ConcentracaoDeDioxidoDeNitrogenio2']['Media'], $nivel['ConcentracaoDeDioxidoDeNitrogenio2']['Limite'],
                    $nivel['TaxaDeEmissaoDeDioxidoDeNitrogenio']['Media'], $nivel['TaxaDeEmissaoDeDioxidoDeNitrogenio']['Limite'],
                    $nivel['Isocinetica']['Media'], $nivel['Isocinetica']['Limite'],
                    $nivel['NumeroDoBequer']['Branco'], $nivel['NumeroDoFiltro']['Branco'],
                    $nivel['MassaDoBequer']['Branco'], $nivel['MassaDoFiltro']['Branco'],
                    $nivel['MassaTotal']['Branco'], $nivel['Volume']['Branco'],
                    $nivel['ConcentracaoDeMaterialParticulado3']['Branco'], $nivel['IncertezaDeMedicaoDeMaterialParticulado']['Branco']
                );
                $stmt->execute();
                $stmt->close();
            }

            // Inserir os dados de Coleta
            if (isset($data['Resultados']) && is_array($data['Resultados'])) {
                foreach ($data['Resultados'] as $coletaKey => $coletaData) {
                    if (preg_match('/Coleta(\d+)/', $coletaKey, $matches)) {
                        $numeroColeta = $matches[1];

                        $stmt = $conn->prepare("
                            INSERT INTO coleta 
                            (perfil_chamine_id, numero_coleta, data_da_coleta, horario_inicial, horario_final, temperatura, umidade, o2, taxa_de_emissao_co2, taxa_de_emissao_o2, taxa_de_emissao_co, taxa_de_emissao_n2, velocidade, vazao1, vazao2, concentracao_de_material_particulado1, concentracao_de_material_particulado2, taxa_de_emissao_de_material_particulado, concentracao_de_monoxido_de_carbono1, concentracao_de_monoxido_de_carbono2, concentracao_de_dioxido_de_nitrogenio1, concentracao_de_dioxido_de_nitrogenio2, taxa_de_emissao_de_dioxido_de_nitrogenio, isocinetica, numero_do_bequer, numero_do_filtro, massa_do_bequer, massa_do_filtro, massa_total, volume, concentracao_de_material_particulado3, incerteza_de_medicao_de_material_particulado) 
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        ");
                        $stmt->bind_param("iissssssssssssssssssssssssssssss",
                            $perfilChamineId, 
                            $numeroColeta, 
                            $coletaData['DataDaColeta']['Valor'], $coletaData['HorarioInicial']['Valor'], $coletaData['Horariofinal']['Valor'], 
                            $coletaData['Temperatura']['Valor'], $coletaData['Umidade']['Valor'], $coletaData['O2']['Valor'], 
                            $coletaData['TaxaDeEmissaoCO2']['Valor'], $coletaData['TaxaDeEmissaoO2']['Valor'], 
                            $coletaData['TaxaDeEmissaoCO']['Valor'], $coletaData['TaxaDeEmissaoN2']['Valor'],
                            $coletaData['Velocidade']['Valor'], $coletaData['Vazao1']['Valor'], $coletaData['Vazao2']['Valor'],
                            $coletaData['ConcentracaoDeMaterialParticulado1']['Valor'], $coletaData['ConcentracaoDeMaterialParticulado2']['Valor'],
                            $coletaData['TaxaDeEmissaoDeMaterialParticulado']['Valor'], $coletaData['ConcentracaoDeMonoxidoDeCarbono1']['Valor'],
                            $coletaData['ConcentracaoDeMonoxidoDeCarbono2']['Valor'], $coletaData['ConcentracaoDeDioxidoDeNitrogenio1']['Valor'],
                            $coletaData['ConcentracaoDeDioxidoDeNitrogenio2']['Valor'], $coletaData['TaxaDeEmissaoDeDioxidoDeNitrogenio']['Valor'],
                            $coletaData['Isocinetica']['Valor'], $coletaData['NumeroDoBequer']['Valor'], $coletaData['NumeroDoFiltro']['Valor'],
                            $coletaData['MassaDoBequer']['Valor'], $coletaData['MassaDoFiltro']['Valor'], $coletaData['MassaTotal']['Valor'],
                            $coletaData['Volume']['Valor'], $coletaData['ConcentracaoDeMaterialParticulado3']['Valor'],
                            $coletaData['IncertezaDeMedicaoDeMaterialParticulado']['Valor']
                        );
                        $stmt->execute();
                        $stmt->close();
                    }
                }
            }

            $conn->commit(); // Confirma todas as inserções

            echo json_encode([
                'status' => 'success',
                'message' => 'Data inserted successfully',
            ]);
        } catch (Exception $e) {
            $conn->rollback(); // Desfaz qualquer alteração se der erro

            echo json_encode([
                'status' => 'error',
                'message' => 'Error inserting data: ' . $e->getMessage()
            ]);
        }
    } else {
        echo json_encode([
            'status' => 'error',
            'message' => 'Invalid JSON data'
        ]);
    }
}

$conn->close();
?>
