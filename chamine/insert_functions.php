<?php
function insertPerfilChamine($pdo, $perfil) {
    $sql = "INSERT INTO perfil_chamine (projeto, fonte, data, equipe, empresa, cidade, processo, parametros, montante, jusante, diametro1, diametro2, diametro3, diametro_medio, jdc, mdc, area_chamine, comprimento, diametro_equivalente, numero_de_pontos, numero_de_eixos, numero_de_pontos_por_eixo, tempo_total_coleta, tempo_por_ponto) 
            VALUES (:projeto, :fonte, :data, :equipe, :empresa, :cidade, :processo, :parametros, :montante, :jusante, :diametro1, :diametro2, :diametro3, :diametro_medio, :jdc, :mdc, :area_chamine, :comprimento, :diametro_equivalente, :numero_de_pontos, :numero_de_eixos, :numero_de_pontos_por_eixo, :tempo_total_coleta, :tempo_por_ponto)";
    $stmt = $pdo->prepare($sql);
    $stmt->execute([
        ':projeto' => $perfil['Projeto'],
        ':fonte' => $perfil['Fonte'],
        ':data' => $perfil['Data'],
        ':equipe' => $perfil['Equipe'],
        ':empresa' => $perfil['Empresa'],
        ':cidade' => $perfil['Cidade'],
        ':processo' => $perfil['Processo'],
        ':parametros' => $perfil['Parametros'],
        ':montante' => $perfil['Montante'],
        ':jusante' => $perfil['Jusante'],
        ':diametro1' => $perfil['Diametro1'],
        ':diametro2' => $perfil['Diametro2'],
        ':diametro3' => $perfil['Diametro3'],
        ':diametro_medio' => $perfil['DiametroMedio'],
        ':jdc' => $perfil['Jdc'],
        ':mdc' => $perfil['Mdc'],
        ':area_chamine' => $perfil['AreaChamine'],
        ':comprimento' => $perfil['Comprimento'],
        ':diametro_equivalente' => $perfil['Diametroequivalente'],
        ':numero_de_pontos' => $perfil['NumeroDePontos'],
        ':numero_de_eixos' => $perfil['NumeroDeEixos'],
        ':numero_de_pontos_por_eixo' => $perfil['NumeroDePontosPorEixo'],
        ':tempo_total_coleta' => $perfil['TempoTotalColeta'],
        ':tempo_por_ponto' => $perfil['TempoporPonto']
    ]);
    return $pdo->lastInsertId();
}

function insertColeta($pdo, $perfilChamineId, $coleta) {
    $sql = "INSERT INTO coleta (perfil_chamine_id, numero_coleta, data_da_coleta, horario_inicial, horario_final, temperatura, umidade, o2, taxa_de_emissao_co2, taxa_de_emissao_o2, taxa_de_emissao_co, taxa_de_emissao_n2, velocidade, vazao1, vazao2, concentracao_de_material_particulado1, concentracao_de_material_particulado2, taxa_de_emissao_de_material_particulado, concentracao_de_monoxido_de_carbono1, concentracao_de_monoxido_de_carbono2, concentracao_de_dioxido_de_nitrogenio1, concentracao_de_dioxido_de_nitrogenio2, taxa_de_emissao_de_dioxido_de_nitrogenio, isocinetica, numero_do_bequer, numero_do_filtro, massa_do_bequer, massa_do_filtro, massa_total, volume, concentracao_de_material_particulado3, incerteza_de_medicao_de_material_particulado) 
            VALUES (:perfil_chamine_id, :numero_coleta, :data_da_coleta, :horario_inicial, :horario_final, :temperatura, :umidade, :o2, :taxa_de_emissao_co2, :taxa_de_emissao_o2, :taxa_de_emissao_co, :taxa_de_emissao_n2, :velocidade, :vazao1, :vazao2, :concentracao_de_material_particulado1, :concentracao_de_material_particulado2, :taxa_de_emissao_de_material_particulado, :concentracao_de_monoxido_de_carbono1, :concentracao_de_monoxido_de_carbono2, :concentracao_de_dioxido_de_nitrogenio1, :concentracao_de_dioxido_de_nitrogenio2, :taxa_de_emissao_de_dioxido_de_nitrogenio, :isocinetica, :numero_do_bequer, :numero_do_filtro, :massa_do_bequer, :massa_do_filtro, :massa_total, :volume, :concentracao_de_material_particulado3, :incerteza_de_medicao_de_material_particulado)";
    $stmt = $pdo->prepare($sql);
    $stmt->execute([
        ':perfil_chamine_id' => $perfilChamineId,
        ':numero_coleta' => $coleta['NumeroColeta'],
        ':data_da_coleta' => $coleta['DataDaColeta']['Valor'],
        ':horario_inicial' => $coleta['HorarioInicial']['Valor'],
        ':horario_final' => $coleta['Horariofinal']['Valor'],
        ':temperatura' => $coleta['Temperatura']['Valor'],
        ':umidade' => $coleta['Umidade']['Valor'],
        ':o2' => $coleta['O2']['Valor'],
        ':taxa_de_emissao_co2' => $coleta['TaxaDeEmissaoCO2']['Valor'],
        ':taxa_de_emissao_o2' => $coleta['TaxaDeEmissaoO2']['Valor'],
        ':taxa_de_emissao_co' => $coleta['TaxaDeEmissaoCO']['Valor'],
        ':taxa_de_emissao_n2' => $coleta['TaxaDeEmissaoN2']['Valor'],
        ':velocidade' => $coleta['Velocidade']['Valor'],
        ':vazao1' => $coleta['Vazao1']['Valor'],
        ':vazao2' => $coleta['Vazao2']['Valor'],
        ':concentracao_de_material_particulado1' => $coleta['ConcentracaoDeMaterialParticulado1']['Valor'],
        ':concentracao_de_material_particulado2' => $coleta['ConcentracaoDeMaterialParticulado2']['Valor'],
        ':taxa_de_emissao_de_material_particulado' => $coleta['TaxaDeEmissaoDeMaterialParticulado']['Valor'],
        ':concentracao_de_monoxido_de_carbono1' => $coleta['ConcentracaoDeMonoxidoDeCarbono1']['Valor'],
        ':concentracao_de_monoxido_de_carbono2' => $coleta['ConcentracaoDeMonoxidoDeCarbono2']['Valor'],
        ':concentracao_de_dioxido_de_nitrogenio1' => $coleta['ConcentracaoDeDioxidoDeNitrogenio1']['Valor'],
        ':concentracao_de_dioxido_de_nitrogenio2' => $coleta['ConcentracaoDeDioxidoDeNitrogenio2']['Valor'],
        ':taxa_de_emissao_de_dioxido_de_nitrogenio' => $coleta['TaxaDeEmissaoDeDioxidoDeNitrogenio']['Valor'],
        ':isocinetica' => $coleta['Isocinetica']['Valor'],
        ':numero_do_bequer' => $coleta['NumeroDoBequer']['Valor'],
        ':numero_do_filtro' => $coleta['NumeroDoFiltro']['Valor'],
        ':massa_do_bequer' => $coleta['MassaDoBequer']['Valor'],
        ':massa_do_filtro' => $coleta['MassaDoFiltro']['Valor'],
        ':massa_total' => $coleta['MassaTotal']['Valor'],
        ':volume' => $coleta['Volume']['Valor'],
        ':concentracao_de_material_particulado3' => $coleta['ConcentracaoDeMaterialParticulado3']['Valor'],
        ':incerteza_de_medicao_de_material_particulado' => $coleta['IncertezaDeMedicaoDeMaterialParticulado']['Valor']
    ]);
}
?>