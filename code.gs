const p_nom = "Página1";
const cel_inicial = "B1";
const cel_rot = "A1";

function onEdit(e) {
  const range = e.range;
  const sheet = e.source.getSheetByName(p_nom);
  const valor_escolhido = e.value;
  const linha_acao = range.getRow();
  const coluna_acao = range.getColumn();

  if (!sheet || sheet.getName() !== p_nom) return;

  // Rotação da Lista
  if (range.getA1Notation() === cel_rot && valor_escolhido === 'TRUE') {
    range.setValue(false); // Reseta o checkbox

    const range_inicionomes = sheet.getRange(cel_inicial);
    const coluna_nomes = range_inicionomes.getColumn();
    const linha_inicialnomes = range_inicionomes.getRow();
    const ulp = encontrarUltimaLinhaComConteudo(sheet, coluna_nomes);

    if (ulp < linha_inicialnomes) return; // Lista vazia
    if (ulp === linha_inicialnomes) {
      sheet.getRange(linha_inicialnomes, coluna_nomes).clearContent();
      return;
    }

    const source_range = sheet.getRange(linha_inicialnomes + 1, coluna_nomes, ulp - linha_inicialnomes);
    const dest_range = sheet.getRange(linha_inicialnomes, coluna_nomes, ulp - linha_inicialnomes);
    source_range.copyTo(dest_range);
    sheet.getRange(ulp, coluna_nomes).clearContent();
    return;
  }

  // Painel de Ações
  const range_inicionomes = sheet.getRange(cel_inicial);
  const coluna_nomes = range_inicionomes.getColumn();
  const linha_inicialnomes = range_inicionomes.getRow();
  const coluna_acoes = coluna_nomes + 1;

  // Verifica se a edição foi na coluna de ações
  if (coluna_acao === coluna_acoes && linha_acao >= linha_inicialnomes && valor_escolhido) {
    range.clearContent(); // Limpa a célula da ação para poder ser usada de novo

    const ulp = encontrarUltimaLinhaComConteudo(sheet, coluna_nomes);
    
    switch (valor_escolhido) {
      case "Remover Nome":
        if (linha_acao < ulp) {
          const source_range = sheet.getRange(linha_acao + 1, coluna_nomes, ulp - linha_acao);
          const dest_range = sheet.getRange(linha_acao, coluna_nomes, ulp - linha_acao);
          source_range.copyTo(dest_range);
        }
        // Limpa o conteúdo da última linha, preserva a chip
        sheet.getRange(ulp, coluna_nomes).clearContent();
        break;

      case "Último Lugar":
        const range_nome = sheet.getRange(linha_acao, coluna_nomes);
        const valor_nome = range_nome.getRichTextValue();
        
        if (!valor_nome.getText() || linha_acao === ulp) {
          return;
        }

        const move_block = sheet.getRange(linha_acao, coluna_nomes, ulp - linha_acao + 1);
        const valor_bloco = move_block.getRichTextValues();

        // Remove o orador da sua posição original no array.
        const nome_movido = valor_bloco.shift();
        // Adiciona o orador no final do array.
        valor_bloco.push(nome_movido);

        // Escreve o array reordenado de volta na planilha.
        move_block.setRichTextValues(valor_bloco);
        break;
    }
  }
}

// Encontra ultima linha com Conteudo
function encontrarUltimaLinhaComConteudo(sheet, coluna) {
    const todos_valores = sheet.getRange(1, coluna, sheet.getMaxRows()).getValues();
    let ultima_linha = 0;
    for (let i = todos_valores.length - 1; i >= 0; i--) {
        if (todos_valores[i][0] !== "") {
            ultima_linha = i + 1;
            break;
        }
    }
    return ultima_linha;
}
