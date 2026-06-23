// ---------------------------------------------------
// MÓDULO: GERAÇÃO DO PLANO DE ESTUDOS
// ---------------------------------------------------

document.addEventListener('DOMContentLoaded', function() {
    const btnCopiarSEI = document.getElementById('btnCopiarSEI');
    const btnBaixarWord = document.getElementById('btnBaixarWord');
    const alunoInfo = document.getElementById('alunoInfo'); 

    if (btnCopiarSEI) {
        btnCopiarSEI.addEventListener('click', function() {
            const html = gerarHTMLPlanoEstudos();
            navigator.clipboard.writeText(html).then(() => {
                alert("Código copiado com sucesso! No SEI, clique em '< >' (Código Fonte) e cole.");
            }).catch(err => {
                alert("Erro ao copiar automaticamente.");
                console.error('Erro ao copiar código: ', err);
            });
        });
    }

    if (btnBaixarWord) {
        btnBaixarWord.addEventListener('click', function() {
            const htmlTabela = gerarHTMLPlanoEstudos();
            const htmlParaWord = `
                <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
                <head>
                    <meta charset='utf-8'>
                    <title>Plano de Estudos</title>
                </head>
                <body>
                    ${htmlTabela}
                </body>
                </html>
            `;
            const blob = new Blob(['\ufeff', htmlParaWord], { type: 'application/msword' });
            const url = URL.createObjectURL(blob);
            
            const link = document.createElement('a');
            link.href = url;
            link.download = "Plano_de_Estudos_FAODO.doc";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        });
    }

    if (alunoInfo) {
        alunoInfo.addEventListener('input', window.atualizarPreviewDocumento);
    }
    
    document.getElementById('listaAproveitamento').addEventListener('change', function(e) {
        if(e.target.classList.contains('chk-aproveitamento')) {
            if (typeof calcularSemestre === "function") calcularSemestre();
        }
    });
});

window.atualizarPreviewDocumento = function() {
    const acoesPlano = document.getElementById('acoesPlano');
    const previewPlano = document.getElementById('previewPlano');
    if (acoesPlano && acoesPlano.style.display === "block" && previewPlano) {
        previewPlano.innerHTML = gerarHTMLPlanoEstudos();
    }
};

function gerarHTMLPlanoEstudos() {
    const elAlunoInfo = document.getElementById('alunoInfo');
    let infoStr = elAlunoInfo ? elAlunoInfo.value.trim() : "";
    
    let nome = "Não informado";
    let rga = "Não informado";

    if (infoStr !== "") {
        let separadorIndex = infoStr.indexOf(" - ");
        if (separadorIndex !== -1) {
            rga = infoStr.substring(0, separadorIndex).trim();
            nome = infoStr.substring(separadorIndex + 3).trim();
        } else {
            nome = infoStr;
        }
    }

    const enquadramento = window.ultimoEnquadramento || "Não calculado";

    let html = `
<p class="Item_Nivel1" style="font-weight: bold;">Identificação do Acadêmico:</p>
<p class="Texto_Justificado">Nome: ${nome}</p>
<p class="Texto_Justificado">RGA na UFMS: ${rga}</p>
<p class="Texto_Justificado">Enquadramento: ${enquadramento}</p>
<p class="Texto_Justificado">&nbsp;</p>

<p class="Item_Nivel1" style="font-weight: bold;">Plano de estudos das disciplinas:</p>
<!-- TABELA 1: Largura travada em 85% para alinhar com os textos do SEI -->
<table class="table" style="border-collapse:collapse;margin-left:auto;margin-right:auto;width:85%;" border="1" cellpadding="5">
    <thead>
        <tr>
            <th style="background-color:#cccccc;"><p class="Tabela_Texto_Centralizado">Disciplinas a serem cursadas no semestre atual</p></th>
            <th style="background-color:#cccccc;"><p class="Tabela_Texto_Centralizado">Indicação da turma prática</p></th>
        </tr>
    </thead>
    <tbody>`;

    if (window.disciplinasMatriculadas && window.disciplinasMatriculadas.length > 0) {
        window.disciplinasMatriculadas.forEach(d => {
            html += `<tr><td><p class="Tabela_Texto_Alinhado_Esquerda">${d.nome}</p></td><td><p class="Tabela_Texto_Centralizado">&nbsp;</p></td></tr>`;
        });
    } else {
        html += `
        <tr><td><p class="Tabela_Texto_Alinhado_Esquerda">&nbsp;</p></td><td><p class="Tabela_Texto_Centralizado">&nbsp;</p></td></tr>
        <tr><td><p class="Tabela_Texto_Alinhado_Esquerda">&nbsp;</p></td><td><p class="Tabela_Texto_Centralizado">&nbsp;</p></td></tr>
        <tr><td><p class="Tabela_Texto_Alinhado_Esquerda">&nbsp;</p></td><td><p class="Tabela_Texto_Centralizado">&nbsp;</p></td></tr>
        <tr><td><p class="Tabela_Texto_Alinhado_Esquerda">&nbsp;</p></td><td><p class="Tabela_Texto_Centralizado">&nbsp;</p></td></tr>`;
    }

    html += `
    </tbody>
</table>
<p class="Texto_Alinhado_Esquerda">&nbsp;</p>

<!-- TABELA 2: Largura travada em 85% para alinhar com os textos do SEI -->
<table class="table" style="border-collapse:collapse;margin-left:auto;margin-right:auto;width:85%;" border="1" cellpadding="5">
    <thead>
        <tr>
            <th style="background-color:#cccccc;width:60%;"><p class="Tabela_Texto_Centralizado">Disciplina a serem cursadas em semestres posteriores</p></th>
            <th style="background-color:#cccccc;width:20%;"><p class="Tabela_Texto_Centralizado">Indicação do Semestre:</p></th>
            <th style="background-color:#cccccc;width:20%;"><p class="Tabela_Texto_Centralizado">Carga horária</p></th>
        </tr>
    </thead>
    <tbody>`;

    let grandTotalCH = 0; 

    if (window.dadosDoCurso && window.dadosDoCurso.semestres) {
        window.dadosDoCurso.semestres.forEach(semestre => {
            let disciplinasPendentes = semestre.disciplinas.filter(d => {
                return !window.disciplinasAprovadas.includes(normalizarNome(d.nome));
            });

            if (disciplinasPendentes.length > 0) {
                html += `
        <tr>
            <td style="width:60%;"><p class="Tabela_Texto_Alinhado_Esquerda"><strong>${semestre.numero}° semestre</strong></p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado">&nbsp;</p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado">&nbsp;</p></td>
        </tr>`;
                
                let totalSemestreCH = 0;
                disciplinasPendentes.forEach(d => {
                    totalSemestreCH += d.ch;
                    grandTotalCH += d.ch; 
                    html += `
        <tr>
            <td style="width:60%;"><p class="Tabela_Texto_Alinhado_Esquerda">${d.nome}</p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado">&nbsp;</p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado">${d.ch}</p></td>
        </tr>`;
                });

                html += `
        <tr>
            <td style="width:60%;"><p class="Tabela_Texto_Alinhado_Esquerda"><strong>TOTAL</strong></p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado">&nbsp;</p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado"><strong>${totalSemestreCH}</strong></p></td>
        </tr>
        <tr>
            <td style="width:60%;"><p class="Tabela_Texto_Alinhado_Esquerda">&nbsp;</p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado">&nbsp;</p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado">&nbsp;</p></td>
        </tr>`;
            }
        });
    }

    html += `
        <tr>
            <td style="width:60%;"><p class="Tabela_Texto_Alinhado_Esquerda">&nbsp;</p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado"><strong>TOTAL DO PLANO</strong></p></td>
            <td style="width:20%;"><p class="Tabela_Texto_Centralizado"><strong>${grandTotalCH}</strong></p></td>
        </tr>`;

    html += `
    </tbody>
</table>`;

    const reqOptativas = (window.dadosDoCurso && window.dadosDoCurso.ch_optativas_exigidas) ? window.dadosDoCurso.ch_optativas_exigidas : 60;
    const reqExtensao = (window.dadosDoCurso && window.dadosDoCurso.ch_extensao_exigida) ? window.dadosDoCurso.ch_extensao_exigida : 409;
    const reqTotalCurso = (window.dadosDoCurso && window.dadosDoCurso.ch_total_curso) ? window.dadosDoCurso.ch_total_curso : 4090;

    const optCursada = window.chOptCursada || 0;
    const optACursar = Math.max(0, reqOptativas - optCursada);
    
    // Cálculo rigoroso com as variáveis passadas pelo main.js
    const totalCursada = window.chTotalCursada || 0;
    const aCursarTotal = Math.max(0, reqTotalCurso - totalCursada);

    const aproveitadasExcel = window.chAproveitamentosExcel || 0;
    const aproveitadasManual = window.chAproveitadaManual || 0;
    const totalAproveitadas = aproveitadasExcel + aproveitadasManual;

    html += `
<p class="Texto_Alinhado_Esquerda">&nbsp;</p>
<p class="Item_Nivel1" style="font-weight: bold;">3. CARGA HORÁRIA DE DISCIPLINAS OPTATIVAS NECESSÁRIAS:</p>
<p class="Texto_Justificado">Carga horária total de disciplinas optativa cursada: ${optCursada} horas</p>
<p class="Texto_Justificado">Carga horária total de disciplinas optativa a cursar: ${optACursar} horas</p>
<p class="Texto_Justificado">&nbsp;</p>

<p class="Item_Nivel1" style="font-weight: bold;">4. CARGA HORÁRIA EM ATIVIDADES DE EXTENSÃO NECESSÁRIAS:</p>
<p class="Texto_Justificado">Carga horária em atividades de extensão cursadas: 0 horas</p>
<p class="Texto_Justificado">Carga horária em atividades de extensão a cursar: ${reqExtensao} horas</p>
<p class="Texto_Justificado">&nbsp;</p>

<p class="Item_Nivel1" style="font-weight: bold;">5. CARGA HORÁRIA DE DISCIPLINAS NECESSÁRIAS PARA A INTEGRALIZAÇÃO CURRICULAR:</p>
<p class="Texto_Justificado">Carga horária total de disciplinas obrigatórias aproveitadas e dispensadas: ${totalAproveitadas} horas</p>
<p class="Texto_Justificado">Carga horária total de disciplinas a cursar: ${aCursarTotal} horas</p>
<p class="Texto_Justificado">Carga horaria total do plano de estudos: ${grandTotalCH} horas</p>

<p class="Texto_Justificado">&nbsp;</p>
<p class="Texto_Justificado">&nbsp;</p>
<p class="Texto_Justificado">&nbsp;</p>
`;

    return html;
}