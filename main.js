window.dadosDoCurso = null;
window.disciplinasAprovadasExcel = [];
window.disciplinasAprovadas = [];
window.disciplinasMatriculadas = [];
window.chOptCursada = 0;
window.chNfcCursada = 0; // Nova variável para o Núcleo de Formação Cidadã
window.chTotalCursada = 0;
window.chAproveitamentosExcel = 0; 
window.chAproveitadaManual = 0;    
window.ultimoEnquadramento = "";
window.nomeDoAlunoPlanilha = ""; 
window.rgaDoAlunoPlanilha = "";  

function normalizarNome(nome) {
    if (!nome) return "";
    return String(nome).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
}

function configurarDragAndDrop(idDropZone, idInput) {
    const dropZone = document.getElementById(idDropZone);
    const inputElement = document.getElementById(idInput);
    if (!dropZone || !inputElement) return;

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    ['dragleave', 'dragend'].forEach(type => {
        dropZone.addEventListener(type, () => dropZone.classList.remove('dragover'));
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            inputElement.files = e.dataTransfer.files;
            inputElement.dispatchEvent(new Event('change')); 
        }
    });
}

document.addEventListener('DOMContentLoaded', function() {
    const btnCalcular = document.getElementById('btnCalcular');
    const inputHistorico = document.getElementById('inputHistorico');
    const inputImportar = document.getElementById('inputImportar');
    const btnExportar = document.getElementById('btnExportar');

    configurarDragAndDrop('dropZoneHistorico', 'inputHistorico');
    configurarDragAndDrop('dropZoneImportar', 'inputImportar');

    if (btnCalcular) btnCalcular.addEventListener('click', calcularSemestre);
    if (inputHistorico) inputHistorico.addEventListener('change', processarHistorico);

    if (inputImportar) {
        inputImportar.addEventListener('change', function(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const jsonImportado = JSON.parse(e.target.result);
                    window.dadosDoCurso = jsonImportado;
                    localStorage.setItem('gradeCurso_v3', JSON.stringify(jsonImportado));
                    atualizarInterfaceCurso();
                    alert(`Grade do curso de ${jsonImportado.curso} importada com sucesso!`);
                } catch (erro) {
                    alert("Erro ao ler o arquivo JSON. Verifique se o formato está correto.");
                }
                event.target.value = ''; 
            };
            reader.readAsText(file);
        });
    }

    if (btnExportar) {
        btnExportar.addEventListener('click', function() {
            if (!window.dadosDoCurso) {
                alert("Não há dados para exportar!"); return;
            }
            const textoJSON = JSON.stringify(window.dadosDoCurso, null, 2);
            const blob = new Blob([textoJSON], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `grade_${window.dadosDoCurso.curso}.json`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        });
    }

    const dadosSalvos = localStorage.getItem('gradeCurso_v3');
    if (dadosSalvos) {
        window.dadosDoCurso = JSON.parse(dadosSalvos);
        atualizarInterfaceCurso();
    } else {
        fetch('dados_curso.json')
            .then(response => response.json())
            .then(data => {
                window.dadosDoCurso = data;
                localStorage.setItem('gradeCurso_v3', JSON.stringify(data));
                atualizarInterfaceCurso();
            })
            .catch(error => console.error("Erro ao carregar o JSON:", error));
    }
});

function atualizarInterfaceCurso() {
    const el = document.getElementById('cursoAtual');
    if (el && window.dadosDoCurso && window.dadosDoCurso.curso) {
        el.textContent = `Grade Ativa: ${window.dadosDoCurso.curso}`;
    }
    renderizarListaAproveitamento();
}

function renderizarListaAproveitamento() {
    const container = document.getElementById('listaAproveitamento');
    if (!container || !window.dadosDoCurso) return;
    
    let html = '';
    let temDisciplinaDisponivel = false;

    window.dadosDoCurso.semestres.forEach(sem => {
        let disciplinasDisponiveis = sem.disciplinas.filter(d => {
            return !window.disciplinasAprovadasExcel.includes(normalizarNome(d.nome));
        });

        if (disciplinasDisponiveis.length > 0) {
            temDisciplinaDisponivel = true;
            html += `<div class="semester-group">${sem.numero}º Semestre</div>`;
            disciplinasDisponiveis.forEach(d => {
                const id = `chk_${normalizarNome(d.nome).replace(/[^a-zA-Z0-9]/g, '_')}`;
                html += `
                <label class="checkbox-item" for="${id}">
                    <input type="checkbox" id="${id}" class="chk-aproveitamento" value="${d.nome}" data-ch="${d.ch}">
                    ${d.nome} (${d.ch}h)
                </label>`;
            });
        }
    });

    if (!temDisciplinaDisponivel) {
        html = '<p style="color: #888; font-style: italic;">Todas as disciplinas já constam no histórico como cursadas/aproveitadas.</p>';
    }

    container.innerHTML = html;
}

function processarHistorico(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            const matrizRaw = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            const nomeEstudante = matrizRaw[0] && matrizRaw[0][0] ? String(matrizRaw[0][0]).trim() : "";
            const rgaEstudante = matrizRaw[1] && matrizRaw[1][0] ? String(matrizRaw[1][0]).trim() : "";

            window.nomeDoAlunoPlanilha = nomeEstudante;
            window.rgaDoAlunoPlanilha = rgaEstudante;

            const elAlunoInfo = document.getElementById('alunoInfo');
            if (elAlunoInfo) {
                if (rgaEstudante && nomeEstudante) {
                    elAlunoInfo.value = `${rgaEstudante} - ${nomeEstudante}`;
                } else if (nomeEstudante) {
                    elAlunoInfo.value = nomeEstudante;
                } else if (rgaEstudante) {
                    elAlunoInfo.value = rgaEstudante;
                }
            }

            const linhas = XLSX.utils.sheet_to_json(worksheet, { range: 2 });
            
            let somaCargaHorariaOBR = 0;
            let somaCargaHorariaOPT = 0;
            let somaCargaHorariaNFC = 0;
            let somaAproveitamentosExcel = 0;
            
            window.disciplinasAprovadasExcel = []; 
            window.disciplinasMatriculadas = [];

            linhas.forEach(linha => {
                const nomeDisciplina = linha["Nome da Disciplina/CCND"];
                const tipo = linha["Tipo"]; 
                const situacao = linha["Situação"];
                const chStr = linha["C.H."];

                if (!tipo || !situacao || chStr === undefined || !nomeDisciplina) return;

                const cargaHoraria = parseFloat(chStr);
                const strTipo = String(tipo).trim().toUpperCase();
                const strSituacao = String(situacao).trim().toUpperCase();

                // Identifica se a disciplina foi dispensada (cobre "Dispensado por Análise de Currículo" e "Aproveitada...")
                const dispensado = strSituacao.includes("APROVEITAD") || strSituacao.includes("DISPENSA") || strSituacao.startsWith("DISP");
                const aprovado = strSituacao === "APROVADO" || strSituacao.startsWith("APR") || dispensado;
                const matriculado = strSituacao === "MATRICULADO" || strSituacao.startsWith("MAT");

                if (matriculado) {
                    window.disciplinasMatriculadas.push({ nome: nomeDisciplina, ch: cargaHoraria });
                }

                if (aprovado && !isNaN(cargaHoraria)) {
                    if (strTipo === "OBR" || strTipo.includes("OBR")) {
                        somaCargaHorariaOBR += cargaHoraria;
                        window.disciplinasAprovadasExcel.push(normalizarNome(nomeDisciplina));
                        if (dispensado) somaAproveitamentosExcel += cargaHoraria;
                    } else if (strTipo === "OPT" || strTipo.includes("OPT")) {
                        somaCargaHorariaOPT += cargaHoraria;
                    } else if (strTipo === "NFC" || strTipo.includes("NFC")) {
                        somaCargaHorariaNFC += cargaHoraria;
                    }
                }
            });

            window.chOptCursada = somaCargaHorariaOPT;
            window.chNfcCursada = somaCargaHorariaNFC;
            window.chAproveitamentosExcel = somaAproveitamentosExcel;

            const elCH = document.getElementById('cargaHoraria');
            if (elCH) elCH.value = somaCargaHorariaOBR;
            
            renderizarListaAproveitamento();
            calcularSemestre();
            
            alert(`Histórico processado! Obrigatórias: ${somaCargaHorariaOBR}h | Optativas: ${somaCargaHorariaOPT}h.`);

        } catch (err) {
            alert("Erro ao processar a planilha. Verifique o formato do arquivo.");
            console.error(err);
        }
        event.target.value = '';
    };
    reader.readAsArrayBuffer(file);
}

function calcularSemestre() {
    const elCH = document.getElementById("cargaHoraria");
    const resultado = document.getElementById("resultado");
    const acoesPlano = document.getElementById('acoesPlano');
    if (!elCH || !resultado) return;

    let chBaseEstudante = parseInt(elCH.value) || 0;

    let chAproveitadaCheckbox = 0;
    let disciplinasAproveitadasCheck = [];
    
    document.querySelectorAll('.chk-aproveitamento:checked').forEach(chk => {
        chAproveitadaCheckbox += parseInt(chk.getAttribute('data-ch')) || 0;
        disciplinasAproveitadasCheck.push(normalizarNome(chk.value));
    });

    window.chAproveitadaManual = chAproveitadaCheckbox;
    
    window.disciplinasAprovadas = [...window.disciplinasAprovadasExcel, ...disciplinasAproveitadasCheck];

    let cargaHorariaEstudanteFinal = chBaseEstudante + chAproveitadaCheckbox;
    
    // Total cursado final agora inclui as OBR, OPT e NFC
    window.chTotalCursada = cargaHorariaEstudanteFinal + window.chOptCursada + window.chNfcCursada;

    if (isNaN(cargaHorariaEstudanteFinal) || cargaHorariaEstudanteFinal < 0) {
        resultado.style.display = "block";
        resultado.textContent = "Por favor, defina uma carga horária válida.";
        if (acoesPlano) acoesPlano.style.display = "none";
        return;
    }

    if (!window.dadosDoCurso) return;

    let chAcumulada = 0;
    let semestreEnquadrado = null;

    for (let i = 0; i < window.dadosDoCurso.semestres.length; i++) {
        let semestre = window.dadosDoCurso.semestres[i];
        let chTotalSemestre = semestre.disciplinas.reduce((soma, disciplina) => soma + disciplina.ch, 0);
        chAcumulada += chTotalSemestre;

        if (chAcumulada > cargaHorariaEstudanteFinal) {
            let diferenca = chAcumulada - cargaHorariaEstudanteFinal;
            if (diferenca <= 136) {
                let proximoSemestre = Math.min(i + 2, window.dadosDoCurso.semestres.length);
                semestreEnquadrado = `${proximoSemestre}º semestre`;
            } else {
                semestreEnquadrado = `${semestre.numero}º semestre`;
            }
            break;
        }
    }

    if (!semestreEnquadrado) semestreEnquadrado = "Formado (ou no último semestre)";
    
    window.ultimoEnquadramento = semestreEnquadrado;
    resultado.style.display = "block";
    resultado.textContent = `O estudante está no ${semestreEnquadrado} (CH Analisada: ${cargaHorariaEstudanteFinal}h)`;
    
    if (acoesPlano) acoesPlano.style.display = "block";
    if (typeof window.atualizarPreviewDocumento === "function") {
        window.atualizarPreviewDocumento();
    }
}