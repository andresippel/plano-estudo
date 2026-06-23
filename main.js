window.dadosDoCurso = null;
window.disciplinasAprovadas = [];
window.disciplinasMatriculadas = [];
window.chOptCursada = 0;
window.chTotalCursada = 0;
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
                alert("Não há dados para exportar!");
                return;
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
            .catch(error => console.error("Erro ao carregar o JSON do curso:", error));
    }
});

function atualizarInterfaceCurso() {
    const el = document.getElementById('cursoAtual');
    if (el && window.dadosDoCurso && window.dadosDoCurso.curso) {
        el.textContent = `Grade Ativa: ${window.dadosDoCurso.curso}`;
    }
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

            // Insere os dados unidos no novo campo único na tela
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
            window.disciplinasAprovadas = []; 
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

                const aprovado = strSituacao === "APROVADO" || strSituacao === "APROVEITADA POR EQUIVALÊNCIA" || strSituacao === "DISPENSADO DE CURSAR" || strSituacao.startsWith("APR") || strSituacao.startsWith("DISP");
                const matriculado = strSituacao === "MATRICULADO" || strSituacao.startsWith("MAT");

                if (matriculado) {
                    window.disciplinasMatriculadas.push({ nome: nomeDisciplina, ch: cargaHoraria });
                }

                if (aprovado && !isNaN(cargaHoraria)) {
                    if (strTipo === "OBR" || strTipo.includes("OBR")) {
                        somaCargaHorariaOBR += cargaHoraria;
                        window.disciplinasAprovadas.push(normalizarNome(nomeDisciplina));
                    } else if (strTipo === "OPT" || strTipo.includes("OPT")) {
                        somaCargaHorariaOPT += cargaHoraria;
                    }
                }
            });

            window.chOptCursada = somaCargaHorariaOPT;
            window.chTotalCursada = somaCargaHorariaOBR + somaCargaHorariaOPT;

            const elCH = document.getElementById('cargaHoraria');
            if (elCH) elCH.value = somaCargaHorariaOBR;
            
            calcularSemestre();

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

    let cargaHorariaEstudante = parseInt(elCH.value);

    if (isNaN(cargaHorariaEstudante) || cargaHorariaEstudante < 0) {
        resultado.style.display = "block";
        resultado.textContent = "Por favor, digite uma carga horária válida!";
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

        if (chAcumulada > cargaHorariaEstudante) {
            let diferenca = chAcumulada - cargaHorariaEstudante;
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
    resultado.textContent = `O estudante está no ${semestreEnquadrado}`;
    
    if (acoesPlano) acoesPlano.style.display = "block";
    if (typeof window.atualizarPreviewDocumento === "function") {
        window.atualizarPreviewDocumento();
    }
}