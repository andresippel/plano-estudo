# Sistema de Enquadramento e Plano de Estudos - UFMS

Ferramenta desenvolvida para automatizar a Análise de Histórico, calcular o enquadramento semestral (com base na IN 40/2019) e gerar o Plano de Estudos oficial para processos de aproveitamento no SEI (Sistema Eletrônico de Informações).

Desenvolvido originalmente para a Faculdade de Odontologia (FAODO), o sistema é **100% dinâmico** e pode ser facilmente adaptado para qualquer curso da Universidade Federal de Mato Grosso do Sul (UFMS).

## 🔒 Privacidade e Segurança (LGPD)
Este sistema funciona inteiramente no navegador do usuário (*Client-Side*). 
* Nenhum dado pessoal do acadêmico ou histórico de notas é enviado para servidores externos. 
* O cruzamento de dados ocorre na memória RAM do computador local, garantindo total conformidade com a Lei Geral de Proteção de Dados (LGPD).

---

## 🚀 Como usar o sistema no dia a dia

1. **Prepare o Arquivo:** Baixe a Análise de Histórico do acadêmico (em formato `.xlsx` ou `.csv`).
   * *Dica de produtividade:* Se você renomear o arquivo para o formato `RGA - Nome do Aluno.xlsx` (Ex: `2023.1234.000-0 - João da Silva.xlsx`), o sistema preencherá a identificação do acadêmico automaticamente!
2. **Faça o Upload:** Arraste e solte o arquivo na área pontilhada de **Processamento de Histórico**.
3. **Aproveitamento de Estudos:** Selecione aqui as disciplinas que foram aproveitadas pela Análise de Curriculo.
4. **Exporte para o SEI:** Role a tela até a prévia do documento. Clique em **Copiar para colar no SEI (<>)**. 
5. **Cole no SEI:** Crie o documento "Plano de Estudos" no SEI, e após a criação, no editor de texto do processo no SEI, clique no botão `< >` (Código Fonte) apague o conteúdo que estava originalmente e cole (`Ctrl + V`) o conteúdo copiado. O documento será gerado com as tabelas na largura exata e com as horas perfeitamente calculadas.
    * *Detalhe:* o cálculo dos itens 4 e 5 não são calculados corretamente, sendo necessário a correção manual.

---

## ⚙️ Como adaptar este sistema para o seu Curso

Se você é de outra coordenação (Ex: Administração, Engenharia, Direito) e deseja utilizar este sistema, não é necessário alterar nenhuma linha de código de programação. Toda a inteligência da ferramenta consome um arquivo de texto chamado `dados_curso.json`.

Siga os passos abaixo para migrar o sistema para a sua grade curricular:

### Passo 1: Entenda a estrutura do JSON
O arquivo `.json` dita as regras do seu curso (cargas horárias totais exigidas) e lista as disciplinas separadas por semestre. Ele deve seguir rigorosamente o formato abaixo:

```json
{
  "curso": "Nome do Seu Curso",
  "ch_total_curso": 4090,
  "ch_optativas_exigidas": 60,
  "ch_extensao_exigida": 409,
  "semestres": [
    {
      "numero": 1,
      "disciplinas": [
        { "nome": "Nome da Disciplina 1", "ch": 60 },
        { "nome": "Nome da Disciplina 2", "ch": 45 }
      ]
    },
    {
      "numero": 2,
      "disciplinas": [
        { "nome": "Nome da Disciplina 3", "ch": 90 }
      ]
    }
  ]
}
```

#### Dicionário de Variáveis:
* `"curso"`: O nome do curso que aparecerá no topo do site.
* `"ch_total_curso"`: A carga horária total para a integralização (usada para o cálculo do Item 5).
* `"ch_optativas_exigidas"`: O total de horas optativas que o seu PPC exige (usado para o cálculo do Item 3).
* `"ch_extensao_exigida"`: O total de horas de extensão que o seu PPC exige (usado para o cálculo do Item 4).
* `"semestres"`: Uma lista que engloba o número do semestre e as disciplinas pertencentes a ele. Cuidado com erros de digitação nos nomes, eles devem ser idênticos aos cadastrados no Siscad.

### Passo 2: Importe no Sistema
1. Crie o seu arquivo através do Bloco de Notas ou de editores de código (como VS Code) e salve-o com a extensão `.json`.
2. Acesse o site do sistema.
3. Role até a área **Gerenciar Grade do Curso**.
4. Arraste e solte o seu arquivo JSON na caixa de importação.
5. Pronto! O navegador gravará a sua grade curricular e o sistema estará pronto para calcular o enquadramento dos seus alunos.

> **Nota:** Para backup ou compartilhamento, você pode clicar em "Exportar Grade Atual (JSON)" a qualquer momento para baixar a grade que está ativa no seu navegador.
