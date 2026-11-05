# Sistema de Automação Operacional de Atendimentos (GESUAS)

**Órgão:** Centro de Referência de Assistência Social — CRAS Sul  
**Município:** Campo Belo / MG  
**Desenvolvimento:** Juliano França da Mata  
**Formação:** Bacharel em Ciência de Dados e Inteligência Artificial | Tecnólogo em Gestão de Tecnologia da Informação  
**Função Operacional:** Agente Administrativo / Operador do Cadastro Único  

---

## 🎯 Objetivo Operacional

Automatizar a rotina de lançamento e registro de atendimentos e atualizações cadastrais no prontuário eletrônico do sistema GESUAS a partir de dados consolidados em planilhas eletrônicas (armazenadas na pasta `DADOS_BRUTOS/`), eliminando retrabalho operacional, prevenindo duplicidades e assegurando a rastreabilidade via relatórios analíticos em PDF.

```text
[ Planilha(s) em DADOS_BRUTOS/ ]
                 │
                 ▼
    [ AUTOMACAO_GESUAS (Robô) ] ──► [ Navegação CDP Edge / GESUAS ]
                 │
                 ▼
    [ Relatorios_GESUAS / PDFs ] ──► Confirmação, Recibos e Auditoria
```

---

## ⚠️ Estrutura Organizacional e Portabilidade no Pendrive

A aplicação foi adaptada para operar em conformidade com a nova estrutura padronizada da unidade removível (pendrive padrão `E:\` ou qualquer letra disponível de `D:` a `Z:`):

```text
[UNIDADE_PENDRIVE]:\CADASTROUNICO\
│
├── DADOS/                        # Diretório com os comprovantes anuais (Auditor)
├── PLANILHAS/                    # Cópias e histórico geral
└── SOTWARES/
    ├── Auditor_GESUAS/           # Módulo de auditoria cadastral prévia
    │
    └── Robo_GESUAS/              # Módulo de automação de atendimentos
        ├── _internal/            # Dependências e bibliotecas embutidas
        ├── DADOS_BRUTOS/         # Diretório das planilhas anuais de entrada
        │   └── CADASTRO_ATUALIZACAO_2026.xlsx
        ├── Relatorios_GESUAS/    # Relatórios executivos em PDF com timestamp
        └── Robo_GESUAS.exe       # Executável do robô
```

---

## 📁 Estrutura do Repositório (`AUTOMACAO_GESUAS`)

```text
AUTOMACAO_GESUAS/
│
├── build/                         # Arquivos temporários de compilação (PyInstaller)
├── DADOS_BRUTOS/                  # Pasta dedicada para as planilhas anuais (.xlsx)
│   └── CADASTRO_ATUALIZACAO_2026.xlsx
├── dist/                          # Distribuição portátil autônoma
│   └── Robo_GESUAS/               # Pasta autocontida para transporte no pendrive
│       ├── _internal/             # Dependências e binários embutidos
│       └── Robo_GESUAS.exe        # Executável compilado
├── IMAGENS/                       # Recursos visuais e ícones corporativos (icon.ico)
├── Relatorios_GESUAS/             # Recibos e relatórios de execução em PDF com timestamp
│
├── SCRIPTS/                       # Módulos operacionais do sistema
│   ├── config.py                  # Caminhos, detecção em DADOS_BRUTOS, CDP e cabeçalho
│   ├── gesuas_bot.py              # Automação de navegação, anti-duplicidade e DOM (Playwright)
│   ├── main.py                    # Orquestrador com menu dinâmico, ESC e loop multi-mês
│   ├── notificacoes.py            # Bipes em canal PCM (winsound) e caixas modais TopMost
│   └── relatorio_pdf.py           # Geração de relatórios executivos em PDF (ReportLab)
│
├── .gitattributes                 # Normalização de quebras de linha e arquivos binários
├── .gitignore                     # Diretrizes de exclusão do Git e proteção LGPD
├── iniciar_automacao.bat          # Disparador facilitado via terminal (duplo clique)
├── LICENSE                        # Licença MIT
├── README.md                      # Documentação técnica e operacional
├── requirements.txt               # Relação congelada de dependências do ambiente Python
└── Robo_GESUAS.spec               # Especificação técnica de compilação do PyInstaller
```

---

## ⚙️ Diretrizes Operacionais e Funcionalidades Técnicas

1. **Varredura Dinâmica em `DADOS_BRUTOS/`:** O robô detecta automaticamente todas as planilhas `.xlsx` presentes na pasta dedicada. Caso haja mais de uma (ex.: 2025, 2026, 2027), apresenta um menu para seleção interativa da planilha de trabalho.
2. **Seleção Inteligente de Abas (Meses):** Exibe as abas disponíveis na planilha, identificando automaticamente o mês corrente e permitindo navegar entre abas sem fechar o sistema.
3. **Pausa e Interrupção via Tecla `ESC`:** O operador pode pressionar a tecla `ESC` a qualquer momento durante a execução no console para pausar o processamento, permitindo encerrar antecipadamente o lote gerando o relatório parcial do que já foi lançado.
4. **Acoplamento CDP ao Microsoft Edge (Porta 9222):** A automação se conecta a uma instância supervisionada do Microsoft Edge via Chrome DevTools Protocol, permitindo login manual seguro com credenciais do operador.
5. **Prevenção Estrita de Duplicidade:** Antes de registrar qualquer atendimento, o robô abre a modal `Editar um Atendimento`, inspeciona o histórico assíncrono via JavaScript e, caso a mesma data/hora já conste lançada, ignora o cadastro e registra o status `JÁ CADASTRADO`.
6. **Classificação Automática de Modalidades e Ações:**
   * **Ações:** Atualização, Inclusão, Consulta (sem alteração de prontuário) e Exclusão (requer ação manual).
   * **Modalidades (coluna OBS):** Unipessoal (CRAS), Unipessoal (Visita Domiciliar), Família (Visita Domiciliar) e Família (Atend. CRAS).
7. **Identificação e Tratamento de Não-RF:** Se a pessoa atendida não for a Pessoa de Referência da família, o robô desmarca o RF e marca a pessoa correta em "Pessoas atendidas", emitindo sinal sonoro e solicitando confirmação visual do operador.
8. **Relatórios Oficiais em PDF Paisagem:** Relatório detalhado gerado com ReportLab contendo cabeçalho executivo, resumo totalizador, legenda cromática compacta e observações tipográficas (`<i><strike>Visita</strike></i>` e `<b>Unipessoal</b>`).

---

## 🚀 Como Executar no Ambiente de Desenvolvimento

### Pré-requisitos

* Microsoft Edge instalado.
* Python 3.10+ com ambiente virtual configurado (`.venv`).
* Planilha colocada na pasta `DADOS_BRUTOS/`.

### Execução via Batch (`.bat`)

Dê **duplo clique em `iniciar_automacao.bat`** na raiz. O arquivo ativa a `.venv` automaticamente e dispara `SCRIPTS/main.py`.

### Execução via Terminal (VS Code)

```powershell
# Com o terminal na raiz AUTOMACAO_GESUAS e a .venv ativa:
python SCRIPTS/main.py
```

---

## 📦 Compilação e Atualização no Pendrive

Para compilar a distribuição executável autônoma (sem necessidade de Python no computador do CRAS):

### 1. Comando de Compilação

No terminal do VS Code, na raiz `AUTOMACAO_GESUAS` com a `.venv` ativada:

```powershell
python -m PyInstaller --noconfirm --onedir --console --icon="IMAGENS/icon.ico" --name "Robo_GESUAS" --collect-all playwright --collect-all reportlab --paths "SCRIPTS" SCRIPTS/main.py
```

### 2. Instalação no Pendrive

1. Acesse a pasta gerada: `AUTOMACAO_GESUAS\dist\Robo_GESUAS\`.
2. Copie a pasta `_internal/` e o executável `Robo_GESUAS.exe`.
3. Cole na pasta do pendrive: `E:\CADASTROUNICO\SOTWARES\Robo_GESUAS\`.
4. Certifique-se de que a planilha anual está dentro de `E:\CADASTROUNICO\SOTWARES\Robo_GESUAS\DADOS_BRUTOS\`.
5. Dê duplo clique em `Robo_GESUAS.exe` para operar.

---

## 🔄 Fluxo Operacional Integrado (Auditoria ➔ Automação)

Para garantir 100% de integridade no prontuário do GESUAS, utilize a rotina institucional recomendada:

1. **Etapa 1 — Auditoria Cadastral (`Auditor_GESUAS.exe`):**  
   Executa a conciliação entre os comprovantes PDF em `CADASTROUNICO\DADOS\` e o GESUAS, gerando o relatório diagnóstico de divergências de membros, códigos familiares ausentes e titularidade de RF.
2. **Etapa 2 — Intervenção Manual Pontual:**  
   O operador resolve previamente apenas as exceções apontadas com o status `INTERVENÇÃO NECESSÁRIA`.
3. **Etapa 3 — Automação Segura em Lote (`Robo_GESUAS.exe`):**  
   O robô processa a planilha mensal em `DADOS_BRUTOS/`, lançando os atendimentos no GESUAS com total fidelidade de horário e emitindo os recibos consolidados em PDF.

---

## 🔒 Governança de Dados e Conformidade (LGPD)

O sistema manipula informações cadastrais e registros de atendimento protegidos pela **Lei Geral de Proteção de Dados (Lei nº 13.709/2018)**:

* **Armazenamento Estritamente Local:** Planilhas, arquivos de depuração e relatórios gerados permanecem restritos à máquina local ou pendrive de trabalho.
* **Isolamento de Repositório:** O arquivo `.gitignore` bloqueia qualquer envio acidental de arquivos `.xlsx`, relatórios `.pdf` ou dados pessoais a repositórios de controle de versão (GitHub/GitLab).
* **Anonimização:** Quaisquer capturas de tela ou exemplos públicos utilizam dados sintéticos ou tarjados.

---

## 📄 Licença e Responsabilidade

Distribuído sob a **Licença MIT** (consulte o arquivo `LICENSE` para detalhes).  
Desenvolvido para apoio operacional no CRAS Sul — Prefeitura Municipal de Campo Belo/MG.