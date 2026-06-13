<p align="center">
  <a href="https://www.seneon.com.br">
    <img src="assets/seneon_horizontal.png" width="280" alt="SENEON">
  </a>
</p>

# MalaDiretaPro

Aplicação desktop **gratuita** para mala direta com integração nativa ao Microsoft Outlook (COM). Todo o processamento é **100% local e offline** — nenhum dado sai da sua máquina (conformidade com LGPD). Sem servidor, sem credenciais SMTP: o envio usa a sessão do Outlook já aberta no seu Windows.

> Um produto [SENEON](https://www.seneon.com.br) · Conheça também o [Tykyra](https://seneon.vercel.app/lps/tykyra/)

---

## Pré-requisitos

- Windows 10 ou 11
- Microsoft Outlook clássico instalado e configurado com sua conta
- Uma planilha de contatos em `.xlsx` ou `.csv`

## Instalação

1. Baixe o `MalaDiretaPro.exe` mais recente na página de [**Releases**](https://github.com/SENEON-BR/MalaDiretaPro/releases)
2. Execute — não precisa instalar nada
3. O app avisa automaticamente quando há nova versão disponível

---

## Guia de uso — os 4 passos

### Passo 1 — Base de Dados

Clique em **Procurar…** e selecione sua planilha (`.xlsx` ou `.csv`). O app exibe uma prévia das primeiras linhas para você conferir se as colunas foram lidas corretamente. Linhas sem e-mail não impedem o uso — elas aparecem no Passo 4 como "Sem e-mail" e ficam registradas na planilha.

### Passo 2 — Mapeamento

Diga ao app qual coluna corresponde a cada campo:

| Campo | Obrigatório | Para que serve |
|---|---|---|
| E-mail | ✅ | Endereço do destinatário |
| Nome | — | Saudação personalizada e rótulo na lista |
| Gênero | — | Concordância automática: `{{Caro}}` vira "Caro"/"Cara" |
| Cargo / Empresa | — | Placeholders adicionais |

Em seguida escolha o **Modo de Envio**:

| Modo | O que faz |
|---|---|
| **Individual** | 1 e-mail por linha da planilha |
| **Em Grupo — individual** | 1 e-mail por pessoa, mas cada grupo tem seu próprio modelo (assunto, corpo, imagens, cópias) |
| **Em Grupo — resumo** | 1 e-mail consolidado por grupo, com tabela `{{TabelaGrupo}}` dos registros |
| **Único Global** | 1 só e-mail para todos (em Para ou Cco) |

Nos modos em grupo, selecione a **coluna de agrupamento** (ex.: ano, turma, categoria).

### Passo 3 — Composição

A tela é dividida: configurações à esquerda, editor do corpo à direita.

- **De (opcional)**: envia por outra conta/caixa compartilhada do seu perfil
- **Assunto**: aceita placeholders, ex.: `Convite {{Nome}}`
- **Cópia (CC) / Cópia Oculta (BCC)**: endereços fixos separados por `;`
- **Testeira / Rodapé**: imagens no topo e no fim do e-mail. Largura ideal: **600px** — imagens maiores são reduzidas proporcionalmente de forma automática
- **Anexos globais**: arraste arquivos ou use "Adicionar…"
- **Placeholders**: duplo-clique insere `{{Coluna}}` no cursor. Com coluna de gênero mapeada, use `{{Caro}}`, `{{Prezado}}`, `{{Sr}}`

**Modelos por grupo** (modos em grupo): a barra "Configurar para:" alterna entre o **Padrão** (herdado por todos) e cada grupo. Só vira personalização (★) o grupo que você **editar** — visitar sem mexer mantém a herança. O botão "🗑 Remover personalização" volta o grupo ao padrão.

### Passo 4 — Envio

- A lista mostra **todos** os registros: os sem e-mail aparecem como "Sem e-mail" e nunca são enviados
- O **preview** é fiel ao e-mail final: largura 600px, fundo branco, testeira/rodapé como o destinatário verá
- Nos modos em grupo, use o **filtro de grupo** para navegar pelos modelos
- **✉ Enviar teste para mim**: dispara uma amostra para o seu endereço antes do envio em massa
- **Salvar como rascunho**: os e-mails vão para a pasta *Rascunhos* do Outlook, para revisão manual
- **Intervalo entre envios**: pausa entre disparos (recomendado ≥ 2s para evitar bloqueio do servidor)
- Durante o envio: **⏸ Pausar** e **✖ Cancelar** a qualquer momento

### Registro na planilha (coluna DataEnvio)

Ao final do disparo, o app grava o resultado **linha a linha** na sua planilha:

| Situação | O que é gravado |
|---|---|
| Enviado com sucesso | `2026-06-12 14:35:00` (timestamp local) |
| Falha no envio | `não enviado: <motivo do erro>` |
| Linha sem e-mail | *(coluna fica em branco — linha não foi processada)* |

**Regras de reprocessamento:**

- **Linhas sem e-mail** não são incluídas na lista de envio — a coluna DataEnvio delas fica em branco.
- **Linhas já enviadas** (DataEnvio com timestamp) são automaticamente ignoradas quando você abre o app novamente e avança para o Passo 4. Isso significa que você pode reenviar com segurança para quem ainda não recebeu, sem risco de duplicar para quem já recebeu.
- **Erros anteriores podem ser re-tentados**: linhas com `não enviado: ...` são incluídas normalmente na próxima execução.
- Um timestamp de sucesso **nunca é sobrescrito**, mesmo que você execute o envio novamente.

---

## Para desenvolvedores

```bash
pip install -r requirements.txt

# Rodar o app
PYTHONPATH=src python -m maladireta

# Modo demonstração (sem Outlook — envio simulado)
PYTHONPATH=src python -m maladireta --demo

# Testes
PYTHONPATH=src python -m pytest

# Gerar o executável
./build.ps1
```

Ou use o menu interativo: `dev.bat`

### Stack

Python 3.11+ · PyQt6 · pandas · openpyxl · pywin32 (COM) · Pillow · PyInstaller

---

## Suporte

- 🐛 Encontrou um problema? [Abra uma issue](https://github.com/SENEON-BR/MalaDiretaPro/issues)
- 🌐 [www.seneon.com.br](https://www.seneon.com.br)
- 🚀 [Tykyra — gestão de revisões sistemáticas](https://seneon.vercel.app/lps/tykyra/)
