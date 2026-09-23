# FileAttachmentControl — PCF para Power Apps Canvas

> **Responsabilidade única:** permitir que usuários anexem arquivos de múltiplos tipos (Office, PDF e imagens) dentro de um aplicativo Canvas, superando a limitação nativa do Power Apps que aceita apenas imagens.

---

## Índice

1. [Visão geral](#visão-geral)
2. [Tipos de arquivo suportados](#tipos-de-arquivo-suportados)
3. [Propriedades do componente](#propriedades-do-componente)
4. [Adicionando ao aplicativo Canvas](#adicionando-ao-aplicativo-canvas)
5. [Lendo os dados no Canvas](#lendo-os-dados-no-canvas)
6. [Integrando com Power Automate](#integrando-com-power-automate)
7. [Exemplos de fórmulas](#exemplos-de-fórmulas)

---

## Visão geral

O `FileAttachmentControl` é um componente PCF (Power Apps Component Framework) do tipo **campo** (`standard`). Ele renderiza uma zona de arrastar-e-soltar (_drag-and-drop_) e uma lista de chips para os arquivos adicionados, expondo o conteúdo via **Output Properties** — prontas para consumo por fórmulas Canvas ou fluxos Power Automate.

O componente **não faz upload** por conta própria; ele apenas lê, valida e codifica os arquivos em Base64, deixando a integração de destino (SharePoint, Dataverse, Blob Storage) a cargo do fluxo Power Automate.

---

## Tipos de arquivo suportados

| Categoria | Extensões |
|---|---|
| PDF | `.pdf` |
| Imagens | `.png`, `.jpg`/`.jpeg`, `.svg` |
| Office (OOXML) | `.docx`, `.pptx`, `.xlsx` |
| Office (legado) | `.doc`, `.ppt`, `.xls` |

**Tamanho máximo:** 25 MB por arquivo (configurável via propriedade `MaxFileSizeMB`).

---

## Propriedades do componente

### Entradas (Input)

| Propriedade | Tipo | Padrão | Descrição |
|---|---|---|---|
| `FilesJsonBound` | Texto (Bound) | — | Campo de texto opcional para persistir o JSON de arquivos no Dataverse/SharePoint. Pode ser deixado sem vínculo em apps Canvas. |
| `MaxFileSizeMB` | Número inteiro | `25` | Tamanho máximo permitido por arquivo, em MB. |
| `AllowMultiple` | Sim/Não | `Sim` | Quando `Não`, substituí qualquer arquivo anterior ao selecionar um novo. |
| `Theme` | Texto | `"light"` | Tema visual. Aceita `"light"` ou `"dark"`. |
| `SavedFiles` | Texto (JSON) | — | Arquivos já enviados (`[{ name, link, mimeType, id }]`). É relido sempre que o valor muda. |
| `ViewOnly` | Sim/Não | `Não` | Oculta a área de upload e os botões de remoção. |
| `ResetToken` | Texto | — | **Sempre que o valor mudar**, o componente descarta os arquivos pendentes e removidos, relê `SavedFiles` e publica `FilesJson`/`FilesRemoved` vazios. Veja [Limpando o componente após salvar](#limpando-o-componente-após-salvar). |

### Saídas (Output)

> Estas são as propriedades que você usa em fórmulas Canvas e em fluxos Power Automate.

| Propriedade | Tipo | Descrição |
|---|---|---|
| **`FilesJson`** | Texto | Array JSON principal. Cada item: `{ name, size, mimeType, base64, lastModified }`. Use este campo no Power Automate. |
| `FileCount` | Número inteiro | Quantidade de arquivos em staging no momento. |
| `HasFiles` | Sim/Não | `true` se pelo menos um arquivo foi adicionado. |
| `ValidationError` | Texto | Última mensagem de erro de validação (tipo não permitido ou tamanho excedido). Vazio quando não há erro. |
| `FilesRemoved` | Texto | Array JSON dos arquivos já enviados que o usuário removeu: `[{ name, link, mimeType, id }]`. |

---

## Adicionando ao aplicativo Canvas

1. **Publicar o componente** no ambiente Power Platform:
   ```powershell
   pac auth create --url https://<seu-org>.crm.dynamics.com
   pac pcf push --publisher-prefix <seu-prefixo>
   ```

2. No **Power Apps Studio**, abrir o aplicativo Canvas desejado.

3. No menu **Inserir → Componentes personalizados**, localizar `FTDEducacao.FileAttachmentControl`.

4. Inserir na tela desejada e redimensionar conforme o layout.

5. No painel de propriedades, configurar:
   - **MaxFileSizeMB** → ex.: `25`
   - **AllowMultiple** → `true` ou `false`
   - **Theme** → `"light"` ou `"dark"`

---

## Lendo os dados no Canvas

Use as Output Properties diretamente em fórmulas do Canvas:

```
// Exibir quantidade de arquivos
FileAttachmentControl1.FileCount

// Habilitar botão apenas quando há arquivos
Button1.DisplayMode = If(FileAttachmentControl1.HasFiles, DisplayMode.Edit, DisplayMode.Disabled)

// Exibir erro de validação
Label1.Text = FileAttachmentControl1.ValidationError

// Capturar o JSON completo para enviar ao Power Automate
Set(varFilesJson, FileAttachmentControl1.FilesJson)
```

---

## Integrando com Power Automate

O campo **`FilesJson`** é o ponto principal de integração. Estrutura de cada item do array:

```json
[
  {
    "name": "proposta.docx",
    "size": 204800,
    "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "base64": "UEsDBBQABgAIAAAAIQA...",
    "lastModified": 1709000000000
  }
]
```

### Fluxo sugerido (SharePoint)

```
Trigger: Power Apps (Canvas) → botão "Salvar"
  Entrada: varFilesJson (texto)

Ação 1: Parse JSON
  Conteúdo: varFilesJson
  Schema: gerar a partir do exemplo acima

Ação 2: Apply to each (item do array)
  Ação interna: Criar arquivo — SharePoint
    Nome do arquivo: item()?['name']
    Conteúdo do arquivo: base64ToBinary(item()?['base64'])
    Pasta: /sites/seu-site/Documentos Compartilhados/Anexos
```

---

## Exemplos de fórmulas

### Enviar ao Power Automate ao clicar em Salvar

```
// OnSelect do botão Salvar
If(
    FileAttachmentControl1.HasFiles,
    MeuFluxo.Run(FileAttachmentControl1.FilesJson),
    Notify("Adicione pelo menos um arquivo antes de salvar.", NotificationType.Warning)
)
```

### Limpando o componente após salvar

No Canvas, as saídas de um code component (`FilesJson`, `FilesRemoved`, ...) guardam o **último valor publicado** durante toda a sessão do app, mesmo depois de navegar para outra tela. Se o app não limpar o componente depois do envio, o próximo salvamento manda os mesmos arquivos de novo ao fluxo e duplica os anexos.

Ligue `ResetToken` a uma variável e dê um valor novo a ela depois de enviar os arquivos e ao abrir outro registro:

```
// Propriedade ResetToken do controle
ResetToken = vResetAnexo

// Depois de chamar o fluxo com FilesJson/FilesRemoved
MeuFluxo.Run(FileAttachmentControl1.FilesJson, FileAttachmentControl1.FilesRemoved);
Set(vResetAnexo, GUID())

// Ao abrir um registro (antes do Navigate para a tela do formulário)
Set(vFilesAttachment, /* JSON dos anexos salvos */);
Set(vResetAnexo, GUID())
```

### Mostrar banner de erro condicionalmente

```
// Visible do label de erro
!IsBlank(FileAttachmentControl1.ValidationError)
```

### Exibir nome do primeiro arquivo (parsing do JSON)

```
First(ParseJSON(FileAttachmentControl1.FilesJson)).name
```
