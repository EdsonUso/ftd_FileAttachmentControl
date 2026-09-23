# Guia de Contribuição — FileAttachmentControl

Bem-vindo! Este documento explica a estrutura interna do PCF, como rodar localmente, como estender o componente e boas práticas para resolver bugs.

---

## Índice

1. [Pré-requisitos](#pré-requisitos)
2. [Estrutura do projeto](#estrutura-do-projeto)
3. [Rodando localmente](#rodando-localmente)
4. [Arquitetura do código](#arquitetura-do-código)
5. [Como estender o componente](#como-estender-o-componente)
6. [Resolvendo bugs comuns](#resolvendo-bugs-comuns)
7. [Publicando uma nova versão](#publicando-uma-nova-versão)
8. [Convenções de código](#convenções-de-código)

---

## Pré-requisitos

| Ferramenta | Versão mínima |
|---|---|
| Node.js | 18+ |
| PAC CLI (`pac`) | instalado e autenticado |
| npm | 9+ |

**Instalar dependências:**
```powershell
cd c:\Users\edson\Documents\trabalho\pcf
cmd /c "npm install"
```

---

## Estrutura do projeto

```
pcf/
├── FileAttachmentControl/
│   ├── ControlManifest.Input.xml   # Manifesto: propriedades, recursos e metadados
│   ├── index.ts                    # Lógica principal do PCF
│   ├── css/
│   │   └── FileAttachmentControl.css  # Estilos (design tokens, temas, animações)
│   └── generated/
│       └── ManifestTypes.d.ts      # Tipos gerados automaticamente (não editar)
├── package.json
├── tsconfig.json
├── eslint.config.mjs
└── pcf.pcfproj
```

> **`generated/ManifestTypes.d.ts`** é regenerado automaticamente a cada `npm run build`. Nunca edite esse arquivo manualmente.

---

## Rodando localmente

```powershell
# Build único
cmd /c "npm run build"

# Modo watch com test harness (http://localhost:8181)
cmd /c "npm start watch"
```

No test harness você pode alterar os valores das Input Properties em tempo real no painel lateral e observar as Output Properties sendo atualizadas.

---

## Arquitetura do código

O componente implementa a interface `ComponentFramework.StandardControl<IInputs, IOutputs>` com quatro métodos obrigatórios:

### `init()`
Chamado **uma única vez** quando o controle é montado. Responsabilidades:
- Ler as Input Properties iniciais (`_readInputs`)
- Construir o DOM (`_buildDOM`)
- Registrar event listeners (`_attachEvents`)

### `updateView()`
Chamado **sempre que uma Input Property muda** no Canvas. Responsabilidades:
- Reler as inputs
- Sincronizar o atributo `multiple` do `<input type="file">`
- Aplicar o tema (`data-theme`)

> ⚠️ **Não recriar o DOM aqui.** `updateView` deve apenas atualizar atributos/classes existentes, não recriar elementos — isso causaria perda de estado e re-renders desnecessários.

### `getOutputs()`
Chamado pelo framework **após** `notifyOutputChanged()` ser disparado. Deve retornar o estado atual de `_stagedFiles`, `_validationError` etc. É o único ponto de saída de dados.

### `destroy()`
Chamado quando o componente é desmontado. Por ora vazio, pois os event listeners estão em elementos DOM que serão coletados pelo GC junto com o componente.

---

### Fluxo de dados interno

```
Usuário arrasta/seleciona arquivo
        │
        ▼
_processFiles(files: File[])
  ├─ Valida MIME type (ALLOWED_MIME)
  ├─ Valida tamanho (MaxFileSizeMB)
  ├─ Verifica duplicatas (nome + tamanho)
  └─ FileReader.readAsDataURL()
        │
        ▼
_stagedFiles.push({ name, size, mimeType, base64, lastModified })
        │
        ▼
_renderChips()        →  Atualiza DOM dos chips
notifyOutputChanged() →  Dispara getOutputs() no framework
        │
        ▼
getOutputs()  →  Retorna FileCount, FilesJson, HasFiles, ValidationError
```

### Constantes importantes

| Constante | Onde | O que é |
|---|---|---|
| `DEFAULT_MAX_MB` | `index.ts` | Limite padrão de tamanho (25 MB) |
| `ALLOWED_MIME` | `index.ts` | `Set<string>` com os MIMEs aceitos |
| `MIME_LABEL` | `index.ts` | Mapa MIME → badge de texto do chip |

---

## Como estender o componente

### Adicionar um novo tipo de arquivo

1. Adicionar o MIME type ao `Set` `ALLOWED_MIME` em `index.ts`:
   ```ts
   "video/mp4",   // .mp4
   ```

2. Adicionar a extensão ao atributo `accept` do `<input>`:
   ```ts
   this._fileInput.accept = [..., ".mp4"].join(",");
   ```

3. (Opcional) Adicionar ao mapa `MIME_LABEL` e criar uma classe CSS `.fac-chip--vid` com a cor do badge em `FileAttachmentControl.css`.

---

### Adicionar uma nova Output Property

1. Declarar a property no `ControlManifest.Input.xml`:
   ```xml
   <property name="MinhaProperty" of-type="SingleLine.Text" usage="output" />
   ```

2. Após o próximo build, `ManifestTypes.d.ts` será regenerado com o tipo novo em `IOutputs`.

3. Retornar o valor em `getOutputs()`:
   ```ts
   public getOutputs(): IOutputs {
     return {
       ...
       MinhaProperty: this._meuValor,
     };
   }
   ```

4. Guardar o valor no estado interno e chamar `this._notifyOutputChanged()` quando ele mudar.

---

### Adicionar uma nova Input Property

1. Declarar no `ControlManifest.Input.xml` com `usage="input"`.
2. Ler em `_readInputs(context)`:
   ```ts
   const raw = context.parameters.MinhaInput?.raw;
   this._minhaInput = raw ?? valorPadrao;
   ```
3. Usar o valor em `init()` ou `updateView()` conforme necessário.

---

### Adicionar preview de imagem

Atualmente os arquivos são exibidos como chips de texto. Para adicionar preview de imagem:

1. Em `_renderChips()`, verificar se `sf.mimeType.startsWith("image/")`.
2. Se sim, criar um `<img>` com `src="data:${sf.mimeType};base64,${sf.base64}"`.
3. Controlar o tamanho via CSS para não impactar a performance.

> ⚠️ **Atenção:** arquivos SVG com `<script>` embutido são um vetor XSS se renderizados diretamente via `<img src="data:...">`. Considere sanitizar SVGs antes do preview ou usar `sandbox`.

---

## Resolvendo bugs comuns

### O Output `FilesJson` não atualiza no Canvas

**Causa:** `notifyOutputChanged()` não foi chamado após a mudança de estado.  
**Fix:** Confirmar que toda alteração em `_stagedFiles` ou `_validationError` é sempre seguida de `this._notifyOutputChanged()`.

---

### Arquivo rejeitado mas o MIME está na lista

**Causa:** Alguns SOs/navegadores reportam o MIME de formas alternativas. Por exemplo, arquivos `.doc` antigos às vezes chegam como `application/octet-stream`.  
**Fix:** Adicionar verificação por extensão como fallback:
```ts
const ext = f.name.split(".").pop()?.toLowerCase() ?? "";
const allowed = ALLOWED_MIME.has(f.type) || FALLBACK_EXTS.has(ext);
```

---

### Drag-and-drop pisca ao arrastar sobre elementos filhos

**Causa:** `dragleave` é disparado ao entrar em elementos filhos.  
**Fix:** Já tratado com `_dragCounter`. Se houver regressão, confirmar que `dragenter` incrementa e `dragleave` decrementa corretamente, e que a classe `fac-dropzone--active` é removida apenas quando `_dragCounter === 0`.

---

### Build falha: `error TS2345` nos tipos de Output

**Causa:** Uma property foi adicionada ao manifesto mas `ManifestTypes.d.ts` ainda não foi regenerado.  
**Fix:** Rodar `npm run build` novamente. Se o erro persistir, apagar `FileAttachmentControl/generated/` e rebuildar.

---

### ESLint: `consistent-generic-constructors`

**Causa:** O ESLint do projeto exige que os type arguments de generics fiquem no lado do construtor.  
**Errado:** `const x: Set<string> = new Set([...])`  
**Certo:** `const x = new Set<string>([...])`

---

## Publicando uma nova versão

1. Incrementar a `version` no `ControlManifest.Input.xml` (ex.: `1.1.0` → `1.2.0`). **Sem isso o Canvas não oferece a atualização do componente.** Mantenha `<Version>` em `Solutions/src/Other/Solution.xml` alinhado.
2. Rebuildar e validar:
   ```powershell
   cmd /c "npm run build"
   cmd /c "npm run lint"
   ```
3. Conferir o ambiente ativo do PAC (o componente está no **ftd-dev2**, solução `Plano_upload_attachment`):
   ```powershell
   pac auth list
   pac env who
   pac auth select --index <n>   # se precisar trocar
   ```
4. Enviar para o ambiente (dev, componente **não gerenciado**):
   ```powershell
   pac pcf push --publisher-prefix ftd --solution-unique-name Plano_upload_attachment
   ```
   Alternativa: gerar o `.zip` da solução com `dotnet build` dentro de `Solutions/` (Debug = não gerenciado, Release = gerenciado, saída em `Solutions/bin/<config>/Solutions.zip`) e importar em **Soluções → Importar**. Use a versão gerenciada só em ambientes onde o componente já é gerenciado.
5. Publicar as personalizações (**Soluções → Publicar todas as personalizações**), se o push/import não tiver publicado.
6. No Power Apps Studio, abrir o app Canvas. Ele avisa que há atualização de code component: aceite, configure as propriedades novas, salve e **publique** o app. Se o aviso não aparecer, feche e reabra o app (o Studio só atualiza code components ao abrir o app).
7. Atualizar o `README.md` se houver mudanças nas propriedades.

---

## Convenções de código

- **Prefixo `_`** em todos os campos privados da classe (ex.: `_stagedFiles`, `_maxMB`).
- **Prefixo `fac-`** em todas as classes CSS (*File Attachment Control*).
- **Português** nos comentários e mensagens de UI voltadas ao usuário final.
- **Inglês** nos identificadores TypeScript (nomes de variáveis, métodos, propriedades).
- Não usar `any` — preferir tipos explícitos ou `unknown` com type guard.
- Não usar `innerHTML` para inserir dados provenientes do usuário (XSS).
