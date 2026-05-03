# VivWord

Sidebar de chat com **Claude (Anthropic API)** dentro do Microsoft Word — iPad, Mac, Windows e web. Add-in proprietário, sideloaded, uso pessoal.

---

## 1. O que é

Sidebar de revisão literária ao lado do manuscrito. Lê selecção ou documento todo como contexto, devolve as respostas do Claude por chat, inserção, substituição ou comentário na margem. Inclui *skills* (workflows guardados — Diagnóstico, Voz, Cortar gordura, Continuidade, Suspensão, Caça IA) e um campo "Voz da casa" onde colas parágrafos âncora teus que viajam como referência calibradora em cada chamada. A conversa, a voz e o system prompt ficam guardados por documento. Sem login, sem telemetria.

---

## 2. Como fazer deploy

Pré-requisitos: conta Vercel, chave Anthropic (https://console.anthropic.com/), `git` e `node`/`npx` instalados.

```bash
# 1. Clonar
git clone https://github.com/vivnasc/vivword.git
cd vivword

# 2. Primeiro deploy (cria o projecto na Vercel)
npx vercel
#  → escolher conta, "link to existing project? no", nome: vivword
#  → directório padrão, sem build command, sem framework

# 3. Configurar a chave da API
#    Vercel → Project: vivword → Settings → Environment Variables
#       Name:  ANTHROPIC_API_KEY
#       Value: sk-ant-…   (a chave criada em console.anthropic.com)
#       Environment: Production (e Preview/Development se quiseres)

# 4. Redeploy para a env var entrar em vigor
npx vercel --prod
```

O domínio final será algo como `https://vivword.vercel.app`. Se for diferente:

- Editar `manifest.xml` e substituir `vivword.vercel.app` pelo domínio real em **3 sítios**: `<IconUrl>`, `<HighResolutionIconUrl>`, `<SourceLocation>` e `<AppDomain>`.
- Voltar a fazer sideload do manifest actualizado.

Validar o manifest antes de instalar:

```bash
npx office-addin-manifest validate manifest.xml
```

---

## 3. Como sideload no iPad (Word iOS)

> O iPad é a plataforma mais picuinhas. A via mais fiável é **sincronizar via pasta partilhada / Mac**, não pelo botão directo no iPad.

### Via mais fiável: através do Mac (recomendada)

1. No **Mac** com a mesma conta Microsoft do iPad, copiar `manifest.xml` para:
   ```
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
   Se a pasta `wef` não existir, criar:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
2. Abrir Word no Mac uma vez (para registar). Verificar em **Insert → My Add-ins → Developer Add-ins** que o **VivWord** aparece.
3. No **iPad**, abrir o app **Microsoft Word** (não Safari) com a mesma conta Microsoft. Abrir um documento e ir a:
   **Insert → Add-ins → My Add-ins**. O VivWord aparece sincronizado.

### Alternativa: Shared Folder Catalog (Word Desktop → iPad)

1. Pôr `manifest.xml` numa pasta de rede (SMB/OneDrive sincronizado localmente).
2. No Word desktop (Mac ou Windows): **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**, adicionar o caminho da pasta, marcar "Show in Menu". Reiniciar Word.
3. **Insert → My Add-ins → Shared Folder** → escolher VivWord.
4. No iPad, com a mesma conta, o add-in aparece em **Insert → Add-ins → My Add-ins**.

### Alternativa directa no iPad (menos fiável)

1. Abrir um documento no app **Microsoft Word** no iPad (necessário ter subscrição Microsoft 365 activa).
2. **Insert → Add-ins → My Add-ins → ⋯ (Manage My Add-ins) → Upload My Add-in**.
3. Escolher o `manifest.xml` (precisa de estar acessível a partir do app Files do iPad — guardar antes em iCloud/OneDrive/Files).
4. Se não aparecer "Upload My Add-in", a tua versão/conta do Word iPad não suporta este caminho — usar a via Mac acima.

---

## 4. Como sideload no Word web

1. Abrir um documento em https://word.office.com .
2. **Insert → Add-ins → Upload My Add-in**.
3. Escolher `manifest.xml`. O botão **VivWord** aparece na ribbon (separador **Home**, ao fundo).

---

## 5. Como sideload no Word para Mac

```bash
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

Reabrir o Word. **Insert → My Add-ins → Developer Add-ins** → VivWord.

## 5b. Como sideload no Word para Windows

Caminho da pasta `wef`:

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

Copiar `manifest.xml` para lá. Reabrir Word. **Insert → My Add-ins → Developer Add-ins** → VivWord.

(Em redes corporativas pode ser mais simples usar **Shared Folder Catalog** — ver secção 3.)

---

## 6. O que faz (v1)

- **Chat com Claude** sobre o documento aberto. Streaming de respostas.
- **Botões de captura:** `Ler` (selecção) e `Doc` (documento todo).
- **Botões de devolução:** `Inserir` (no cursor), `Subst.` (sobre a selecção) e `Comentar` (na margem direita do Word, como nota literária — aparece se o host suportar WordApi 1.4).
- **Skills literárias** — uma fila de pílulas com workflows guardados: Diagnóstico, Voz, Cortar gordura, Continuidade, Suspensão, Caça IA. Cada uma é um ficheiro de texto em `public/skills/` que podes editar.
- **Voz da casa** — campo "Voz da casa" onde colas um a três parágrafos teus reconhecidos como inquestionavelmente da tua voz. São anexados como referência calibradora a cada chamada. Guardado por documento.
- **Persistência por documento** — a conversa, o modelo escolhido, o system prompt editado e a Voz da casa ficam guardados nos `Office.context.document.settings`. Cada manuscrito tem o seu próprio histórico.
- **Exportar conversa** — copia a conversa actual como Markdown para o clipboard.

## 6b. Limitações conhecidas

- **Documentos grandes (>50 000 palavras)** disparam aviso. `max_tokens=4096` por defeito.
- **Sem tracked changes** — `Subst.` troca o texto sem revisão. Para revisão lado-a-lado, usar `Comentar`.
- **Sem partilha** com Excel/PowerPoint — só Word.
- **`Comentar` requer WordApi 1.4** — botão fica oculto em hosts mais antigos. iPad e Word web modernos suportam.

---

## 7. Roadmap

- **v2** — comparar duas versões (slot A / slot B) numa skill dedicada.
- **v3** — exportação de conversa para .docx, não só Markdown.
- **v4** — sincronização das skills entre documentos (catálogo OneDrive em vez de só o repo).
- **v5** — modo "extended thinking" para análises estruturais profundas.

---

## 8. Estrutura

```
vivword/
├── manifest.xml                    ← XML clássico, WordApi 1.4
├── public/
│   ├── taskpane.html
│   ├── taskpane.js                 ← Office.js + chat + skills + persistência
│   ├── taskpane.css                ← tema escuro / terracota
│   ├── default-system-prompt.txt   ← instrução padrão (revisor literário)
│   ├── skills/
│   │   ├── index.json              ← catálogo de skills
│   │   ├── diagnostico.txt
│   │   ├── voz.txt
│   │   ├── cortar-gordura.txt
│   │   ├── continuidade.txt
│   │   ├── suspensao.txt
│   │   └── ia.txt
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
├── api/
│   └── chat.js               ← serverless Vercel → api.anthropic.com
├── package.json
├── vercel.json
├── .env.example
└── README.md
```

---

## 9. Instrução padrão, voz e skills

**Instrução para o Claude** (system prompt): vem pré-preenchida de `public/default-system-prompt.txt`. Edita o ficheiro no repo e faz commit + redeploy para mudar o default. O botão "Repor instrução padrão" recarrega a quente. Edições por documento ficam guardadas nas settings do próprio documento.

**Voz da casa**: por documento. Cola um a três parágrafos teus que tu reconheces como inquestionavelmente da tua voz. São anexados ao system prompt em cada chamada como bloco "Voz canónica desta obra" — o Claude passa a ter a tua voz à vista, não só descrita. Cada manuscrito tem a sua própria voz registada.

**Skills**: ficheiros em `public/skills/`. O catálogo é `skills/index.json`, cada entrada aponta para um `.txt` com o prompt da skill. Para adicionar uma skill:

1. Cria `public/skills/minha-skill.txt` com o prompt.
2. Acrescenta uma entrada em `skills/index.json` com `id`, `label`, `tip`, `context` (`selection` | `document`) e `file`.
3. Commit + redeploy.

A próxima vez que abres o taskpane, a skill aparece na fila.

---

## 10. Desenvolvimento local


```bash
cp .env.example .env
# preencher ANTHROPIC_API_KEY no .env
npx vercel dev
# abrir http://localhost:3000/taskpane.html
```

Para testar dentro do Word web sem deploy, usar um túnel (`cloudflared tunnel --url http://localhost:3000`) e apontar o `<SourceLocation>` do manifest para o URL do túnel temporariamente.
