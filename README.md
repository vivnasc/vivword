# VivWord

Sidebar de chat com **Claude (Anthropic API)** dentro do Microsoft Word — iPad, Mac, Windows e web. Add-in proprietário, sideloaded, uso pessoal.

---

## 1. O que é

Um taskpane que abre ao lado do documento e permite conversar com o Claude, ler a selecção ou o documento todo como contexto, e inserir/substituir texto. Sem login, sem persistência, sem telemetria.

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

## 6. Limitações conhecidas (v1)

- **Sem persistência de chat** entre fechar e reabrir o taskpane. O Office.js no iPad não dá storage 100% fiável; a sessão vive em memória.
- **Sem tracked changes** automáticas — o "Substituir" troca o texto seleccionado de forma directa.
- **Sem partilha** com Excel/PowerPoint — só Word.
- **Sem streaming** das respostas na v1 (pedido completo, depois aparece de uma vez). Adicionar streaming SSE numa próxima versão se for confortável no iPad.
- **Documentos grandes (>50.000 palavras)** disparam aviso. O backend impõe `max_tokens=4096` por defeito; para responder a manuscritos inteiros a janela de contexto chega, mas a resposta fica limitada.

---

## 7. Roadmap

- **v2** — tracked changes via `range.insertText` com formatação custom; histórico recente de respostas inseríveis.
- **v3** — persistência de conversa via OneDrive (file pickers do Office).
- **v4** — **SKILLS** reutilizáveis: templates de prompt guardados (ex.: "Revisão literária", "Resumo executivo BM público", "Tradução PT↔EN").
- **v5** — streaming SSE estável em iPad, exportação de conversa para .docx.

---

## 8. Política de uso (BM)

A autora trabalha no **Banco de Moçambique** sob **Ordem de Serviço n.º 18/2026**. O add-in inclui um aviso fixo no rodapé:

> ⚠ Não inserir dados confidenciais BM. Uso pessoal apenas.

E um toggle **Modo BM** que, quando activo, redacta no servidor (antes de chamar a Anthropic) padrões tipo NUITs (`\b\d{9}\b`) e códigos `BM-XXX` antes do envio. **Não substitui** uma triagem humana; é um cinto de segurança, não um cofre. Para conteúdo BM real, **não usar**.

---

## 9. Estrutura

```
vivword/
├── manifest.xml              ← XML clássico, WordApi 1.1
├── public/
│   ├── taskpane.html
│   ├── taskpane.js                 ← Office.js + chat
│   ├── taskpane.css                ← tema escuro / terracota
│   ├── default-system-prompt.txt   ← instrução padrão (revisor literário)
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

## 9b. Instrução padrão (system prompt)

O campo **"Instrução para o Claude"** no taskpane vem pré-preenchido a partir de `public/default-system-prompt.txt` (carregado por fetch no boot). Para mudar o default sem tocar em código, edita esse ficheiro e faz commit + redeploy. O botão **"Repor instrução padrão"** (visível quando a secção está expandida) recarrega o ficheiro a quente. O valor da sessão pode sempre ser editado livremente — só é enviado como `system` ao chamar a API.

---

## 10. Desenvolvimento local

```bash
cp .env.example .env
# preencher ANTHROPIC_API_KEY no .env
npx vercel dev
# abrir http://localhost:3000/taskpane.html
```

Para testar dentro do Word web sem deploy, usar um túnel (`cloudflared tunnel --url http://localhost:3000`) e apontar o `<SourceLocation>` do manifest para o URL do túnel temporariamente.
