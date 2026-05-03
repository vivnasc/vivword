// VivWord — sidebar de chat com Claude dentro do Word.
// Vanilla JS. Sem build. Compatível com WordApi 1.1 (iPad-friendly).

(function () {
  'use strict';

  // ---------- Estado de sessão (não persiste entre fechar/abrir) ----------
  const state = {
    messages: [],          // [{ role: 'user'|'assistant', content: '...' }]
    pendingContext: null,  // { kind: 'selection'|'document', text, words }
    lastAssistant: '',     // última resposta do Claude (para Inserir/Substituir)
    inFlight: false,
  };

  const WORDS_WARN = 50000;

  // ---------- DOM refs ----------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);
  const els = {
    chat: $('#chat'),
    input: $('#input'),
    composer: $('#composer'),
    sendBtn: $('#sendBtn'),
    loading: $('#loading'),
    model: $('#model'),
    bmMode: $('#bmMode'),
    sysToggle: $('#sysToggle'),
    sysPrompt: $('#sysPrompt'),
    sysReset: $('#sysReset'),
    contextChip: $('#contextChip'),
    contextLabel: $('#contextLabel'),
    contextClear: $('#contextClear'),
    clearChat: $('#clearChat'),
    app: $('#app'),
  };

  // ---------- Default system prompt (carregado de ficheiro) ----------
  async function loadDefaultSystemPrompt({ force = false } = {}) {
    try {
      const r = await fetch('default-system-prompt.txt', { cache: 'no-store' });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const text = (await r.text()).trim();
      if (!text) return;
      // Só sobrescreve se o campo estiver vazio, ou se for reposição explícita.
      if (force || !els.sysPrompt.value.trim()) {
        els.sysPrompt.value = text;
      }
      els.sysReset.dataset.available = '1';
      // Visível apenas quando a secção do prompt está expandida.
      const expanded = els.sysToggle.getAttribute('aria-expanded') === 'true';
      els.sysReset.hidden = !expanded;
      els.sysPrompt.placeholder = 'Ex.: És revisor literário. Mantém o estilo da autora.';
    } catch (e) {
      els.sysPrompt.placeholder = 'Ex.: És revisor literário. Mantém o estilo da autora.';
      // Falha silenciosa; o campo continua editável e a sessão funciona.
      if (typeof console !== 'undefined') console.warn('default-system-prompt:', e.message);
    }
  }

  // ---------- Office.js boot ----------
  let officeReady = false;
  if (typeof Office !== 'undefined') {
    Office.onReady(() => {
      officeReady = true;
      loadDefaultSystemPrompt();
    });
  } else {
    // Fora do Word (preview no browser): carregar à mesma.
    loadDefaultSystemPrompt();
  }

  // ---------- Render ----------
  function clearEmptyHint() {
    const hint = els.chat.querySelector('.empty-hint');
    if (hint) hint.remove();
  }

  function appendMessage(role, text, meta) {
    clearEmptyHint();
    const div = document.createElement('div');
    div.className = 'msg ' + (role === 'user' ? 'user' : role === 'error' ? 'error' : 'claude');
    div.textContent = text;
    if (meta) {
      const m = document.createElement('span');
      m.className = 'meta';
      m.textContent = meta;
      div.appendChild(m);
    }
    els.chat.appendChild(div);
    els.chat.scrollTop = els.chat.scrollHeight;
    return div;
  }

  function setLoading(on) {
    state.inFlight = on;
    els.loading.hidden = !on;
    els.sendBtn.disabled = on;
    els.app.setAttribute('aria-busy', on ? 'true' : 'false');
    if (on) els.chat.scrollTop = els.chat.scrollHeight;
  }

  function showContextChip() {
    if (!state.pendingContext) {
      els.contextChip.hidden = true;
      els.contextLabel.textContent = '';
      return;
    }
    const c = state.pendingContext;
    const label =
      c.kind === 'selection'
        ? `Selecção anexa (${c.words} palavra${c.words === 1 ? '' : 's'})`
        : `Documento anexo (${c.words} palavra${c.words === 1 ? '' : 's'})`;
    els.contextLabel.textContent = label;
    els.contextChip.hidden = false;
  }

  function countWords(s) {
    if (!s) return 0;
    const m = s.trim().match(/\S+/g);
    return m ? m.length : 0;
  }

  // ---------- Office acções (WordApi 1.1) ----------
  function ensureOffice() {
    if (typeof Word === 'undefined' || typeof Office === 'undefined') {
      throw new Error('Office.js não disponível. Abre o add-in dentro do Word.');
    }
  }

  async function readSelection() {
    ensureOffice();
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load('text');
      await context.sync();
      return sel.text || '';
    });
  }

  async function readDocument() {
    ensureOffice();
    return Word.run(async (context) => {
      const body = context.document.body.getRange();
      body.load('text');
      await context.sync();
      return body.text || '';
    });
  }

  async function insertAtCursor(text) {
    ensureOffice();
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertText(text, 'Replace');
      await context.sync();
    });
  }

  async function replaceSelection(text) {
    // Em WordApi 1.1, "Replace" no insertText sobre a selecção substitui-a.
    return insertAtCursor(text);
  }

  // ---------- Acções da toolbar ----------
  async function actionReadSelection() {
    try {
      const text = await readSelection();
      if (!text || !text.trim()) {
        appendMessage('error', 'Nada seleccionado no documento.');
        return;
      }
      const words = countWords(text);
      state.pendingContext = { kind: 'selection', text, words };
      showContextChip();
      appendMessage('claude', `Selecção lida: ${words} palavra${words === 1 ? '' : 's'}. Anexada à próxima mensagem.`);
    } catch (e) {
      appendMessage('error', 'Erro ao ler selecção: ' + e.message);
    }
  }

  async function actionReadDocument() {
    try {
      const text = await readDocument();
      if (!text || !text.trim()) {
        appendMessage('error', 'Documento vazio.');
        return;
      }
      const words = countWords(text);
      state.pendingContext = { kind: 'document', text, words };
      showContextChip();
      let note = `Documento lido: ${words.toLocaleString('pt-PT')} palavras. Anexado à próxima mensagem.`;
      if (words > WORDS_WARN) {
        note += `\n\n⚠ Documento grande (>${WORDS_WARN.toLocaleString('pt-PT')} palavras). Pode demorar e consumir muitos tokens.`;
      }
      appendMessage('claude', note);
    } catch (e) {
      appendMessage('error', 'Erro ao ler documento: ' + e.message);
    }
  }

  async function actionInsert() {
    if (!state.lastAssistant) {
      appendMessage('error', 'Ainda não há resposta do Claude para inserir.');
      return;
    }
    try {
      await insertAtCursor(state.lastAssistant);
      appendMessage('claude', 'Resposta inserida no documento.');
    } catch (e) {
      appendMessage('error', 'Erro ao inserir: ' + e.message);
    }
  }

  async function actionReplace() {
    if (!state.lastAssistant) {
      appendMessage('error', 'Ainda não há resposta do Claude para usar.');
      return;
    }
    try {
      await replaceSelection(state.lastAssistant);
      appendMessage('claude', 'Selecção substituída.');
    } catch (e) {
      appendMessage('error', 'Erro ao substituir: ' + e.message);
    }
  }

  // ---------- Envio para o backend ----------
  function buildMessagesForApi(userText) {
    // Se há contexto pendente, prepende-o à mensagem do utilizador como bloco rotulado.
    let composed = userText;
    if (state.pendingContext) {
      const c = state.pendingContext;
      const label = c.kind === 'selection' ? 'SELECÇÃO DO DOCUMENTO' : 'DOCUMENTO COMPLETO';
      composed = `[${label}]\n${c.text}\n[/${label}]\n\n${userText}`;
    }
    const history = state.messages.map((m) => ({ role: m.role, content: m.content }));
    history.push({ role: 'user', content: composed });
    return history;
  }

  async function sendToClaude(userText) {
    const apiMessages = buildMessagesForApi(userText);
    const payload = {
      model: els.model.value,
      messages: apiMessages,
      system: els.sysPrompt.value.trim(),
      redactBM: !!els.bmMode.checked,
      stream: false,
    };

    const r = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(payload),
    });

    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try {
        const j = await r.json();
        if (j && j.error) msg = j.error;
      } catch (_) {}
      throw new Error(msg);
    }

    const data = await r.json();
    const blocks = Array.isArray(data.content) ? data.content : [];
    const text = blocks.filter((b) => b && b.type === 'text').map((b) => b.text).join('\n').trim();
    if (!text) throw new Error('Resposta vazia da Anthropic.');
    return text;
  }

  async function handleSend() {
    if (state.inFlight) return;
    const userText = els.input.value.trim();
    if (!userText && !state.pendingContext) return;

    const effective = userText || '(Analisa o texto anexo.)';
    appendMessage('user', effective);

    // Guarda no histórico a versão composta (com contexto), para o Claude poder referir-se.
    const composedForHistory = (() => {
      if (!state.pendingContext) return effective;
      const c = state.pendingContext;
      const label = c.kind === 'selection' ? 'SELECÇÃO DO DOCUMENTO' : 'DOCUMENTO COMPLETO';
      return `[${label}]\n${c.text}\n[/${label}]\n\n${effective}`;
    })();

    els.input.value = '';
    autoresize();

    setLoading(true);
    try {
      const reply = await sendToClaude(effective);
      state.messages.push({ role: 'user', content: composedForHistory });
      state.messages.push({ role: 'assistant', content: reply });
      state.lastAssistant = reply;
      state.pendingContext = null;
      showContextChip();
      appendMessage('claude', reply);
    } catch (e) {
      appendMessage('error', e.message);
    } finally {
      setLoading(false);
    }
  }

  // ---------- UI helpers ----------
  function autoresize() {
    els.input.style.height = 'auto';
    els.input.style.height = Math.min(els.input.scrollHeight, 140) + 'px';
  }

  function clearChat() {
    state.messages = [];
    state.lastAssistant = '';
    state.pendingContext = null;
    showContextChip();
    els.chat.innerHTML = '<div class="empty-hint"><p>Conversa limpa. Pergunta algo novo.</p></div>';
  }

  // ---------- Bindings ----------
  document.addEventListener('DOMContentLoaded', () => {
    $$('.tool').forEach((btn) => {
      btn.addEventListener('click', () => {
        const action = btn.dataset.action;
        if (action === 'read-selection') actionReadSelection();
        else if (action === 'read-document') actionReadDocument();
        else if (action === 'insert') actionInsert();
        else if (action === 'replace') actionReplace();
      });
    });

    els.composer.addEventListener('submit', (e) => {
      e.preventDefault();
      handleSend();
    });

    els.input.addEventListener('input', autoresize);
    els.input.addEventListener('keydown', (e) => {
      // Enter envia, Shift+Enter quebra linha. No iPad, o teclado mostra "return"; mantemos o botão Enviar visível.
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    });

    els.sysToggle.addEventListener('click', () => {
      const expanded = els.sysToggle.getAttribute('aria-expanded') === 'true';
      els.sysToggle.setAttribute('aria-expanded', expanded ? 'false' : 'true');
      els.sysPrompt.hidden = expanded;
      els.sysReset.hidden = expanded || !els.sysReset.dataset.available;
      if (!expanded) els.sysPrompt.focus();
    });

    els.sysReset.addEventListener('click', () => {
      const hasContent = els.sysPrompt.value.trim().length > 0;
      if (hasContent && !confirm('Substituir a instrução actual pela versão padrão?')) return;
      loadDefaultSystemPrompt({ force: true });
    });

    els.contextClear.addEventListener('click', () => {
      state.pendingContext = null;
      showContextChip();
    });

    els.clearChat.addEventListener('click', clearChat);
  });
})();
