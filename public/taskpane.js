// VivWord — sidebar de revisão literária com Claude dentro do Word.
// Vanilla JS. Sem build. Targets WordApi 1.4 (margin comments + chat).

(function () {
  'use strict';

  // ---------- Estado de sessão ----------
  const state = {
    messages: [],          // [{ role, content, displayContent? }]
    pendingContext: null,  // { kind: 'selection'|'document', text, words }
    lastAssistant: '',
    inFlight: false,
    skills: [],
    canComment: false,
    inWord: false,
  };

  const WORDS_WARN = 50000;
  const SETTINGS_KEYS = {
    history: 'vivword.history',
    voice: 'vivword.voice',
    model: 'vivword.model',
    sysOverride: 'vivword.sysOverride',
  };

  // ---------- DOM refs ----------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));
  const els = {
    chat: $('#chat'),
    input: $('#input'),
    composer: $('#composer'),
    sendBtn: $('#sendBtn'),
    loading: $('#loading'),
    model: $('#model'),
    sysToggle: $('#sysToggle'),
    sysPrompt: $('#sysPrompt'),
    sysReset: $('#sysReset'),
    sysBody: $('#sysToggle + .sys-body'),
    voiceToggle: $('#voiceToggle'),
    voiceFingerprint: $('#voiceFingerprint'),
    voiceStatus: $('#voiceStatus'),
    voiceBody: $('#voiceToggle + .sys-body'),
    skills: $('#skills'),
    contextChip: $('#contextChip'),
    contextLabel: $('#contextLabel'),
    contextClear: $('#contextClear'),
    clearChat: $('#clearChat'),
    exportChat: $('#exportChat'),
    commentBtn: $('.tool[data-action="comment"]'),
    app: $('#app'),
  };

  // ---------- Helpers de Office ----------
  function officeAvailable() {
    return typeof Office !== 'undefined' && typeof Word !== 'undefined';
  }

  function isApiSupported(set, version) {
    try {
      return !!(Office.context && Office.context.requirements &&
        Office.context.requirements.isSetSupported(set, version));
    } catch (_) { return false; }
  }

  function settingsAvailable() {
    return officeAvailable() && Office.context && Office.context.document &&
      Office.context.document.settings;
  }

  function settingsGet(key) {
    if (!settingsAvailable()) return null;
    try { return Office.context.document.settings.get(key); } catch (_) { return null; }
  }

  function settingsSet(key, value) {
    if (!settingsAvailable()) return;
    try {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync(() => {});
    } catch (_) {}
  }

  // Debounce factory para autosave.
  function debounce(fn, ms) {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  }

  // ---------- Word actions (WordApi 1.4) ----------
  function ensureWord() {
    if (!officeAvailable()) {
      throw new Error('Office.js indisponível. Abre dentro do Word.');
    }
  }

  // Wrappers do Office. A API legada (Common API: getSelectedDataAsync,
  // setSelectedDataAsync) é muito mais estável no iPad Safari + Word web do
  // que `Word.run + getSelection`. Usamos a legada para selecção e para
  // inserção/substituição. A API moderna (Word.run) só é usada onde a legada
  // não chega: ler o documento todo e inserir comentário na margem.

  function getSelectedTextAsync() {
    return new Promise((resolve, reject) => {
      if (!state.inWord) {
        reject(new Error('Esta acção precisa do Word à volta. Abre o add-in dentro de word.office.com (ou no app Word) — não funciona na pré-visualização.'));
        return;
      }
      try {
        Office.context.document.getSelectedDataAsync(
          Office.CoercionType.Text,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || '');
            } else {
              const msg = (result.error && result.error.message) || 'Falha ao ler selecção.';
              reject(new Error(msg));
            }
          }
        );
      } catch (e) {
        reject(e);
      }
    });
  }

  function setSelectedTextAsync(text) {
    return new Promise((resolve, reject) => {
      if (!state.inWord) {
        reject(new Error('Esta acção precisa do Word à volta.'));
        return;
      }
      try {
        Office.context.document.setSelectedDataAsync(
          text,
          { coercionType: Office.CoercionType.Text },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error((result.error && result.error.message) || 'Falha ao inserir.'));
          }
        );
      } catch (e) {
        reject(e);
      }
    });
  }

  async function readSelection() {
    return getSelectedTextAsync();
  }

  async function readDocument() {
    ensureWord();
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      return body.text || '';
    });
  }

  async function insertAtCursor(text) {
    return setSelectedTextAsync(text);
  }

  async function commentOnSelection(text) {
    ensureWord();
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertComment(text);
      await context.sync();
    });
  }

  // ---------- Render ----------
  function clearEmptyHint() {
    const hint = els.chat.querySelector('.empty-hint');
    if (hint) hint.remove();
  }

  function appendMessage(role, text, opts = {}) {
    clearEmptyHint();
    const div = document.createElement('div');
    const cls = role === 'user' ? 'user'
              : role === 'error' ? 'error'
              : role === 'system-note' ? 'system-note'
              : 'claude';
    div.className = 'msg ' + cls;
    div.textContent = text;
    if (opts.streaming) div.classList.add('streaming');
    if (opts.meta) {
      const m = document.createElement('span');
      m.className = 'meta';
      m.textContent = opts.meta;
      div.appendChild(m);
    }
    els.chat.appendChild(div);
    els.chat.scrollTop = els.chat.scrollHeight;
    return div;
  }

  function renderHistory() {
    els.chat.innerHTML = '';
    if (!state.messages.length) {
      const hint = document.createElement('div');
      hint.className = 'empty-hint';
      hint.innerHTML = '<p>Olá. Pergunta algo, usa os botões para passar texto do documento, ou toca uma <em>skill</em>.</p>';
      els.chat.appendChild(hint);
      return;
    }
    for (const m of state.messages) {
      const text = m.displayContent || m.content;
      appendMessage(m.role === 'assistant' ? 'claude' : 'user', text);
    }
  }

  function setLoading(on) {
    state.inFlight = on;
    els.loading.hidden = !on;
    els.sendBtn.disabled = on;
    els.app.setAttribute('aria-busy', on ? 'true' : 'false');
    $$('.skill').forEach((b) => { b.disabled = on; });
    if (on) els.chat.scrollTop = els.chat.scrollHeight;
  }

  function showContextChip() {
    if (!state.pendingContext) {
      els.contextChip.hidden = true;
      els.contextLabel.textContent = '';
      return;
    }
    const c = state.pendingContext;
    const noun = c.kind === 'selection' ? 'Selecção' : 'Documento';
    els.contextLabel.textContent = `${noun} anexo (${c.words.toLocaleString('pt-PT')} palavra${c.words === 1 ? '' : 's'})`;
    els.contextChip.hidden = false;
  }

  function countWords(s) {
    if (!s) return 0;
    const m = s.trim().match(/\S+/g);
    return m ? m.length : 0;
  }

  // ---------- Persistência ----------
  function persistHistory() {
    settingsSet(SETTINGS_KEYS.history, JSON.stringify(state.messages.slice(-40)));
  }
  const persistHistoryDebounced = debounce(persistHistory, 500);

  function persistVoice() {
    settingsSet(SETTINGS_KEYS.voice, els.voiceFingerprint.value || '');
    flashVoiceStatus('guardado');
  }
  const persistVoiceDebounced = debounce(persistVoice, 700);

  function flashVoiceStatus(text) {
    els.voiceStatus.textContent = text;
    setTimeout(() => { els.voiceStatus.textContent = ''; }, 1400);
  }

  function loadFromSettings() {
    const hist = settingsGet(SETTINGS_KEYS.history);
    if (hist) {
      try {
        const parsed = JSON.parse(hist);
        if (Array.isArray(parsed)) {
          state.messages = parsed;
          const last = [...parsed].reverse().find((m) => m.role === 'assistant');
          if (last) state.lastAssistant = last.content;
        }
      } catch (_) {}
    }
    const voice = settingsGet(SETTINGS_KEYS.voice);
    if (voice) els.voiceFingerprint.value = voice;
    const model = settingsGet(SETTINGS_KEYS.model);
    if (model) els.model.value = model;
    const sysOverride = settingsGet(SETTINGS_KEYS.sysOverride);
    if (sysOverride) els.sysPrompt.value = sysOverride;
  }

  // ---------- System prompt + voz ----------
  let defaultSystemLoaded = false;

  async function loadDefaultSystemPrompt({ force = false } = {}) {
    try {
      const r = await fetch('default-system-prompt.txt', { cache: 'no-store' });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const text = (await r.text()).trim();
      if (!text) return;
      if (force || !els.sysPrompt.value.trim()) {
        els.sysPrompt.value = text;
        // Reposição apaga o override persistido para voltar a seguir o default.
        if (force) settingsSet(SETTINGS_KEYS.sysOverride, '');
      }
      defaultSystemLoaded = true;
      els.sysReset.dataset.available = '1';
      const expanded = els.sysToggle.getAttribute('aria-expanded') === 'true';
      els.sysReset.hidden = !expanded;
      els.sysPrompt.placeholder = 'Edita ou repõe a instrução padrão.';
    } catch (e) {
      els.sysPrompt.placeholder = 'Ex.: És revisor literário. Mantém o estilo da autora.';
      if (typeof console !== 'undefined') console.warn('default-system-prompt:', e.message);
    }
  }

  function buildEffectiveSystem() {
    const base = els.sysPrompt.value.trim();
    const voice = els.voiceFingerprint.value.trim();
    if (!voice) return base;
    const voiceBlock = `\n\n## Voz canónica desta obra\n\nO que se segue são parágrafos que a autora reconhece como inquestionavelmente da sua voz. Usa-os como referência calibradora — não os cites de volta, não os repitas, apenas mede o que ela escreve por aqui.\n\n---\n${voice}\n---`;
    return (base ? base : '') + voiceBlock;
  }

  // ---------- Skills ----------
  async function loadSkills() {
    try {
      const r = await fetch('skills/index.json', { cache: 'no-store' });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      state.skills = await r.json();
      renderSkills();
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('skills index:', e.message);
    }
  }

  function renderSkills() {
    els.skills.innerHTML = '';
    for (const sk of state.skills) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'skill';
      btn.dataset.id = sk.id;
      const needsHost = sk.context === 'selection' || sk.context === 'document';
      if (needsHost && !state.inWord) {
        btn.disabled = true;
        btn.title = 'Disponível apenas dentro do Word';
      } else {
        btn.title = sk.tip || sk.label;
      }
      btn.textContent = sk.label;
      btn.addEventListener('click', () => runSkill(sk, btn));
      els.skills.appendChild(btn);
    }
  }

  async function runSkill(sk, btn) {
    if (state.inFlight) return;
    btn.classList.add('running');
    try {
      // Preparar contexto: lê a fonte indicada pela skill.
      let ctxText = '';
      let ctxWords = 0;
      if (sk.context === 'selection') {
        ctxText = await readSelection();
        if (!ctxText.trim()) {
          appendMessage('error', 'Skill cancelada: nada seleccionado no documento.');
          return;
        }
      } else if (sk.context === 'document') {
        ctxText = await readDocument();
        if (!ctxText.trim()) {
          appendMessage('error', 'Skill cancelada: documento vazio.');
          return;
        }
      }
      ctxWords = countWords(ctxText);

      // Carregar prompt da skill.
      const promptText = await fetchSkillPrompt(sk);
      if (!promptText) {
        appendMessage('error', `Skill "${sk.label}" sem prompt definido.`);
        return;
      }

      // Anexar contexto e enviar.
      if (ctxText) {
        state.pendingContext = { kind: sk.context, text: ctxText, words: ctxWords };
        showContextChip();
      }
      appendMessage('system-note', `▸ skill "${sk.label}" sobre ${sk.context === 'selection' ? `selecção (${ctxWords} palavra${ctxWords === 1 ? '' : 's'})` : `documento (${ctxWords.toLocaleString('pt-PT')} palavras)`}`);
      if (ctxWords > WORDS_WARN) {
        appendMessage('system-note', `aviso: documento grande (>${WORDS_WARN.toLocaleString('pt-PT')} palavras). Pode demorar.`);
      }
      await dispatchUserMessage(promptText);
    } catch (e) {
      appendMessage('error', `Erro na skill: ${e.message}`);
    } finally {
      btn.classList.remove('running');
    }
  }

  const skillPromptCache = new Map();
  async function fetchSkillPrompt(sk) {
    if (skillPromptCache.has(sk.id)) return skillPromptCache.get(sk.id);
    const r = await fetch(`skills/${sk.file}`, { cache: 'no-store' });
    if (!r.ok) return null;
    const text = (await r.text()).trim();
    skillPromptCache.set(sk.id, text);
    return text;
  }

  // ---------- Toolbar actions ----------
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
      appendMessage('system-note', `selecção lida — ${words} palavra${words === 1 ? '' : 's'} anexa${words === 1 ? '' : 's'} à próxima mensagem`);
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
      appendMessage('system-note', `documento lido — ${words.toLocaleString('pt-PT')} palavras anexas à próxima mensagem`);
      if (words > WORDS_WARN) {
        appendMessage('system-note', `aviso: documento grande (>${WORDS_WARN.toLocaleString('pt-PT')} palavras). Pode demorar e consumir muitos tokens.`);
      }
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
      appendMessage('system-note', 'resposta inserida no documento');
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
      await insertAtCursor(state.lastAssistant);
      appendMessage('system-note', 'selecção substituída');
    } catch (e) {
      appendMessage('error', 'Erro ao substituir: ' + e.message);
    }
  }

  async function actionComment() {
    if (!state.canComment) {
      appendMessage('error', 'O Word desta máquina não suporta comentários por API.');
      return;
    }
    if (!state.lastAssistant) {
      appendMessage('error', 'Ainda não há resposta do Claude para comentar.');
      return;
    }
    try {
      await commentOnSelection(state.lastAssistant);
      appendMessage('system-note', 'comentário do Claude inserido na margem');
    } catch (e) {
      appendMessage('error', 'Erro ao comentar: ' + e.message);
    }
  }

  // ---------- Chat: composer ↔ API ----------
  function buildMessagesForApi(userText) {
    let composed = userText;
    if (state.pendingContext) {
      const c = state.pendingContext;
      const label = c.kind === 'selection' ? 'SELECÇÃO DO DOCUMENTO' : 'DOCUMENTO COMPLETO';
      composed = `[${label}]\n${c.text}\n[/${label}]\n\n${userText}`;
    }
    const history = state.messages.map((m) => ({ role: m.role, content: m.content }));
    history.push({ role: 'user', content: composed });
    return { apiMessages: history, composed };
  }

  async function sendStreaming(payload, bubble) {
    const r = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ ...payload, stream: true }),
    });
    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try { const j = await r.json(); if (j && j.error) msg = j.error; } catch (_) {}
      throw new Error(msg);
    }
    if (!r.body || !r.body.getReader) {
      // Fallback se o ambiente não suporta streams: lê tudo de uma vez.
      const text = await r.text();
      return parseSseAccumulated(text, bubble);
    }
    const reader = r.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let assembled = '';
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const events = buffer.split('\n\n');
      buffer = events.pop() || '';
      for (const ev of events) {
        const piece = parseSseEvent(ev);
        if (piece) {
          assembled += piece;
          bubble.firstChild ? (bubble.firstChild.nodeValue = assembled) : (bubble.textContent = assembled);
          els.chat.scrollTop = els.chat.scrollHeight;
        }
      }
    }
    if (buffer.trim()) {
      const piece = parseSseEvent(buffer);
      if (piece) {
        assembled += piece;
        bubble.textContent = assembled;
      }
    }
    return assembled;
  }

  function parseSseEvent(raw) {
    // Anthropic envia: event: <name>\ndata: {...}
    let dataLine = '';
    for (const line of raw.split('\n')) {
      if (line.startsWith('data:')) dataLine += line.slice(5).trim();
    }
    if (!dataLine) return '';
    if (dataLine === '[DONE]') return '';
    try {
      const j = JSON.parse(dataLine);
      if (j.type === 'content_block_delta' && j.delta && j.delta.type === 'text_delta') {
        return j.delta.text || '';
      }
    } catch (_) {}
    return '';
  }

  function parseSseAccumulated(text, bubble) {
    let assembled = '';
    for (const ev of text.split('\n\n')) {
      const piece = parseSseEvent(ev);
      if (piece) assembled += piece;
    }
    bubble.textContent = assembled;
    return assembled;
  }

  async function sendNonStreaming(payload) {
    const r = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try { const j = await r.json(); if (j && j.error) msg = j.error; } catch (_) {}
      throw new Error(msg);
    }
    const data = await r.json();
    const blocks = Array.isArray(data.content) ? data.content : [];
    const text = blocks.filter((b) => b && b.type === 'text').map((b) => b.text).join('\n').trim();
    if (!text) throw new Error('Resposta vazia da Anthropic.');
    return text;
  }

  async function dispatchUserMessage(userText) {
    if (state.inFlight) return;
    if (!userText.trim() && !state.pendingContext) return;

    const effective = userText.trim() || '(Analisa o texto anexo segundo a instrução.)';
    const { apiMessages, composed } = buildMessagesForApi(effective);

    appendMessage('user', effective);
    state.messages.push({ role: 'user', content: composed, displayContent: effective });
    persistHistoryDebounced();

    setLoading(true);
    const bubble = appendMessage('claude', '', { streaming: true });
    try {
      const payload = {
        model: els.model.value,
        messages: apiMessages,
        system: buildEffectiveSystem(),
      };
      let reply;
      try {
        reply = await sendStreaming(payload, bubble);
      } catch (streamErr) {
        // Fallback gracioso para não-streaming.
        bubble.textContent = '';
        reply = await sendNonStreaming({ ...payload, stream: false });
        bubble.textContent = reply;
      }
      bubble.classList.remove('streaming');
      if (!reply || !reply.trim()) {
        bubble.remove();
        appendMessage('error', 'Resposta vazia da Anthropic.');
        return;
      }
      state.messages.push({ role: 'assistant', content: reply });
      state.lastAssistant = reply;
      state.pendingContext = null;
      showContextChip();
      persistHistoryDebounced();
    } catch (e) {
      bubble.remove();
      appendMessage('error', e.message || String(e));
    } finally {
      setLoading(false);
    }
  }

  async function handleSend() {
    const userText = els.input.value;
    if (!userText.trim() && !state.pendingContext) return;
    els.input.value = '';
    autoresize();
    await dispatchUserMessage(userText);
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
    persistHistory();
    renderHistory();
  }

  function exportConversationMarkdown() {
    if (!state.messages.length) {
      appendMessage('system-note', 'nada para exportar');
      return;
    }
    const lines = ['# Conversa VivWord', ''];
    for (const m of state.messages) {
      const text = m.displayContent || m.content;
      lines.push(`## ${m.role === 'assistant' ? 'Claude' : 'Vivianne'}`);
      lines.push('');
      lines.push(text);
      lines.push('');
    }
    const md = lines.join('\n');
    const fallback = () => {
      const ta = document.createElement('textarea');
      ta.value = md;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); } catch (_) {}
      document.body.removeChild(ta);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(md).then(
        () => appendMessage('system-note', 'conversa copiada como Markdown — cola onde quiseres'),
        () => { fallback(); appendMessage('system-note', 'conversa copiada (fallback)'); }
      );
    } else {
      fallback();
      appendMessage('system-note', 'conversa copiada como Markdown');
    }
  }

  function bindToggle(toggleEl, bodyEl, onExpand) {
    toggleEl.addEventListener('click', () => {
      const expanded = toggleEl.getAttribute('aria-expanded') === 'true';
      toggleEl.setAttribute('aria-expanded', expanded ? 'false' : 'true');
      bodyEl.hidden = expanded;
      if (!expanded && typeof onExpand === 'function') onExpand();
    });
  }

  // ---------- Boot ----------
  function bindEvents() {
    $$('.tool').forEach((btn) => {
      btn.addEventListener('click', () => {
        const action = btn.dataset.action;
        if (action === 'read-selection') actionReadSelection();
        else if (action === 'read-document') actionReadDocument();
        else if (action === 'insert') actionInsert();
        else if (action === 'replace') actionReplace();
        else if (action === 'comment') actionComment();
      });
    });

    els.composer.addEventListener('submit', (e) => {
      e.preventDefault();
      handleSend();
    });

    els.input.addEventListener('input', autoresize);
    els.input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    });

    bindToggle(els.sysToggle, els.sysBody, () => {
      els.sysPrompt.focus();
      if (els.sysReset.dataset.available) els.sysReset.hidden = false;
    });
    bindToggle(els.voiceToggle, els.voiceBody, () => els.voiceFingerprint.focus());

    const persistSysOverrideDebounced = debounce(() => {
      settingsSet(SETTINGS_KEYS.sysOverride, els.sysPrompt.value);
    }, 700);
    els.sysPrompt.addEventListener('input', persistSysOverrideDebounced);
    els.voiceFingerprint.addEventListener('input', () => {
      persistVoiceDebounced();
    });
    els.model.addEventListener('change', () => {
      settingsSet(SETTINGS_KEYS.model, els.model.value);
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

    els.clearChat.addEventListener('click', () => {
      if (state.messages.length && !confirm('Limpar a conversa deste documento?')) return;
      clearChat();
    });

    els.exportChat.addEventListener('click', exportConversationMarkdown);
  }

  function detectCapabilities() {
    state.canComment = state.inWord && isApiSupported('WordApi', '1.4');
    if (els.commentBtn) {
      els.commentBtn.hidden = !state.canComment;
    }
    // Desactiva visualmente as acções que precisam do Word à volta.
    $$('.tool').forEach((btn) => {
      const action = btn.dataset.action;
      const needsHost = ['read-selection', 'read-document', 'insert', 'replace', 'comment'].includes(action);
      if (needsHost && !state.inWord) {
        btn.disabled = true;
        btn.title = 'Disponível apenas dentro do Word';
      }
    });
    if (!state.inWord) {
      const note = document.createElement('div');
      note.className = 'msg system-note';
      note.textContent = 'Modo pré-visualização: estás a abrir o taskpane fora do Word. O chat funciona, mas as acções e skills que tocam no documento precisam do Word à volta. Abre via word.office.com com o add-in carregado para usares tudo.';
      els.chat.appendChild(note);
    }
  }

  function boot() {
    bindEvents();
    if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
      Office.onReady((info) => {
        const HostType = (typeof Office !== 'undefined' && Office.HostType) || {};
        state.inWord = !!(info && info.host && HostType.Word && info.host === HostType.Word);
        detectCapabilities();
        if (state.inWord) loadFromSettings();
        loadDefaultSystemPrompt();
        loadSkills();
        renderHistory();
      });
    } else {
      // Office.js falhou a carregar. Modo standalone básico.
      loadDefaultSystemPrompt();
      loadSkills();
      renderHistory();
      detectCapabilities();
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
