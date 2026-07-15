// ── Lógica de renderização do slide de fechamento por operadora ──────────
// Compartilhada entre fechamento.html (um slide) e fechamento-completo.html
// (capa + um slide por operadora) — mesma função, mesmo resultado visual,
// sem duplicar HTML/CSS/JS entre as duas telas.

// Paleta fixa por categoria (mesma cor sempre, em qualquer operadora/mês)
const CATEGORY_COLORS = {
  'Agendamento EQ':                 '#03914C',
  'Reenvio de assinatura médica':   '#F15A24',
  'Aguardando assinatura do médico':'#1F5C18',
  'N° Incorreto':                   '#A6A6A6',
  'Reagendamento EQ':               '#D9530A',
  'Modificar Cadastro':             '#8A9586',
  'Link de assinatura':             '#3B6FA0',
  'Cadastro não localizado':        '#B23A48',
  'Outro':                          '#8C8C8C',
}

// Cor determinística (mesma categoria → mesma cor sempre) para tipos futuros
// que ainda não estão no mapa acima — nunca quebra por categoria nova.
function colorFor(tipo) {
  if (CATEGORY_COLORS[tipo]) return CATEGORY_COLORS[tipo]
  let h = 0
  for (const ch of String(tipo)) h = (h * 31 + ch.charCodeAt(0)) >>> 0
  return `hsl(${h % 360} 52% 40%)`
}

function computeDerived(data) {
  const tipos = [...(data.tipos || [])].filter(t => t && t.tipo).sort((a, b) => b.count - a.count)
  const total = data.total ?? tipos.reduce((a, t) => a + t.count, 0)
  const maior = tipos[0] || { tipo: '—', count: 0 }
  const countOf = name => (tipos.find(t => t.tipo === name) || { count: 0 }).count
  const assinatura = countOf('Reenvio de assinatura médica') + countOf('Aguardando assinatura do médico')
  const maxCount = tipos.length ? tipos[0].count : 0
  return { tipos, total, maior, assinatura, maxCount }
}

// Markup de um slide de operadora — instanciado 1x em fechamento.html, ou
// N vezes (um por operadora) em fechamento-completo.html.
function operatorSlideMarkup() {
  return `
    <div class="hd">
      <div class="hd-title">Fechamento</div>
      <div class="hd-logo-wrap"></div>
    </div>
    <div class="card">
      <div class="card-title">Relatório mensal de demandas</div>
      <div class="card-subtitle">—</div>
      <div class="kpi-row">
        <div class="kpi-box">
          <div class="kpi-label">Total</div>
          <div class="kpi-line">
            <span class="kpi-num orange kpi-total-n">0</span>
            <span class="kpi-desc">demandas registradas</span>
          </div>
        </div>
        <div class="kpi-box">
          <div class="kpi-label">Maior demanda</div>
          <div class="kpi-line">
            <span class="kpi-num green kpi-maior-n">0</span>
            <span class="kpi-desc kpi-maior-desc">—</span>
          </div>
        </div>
        <div class="kpi-box">
          <div class="kpi-label">Assinatura médica</div>
          <div class="kpi-line">
            <span class="kpi-num orange kpi-assin-n">0</span>
            <span class="kpi-desc">Reenvio + aguardando assinatura</span>
          </div>
        </div>
      </div>
      <div class="body-row">
        <div class="donut-col">
          <div class="section-title">Distribuição por tipo de demanda</div>
          <div class="donut-wrap"><canvas class="donut" width="386" height="386"></canvas></div>
        </div>
        <div class="table-col">
          <div class="section-title">Volume por tipo de demanda</div>
          <div class="tbl-head"><span>Demanda</span><span>Volume</span><span>Qtd</span><span>%</span></div>
          <div class="tbl-rows"></div>
        </div>
      </div>
    </div>`
}

// Ajusta o tamanho do título até caber na largura disponível.
function fitTitle(el) {
  let size = 50
  el.style.fontSize = size + 'px'
  const maxWidth = el.parentElement ? el.parentElement.clientWidth - 420 - 56 - 62 : 1050
  while (el.scrollWidth > maxWidth && size > 22) {
    size -= 2
    el.style.fontSize = size + 'px'
  }
}

function drawDonut(canvas, derived) {
  const ctx = canvas.getContext('2d')
  const W = canvas.width, H = canvas.height, cx = W / 2, cy = H / 2
  const R = 193, r = 92
  const { tipos, total } = derived
  ctx.clearRect(0, 0, W, H)

  if (!total) {
    ctx.beginPath(); ctx.arc(cx, cy, R, 0, Math.PI * 2)
    ctx.fillStyle = '#EDEAE3'; ctx.fill()
  } else {
    let angle = -Math.PI / 2
    tipos.forEach(t => {
      if (!t.count) return
      const frac = t.count / total, sweep = frac * Math.PI * 2
      ctx.beginPath(); ctx.moveTo(cx, cy)
      ctx.arc(cx, cy, R, angle, angle + sweep); ctx.closePath()
      ctx.fillStyle = colorFor(t.tipo); ctx.fill()
      ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 3; ctx.stroke()

      const pct = Math.round(frac * 100)
      if (pct >= 6) {
        const mid = angle + sweep / 2, lr = (R + r) / 2
        ctx.save()
        ctx.fillStyle = '#fff'; ctx.font = '700 23px Segoe UI,Arial'
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
        ctx.shadowColor = 'rgba(0,0,0,.25)'; ctx.shadowBlur = 3
        ctx.fillText(pct + '%', cx + Math.cos(mid) * lr, cy + Math.sin(mid) * lr)
        ctx.restore()
      }
      angle += sweep
    })
  }

  ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2)
  ctx.fillStyle = '#FFFFFF'; ctx.fill()

  ctx.textAlign = 'center'; ctx.textBaseline = 'alphabetic'
  ctx.fillStyle = '#0B1F4A'
  ctx.font = `800 ${String(total).length >= 3 ? 46 : 54}px Segoe UI,Arial`
  ctx.fillText(String(total), cx, cy + 4)
  ctx.font = '500 20px Segoe UI,Arial'; ctx.fillStyle = '#8F8B82'
  ctx.fillText('demandas', cx, cy + 32)
}

function buildLogo(root, data, logoDataUrl) {
  const wrap = root.querySelector('.hd-logo-wrap')
  if (logoDataUrl) {
    wrap.innerHTML = `<img src="${logoDataUrl}" alt="">`
  } else {
    const label = (data.opDisplay || data.op || 'U').trim()
    wrap.innerHTML = `<div class="hd-logo-fallback">${label}</div>`
  }
}

// Popula um slide (root = elemento .slide, já com operatorSlideMarkup() dentro)
// com os dados de uma operadora. Busca a logo via IPC (com fallback de
// iniciais) e retorna uma Promise que resolve quando tudo — inclusive a
// logo — estiver pronto para captura/impressão.
async function renderOperatorSlide(root, data) {
  const derived = computeDerived(data)

  const titleEl = root.querySelector('.hd-title')
  const nome = data.opDisplay || (data.op || '').replace(/^UNIMED\s+/i, '')
  titleEl.textContent = `Fechamento ${nome} ${data.mes}`
  fitTitle(titleEl)

  root.querySelector('.card-subtitle').textContent = `${data.mes} de ${data.ano}`
  root.querySelector('.kpi-total-n').textContent = derived.total
  root.querySelector('.kpi-maior-n').textContent = derived.maior.count
  root.querySelector('.kpi-maior-desc').textContent = derived.maior.tipo
  root.querySelector('.kpi-assin-n').textContent = derived.assinatura

  root.querySelector('.tbl-rows').innerHTML = derived.tipos.map(t => {
    const color  = colorFor(t.tipo)
    const pct    = derived.total ? Math.round(t.count / derived.total * 100) : 0
    const barPct = derived.maxCount ? Math.round(t.count / derived.maxCount * 100) : 0
    return `<div class="tbl-row">
      <span class="tbl-name" title="${t.tipo}">${t.tipo}</span>
      <div class="tbl-bar-track"><div class="tbl-bar-fill" style="width:${barPct}%;background:${color}"></div></div>
      <span class="tbl-qtd">${t.count}</span>
      <span class="tbl-pct">${pct}%</span>
    </div>`
  }).join('')

  drawDonut(root.querySelector('.donut'), derived)

  const card = root.querySelector('.card')
  const oldBanner = card.querySelector('.empty-banner')
  if (oldBanner) oldBanner.remove()
  if (!derived.total) {
    card.insertAdjacentHTML('beforeend',
      `<div class="empty-banner"><div class="icon">🗂️</div><div class="txt">Sem demandas registradas neste período</div></div>`)
  }

  let logoDataUrl = null
  if (window.csApp && window.csApp.getOperatorLogo) {
    try { logoDataUrl = await window.csApp.getOperatorLogo(data.op || data.opDisplay || '') } catch {}
  }
  buildLogo(root, data, logoDataUrl)
}
