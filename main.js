// CONFIGURAÇÃO: Coloca aqui o link do teu Google Sheets publicado como CSV
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTMBrUrS0Sz40SavUMVWWeeUUp4EJLHqtefxgnYWpYLSCIvzx5fy0J5v2MjEnugZ5Vd0ylNyYnw8Ptc/pub?output=csv";

let catalogProducts = [];
let carouselResizeBound = false;

// --- ANIMAÇÕES E INTERFACE (Mantido do teu original) ---

function setupRevealAnimations() {
  const revealElements = document.querySelectorAll(".reveal");
  if (!revealElements.length) return;

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("visible");
        }
      });
    },
    { threshold: 0.15 }
  );

  revealElements.forEach((section) => observer.observe(section));
}

// --- LÓGICA DE TRATAMENTO DE DADOS (Mantido e Otimizado) ---

function sanitizeKey(key) {
  return String(key)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function normalizeRecord(raw) {
  const normalized = {};
  Object.keys(raw).forEach((key) => {
    normalized[sanitizeKey(key)] = raw[key];
  });

  const image = normalized.imagem || normalized.image || normalized.foto || "";
  const reference = normalized.referencia || normalized.ref || normalized.codigo || "";
  const description = normalized.descricao || normalized.description || "";
  const priceRaw = normalized.preco || normalized.price || "";
  const highlightRaw = normalized.destaque || normalized.highlight || "";
  const category = normalized.categoria || normalized.category || "";
  const color = normalized.cor || normalized.color || "";

  return {
    imagem: String(image || "").trim(),
    referencia: String(reference || "").trim(),
    descricao: String(description || "").trim(),
    categoria: String(category || "").trim(),
    cor: String(color || "").trim(),
    preco: normalizePrice(priceRaw),
    destaque: normalizeBoolean(highlightRaw)
  };
}

function normalizeBoolean(value) {
  const text = String(value || "").toLowerCase().trim();
  return ["1", "sim", "yes", "true", "x", "destaque"].includes(text);
}

function normalizePrice(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  let text = String(value || "").replace(/\s/g, "").replace(/[^\d,.-]/g, "");
  if (text.includes(",") && text.includes(".")) {
    text = text.lastIndexOf(",") > text.lastIndexOf(".") ? text.replace(/\./g, "").replace(",", ".") : text.replace(/,/g, "");
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  }
  const parsed = Number.parseFloat(text);
  return Number.isFinite(parsed) ? parsed : value;
}

function formatPrice(value) {
  const parsed = typeof value === "number" ? value : normalizePrice(value);
  if (typeof parsed === "number" && Number.isFinite(parsed)) {
    return new Intl.NumberFormat("pt-PT", { style: "currency", currency: "EUR" }).format(parsed);
  }
  return "Sob consulta";
}

function escapeHtml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;").replace(/'/g, "&#039;");
}

// --- CARREGAMENTO DE DADOS (A grande mudança para o Vercel/Gemini) ---

async function loadProductsFromExcel() {
  // Buscamos os dados diretamente do Google Sheets (CSV)
  // O cache: "no-store" garante que o site atualize assim que a tua mãe falar com o Gemini
  const response = await fetch(SHEET_CSV_URL, { cache: "no-store" });
  
  if (!response.ok) {
    throw new Error("Erro ao aceder à base de dados do Google Sheets.");
  }

  const csvText = await response.text();
  
  // A biblioteca XLSX converte o CSV para o formato que o teu site já entende
  const workbook = XLSX.read(csvText, { type: "string" });
  const firstSheet = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });

  return rows.map(normalizeRecord).filter((item) => item.referencia || item.descricao);
}

// --- RENDERIZAÇÃO (Mantido do teu original) ---

function renderHighlights(products) {
  const gallery = document.getElementById("highlight-gallery");
  if (!gallery) return;

  const highlighted = products.filter((item) => item.destaque);
  const selected = highlighted.length ? highlighted : products.slice(0, 3);
  const itemsPerPage = window.innerWidth < 640 ? 2 : (window.innerWidth < 1180 ? 3 : 4);

  if (!selected.length) {
    gallery.innerHTML = "<p class=\"empty-state\">Sem destaques disponíveis.</p>";
    return;
  }

  const chunks = [];
  for (let i = 0; i < selected.length; i += itemsPerPage) chunks.push(selected.slice(i, i + itemsPerPage));

  gallery.innerHTML = `
    <div class="highlight-viewport">
      <div class="highlight-track">
        ${chunks.map(page => `
          <div class="highlight-slide">
            <div class="highlight-page page-size-${itemsPerPage}">
              ${page.map(p => `
                <article class="highlight-card">
                  <img src="${escapeHtml(p.imagem || 'img/placeholder.jpg')}" alt="${escapeHtml(p.referencia)}" />
                  <div class="highlight-info">
                    <p class="highlight-kicker">Destaque</p>
                    <h3>${escapeHtml(p.referencia)}</h3>
                    <p>${escapeHtml(p.descricao)}</p>
                    <span class="price">${formatPrice(p.preco)}</span>
                  </div>
                </article>
              `).join('')}
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
}

function renderProductGrid(products) {
  const grid = document.getElementById("product-grid");
  if (!grid) return;

  if (!products.length) {
    grid.innerHTML = '<p class="empty-state">Nenhum produto encontrado na base de dados.</p>';
    return;
  }

  grid.innerHTML = products.map(p => `
    <article class="product-card">
      <img class="product-media" src="${escapeHtml(p.imagem || 'img/placeholder.jpg')}" alt="${escapeHtml(p.referencia)}" />
      <div class="product-content">
        <p class="product-ref">Ref: ${escapeHtml(p.referencia)}</p>
        <p>${escapeHtml(p.descricao)}</p>
        <p class="product-category">${escapeHtml(p.categoria || "")}</p>
        <p class="product-color">${escapeHtml(p.cor || "")}</p>
        <span class="price">${formatPrice(p.preco)}</span>
      </div>
    </article>
  `).join("");
}

function applyFilters() {
  const categorySelect = document.getElementById("categoria-filter");
  const colorSelect = document.getElementById("cor-filter");

  if (!categorySelect || !colorSelect) {
    renderProductGrid(catalogProducts);
    return;
  }

  const selectedCategory = categorySelect.value;
  const selectedColor = colorSelect.value;

  const filteredProducts = catalogProducts.filter((product) => {
    const categoryMatch = !selectedCategory || product.categoria === selectedCategory;
    const colorMatch = !selectedColor || product.cor === selectedColor;

    return categoryMatch && colorMatch;
  });

  renderProductGrid(filteredProducts);
}

function setupFilters() {
  const grid = document.getElementById("product-grid");

  if (!grid) return;

  const filtersHTML = `
    <div class="catalog-filters">
      <select id="categoria-filter">
        <option value="">Todas as categorias</option>
        <option value="Colares curtos">Colares curtos</option>
        <option value="Colares compridos">Colares compridos</option>
        <option value="Pulseiras">Pulseiras</option>
        <option value="Brincos">Brincos</option>
        <option value="Anéis">Anéis</option>
      </select>

      <select id="cor-filter">
        <option value="">Todas as cores</option>
        <option value="Dourado">Dourado</option>
        <option value="Prateado">Prateado</option>
      </select>
    </div>
  `;

  grid.insertAdjacentHTML("beforebegin", filtersHTML);

  document
    .getElementById("categoria-filter")
    .addEventListener("change", applyFilters);

  document
    .getElementById("cor-filter")
    .addEventListener("change", applyFilters);
}

async function initCatalog() {
  const note = document.getElementById("data-note");
  try {
    const products = await loadProductsFromExcel();
    catalogProducts = products;
    setupFilters();
    renderHighlights(products);
    renderProductGrid(products);
    if(note) note.textContent = `Catálogo atualizado: ${products.length} produtos.`;
  } catch (error) {
    console.error(error);
    if(note) {
      note.textContent = "Erro ao ligar ao Google Sheets. Verifica o link e a publicação.";
      note.classList.add("error");
    }
  }
}

// Inicialização
setupRevealAnimations();
initCatalog();

// Resize do Carousel
window.addEventListener("resize", () => {
  if (catalogProducts.length) renderHighlights(catalogProducts);
});