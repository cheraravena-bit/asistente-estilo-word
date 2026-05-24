const analyzeButton = document.getElementById("analyzeButton");
const statusBadge = document.getElementById("status");
const summaryText = document.getElementById("summaryText");
const suggestionsList = document.getElementById("suggestions");

const styleRules = [
  {
    title: "Abuso de adverbios en \"-mente\"",
    description:
      "Los adverbios pueden ralentizar el ritmo del texto y explicar demasiado. Intenta reemplazarlos por verbos o acciones mas precisas que transmitan la idea sin necesidad de reforzarla.",
    objective: "Favorecer una escritura mas dinamica, visual y directa.",
    findMatches: findMenteAdverbs
  },
  {
    title: "Uso de voz pasiva",
    description:
      "La voz pasiva crea distancia entre el lector y la accion. Prioriza la voz activa para lograr frases mas claras, directas y con mayor impacto narrativo.",
    objective: "Mejorar claridad, ritmo y cercania del relato.",
    findMatches: findPassiveVoice
  },
  {
    title: "Repeticiones lexicas cercanas",
    description:
      "La repeticion constante puede volver el texto monotono. Evalua si puedes variar el vocabulario, reorganizar la frase o reemplazar terminos sin perder tu voz.",
    objective: "Aumentar fluidez y riqueza linguistica sin sobrecorregir el estilo personal.",
    findMatches: findNearbyRepetitions
  },
  {
    title: "Oraciones excesivamente largas",
    description:
      "Las oraciones demasiado extensas pueden dificultar la lectura y diluir la idea principal. Considera dividir la informacion en dos frases o jerarquizar mejor las ideas.",
    objective: "Mejorar legibilidad y comprension del texto.",
    findMatches: findLongSentences
  },
  {
    title: "Cliches o frases demasiado usadas",
    description:
      "Los cliches reducen originalidad y fuerza expresiva. Busca una imagen, comparacion o formulacion mas propia de tu voz narrativa.",
    objective: "Potenciar autenticidad y construccion de estilo autoral.",
    findMatches: findCliches
  }
];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    analyzeButton.disabled = false;
    analyzeButton.addEventListener("click", analyzeDocument);
    return;
  }

  analyzeButton.disabled = true;
  setStatus("Solo Word");
  summaryText.textContent = "Este prototipo esta pensado para ejecutarse dentro de Microsoft Word.";
});

async function analyzeDocument() {
  setLoading(true);
  clearSuggestions();

  try {
    const text = await readDocumentText();
    const suggestions = buildSuggestions(text);
    renderSuggestions(suggestions, text);
  } catch (error) {
    summaryText.textContent = "No se pudo leer el documento. Revisa que Word haya cargado el complemento.";
    console.error(error);
  } finally {
    setLoading(false);
  }
}

async function readDocumentText() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text.trim();
  });
}

function buildSuggestions(text) {
  return styleRules.flatMap((rule) =>
    rule.findMatches(text).map((match) => ({
      title: rule.title,
      description: rule.description,
      objective: rule.objective,
      quote: match
    }))
  );
}

function renderSuggestions(suggestions, text) {
  if (!text) {
    summaryText.textContent = "El documento no tiene texto para analizar.";
    return;
  }

  if (suggestions.length === 0) {
    summaryText.textContent =
      "No hay sugerencias todavia. Completa las 5 reglas de estilo para activar el analisis.";
    return;
  }

  summaryText.textContent = `${suggestions.length} sugerencia(s) encontradas. El documento no fue modificado.`;

  for (const suggestion of suggestions) {
    const item = document.createElement("li");
    item.className = "suggestion";
    item.innerHTML = `
      <strong>${escapeHtml(suggestion.title)}</strong>
      <p>${escapeHtml(suggestion.description)}</p>
      <p><strong>Objetivo:</strong> ${escapeHtml(suggestion.objective)}</p>
      <p class="quote">${escapeHtml(suggestion.quote)}</p>
    `;
    suggestionsList.appendChild(item);
  }
}

function findMenteAdverbs(text) {
  const matches = text.match(/\b[a-zA-ZáéíóúÁÉÍÓÚñÑüÜ]+mente\b/gi) || [];
  return unique(matches).slice(0, 5).map((word) => `Adverbio detectado: "${word}"`);
}

function findPassiveVoice(text) {
  const passivePattern =
    /\b(?:fue|fueron|era|eran|ha sido|han sido|habia sido|habian sido)\s+[a-zA-ZáéíóúÁÉÍÓÚñÑüÜ]+(?:ado|ada|ados|adas|ido|ida|idos|idas)\s+por\b/gi;
  return unique(text.match(passivePattern) || [])
    .slice(0, 5)
    .map((phrase) => `Posible voz pasiva: "${phrase}"`);
}

function findNearbyRepetitions(text) {
  const paragraphs = text.split(/\n+/).filter(Boolean);
  const suggestions = [];

  for (const paragraph of paragraphs) {
    const words = normalizeWords(paragraph).filter((word) => word.length > 4);
    const counts = new Map();

    for (const word of words) {
      counts.set(word, (counts.get(word) || 0) + 1);
    }

    for (const [word, count] of counts) {
      if (count >= 3) {
        suggestions.push(`Repeticion cercana: "${word}" aparece ${count} veces en un parrafo.`);
      }
    }
  }

  return suggestions.slice(0, 5);
}

function findLongSentences(text) {
  return splitSentences(text)
    .map((sentence) => sentence.trim())
    .filter(Boolean)
    .filter((sentence) => {
      const wordCount = normalizeWords(sentence).length;
      const hasClearPause = /[,;:]/.test(sentence);
      return wordCount > 40 || (wordCount > 35 && !hasClearPause);
    })
    .slice(0, 5)
    .map((sentence) => `Oracion extensa: "${extractSample(sentence)}"`);
}

function findCliches(text) {
  const cliches = [
    "al final del dia",
    "sin lugar a dudas",
    "desde tiempos inmemoriales",
    "una montana rusa de emociones",
    "una montaña rusa de emociones"
  ];
  const normalizedText = removeDiacritics(text).toLowerCase();

  return cliches
    .filter((phrase) => normalizedText.includes(removeDiacritics(phrase).toLowerCase()))
    .map((phrase) => `Cliche detectado: "${phrase}"`);
}

function extractSample(text) {
  const compact = text.replace(/\s+/g, " ").trim();
  if (compact.length <= 140) {
    return compact;
  }

  return `${compact.slice(0, 137)}...`;
}

function normalizeWords(text) {
  return removeDiacritics(text)
    .toLowerCase()
    .match(/\b[a-zñü]{2,}\b/g) || [];
}

function removeDiacritics(value) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function splitSentences(text) {
  return text
    .replace(/([.!?])\s+/g, "$1|")
    .split("|");
}

function unique(values) {
  return [...new Set(values)];
}

function clearSuggestions() {
  suggestionsList.replaceChildren();
}

function setLoading(isLoading) {
  analyzeButton.disabled = isLoading;
  analyzeButton.textContent = isLoading ? "Analizando..." : "Analizar documento";
  setStatus(isLoading ? "Leyendo" : "Listo");
}

function setStatus(label) {
  statusBadge.textContent = label;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
