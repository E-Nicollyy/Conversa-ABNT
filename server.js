// ==========================================
// ASSISTENTE ACADÊMICO - servidor Node/Express
// Migrado de Python (Colab + Gemini + Gradio) para JS
// ==========================================
require("dotenv").config();

const path = require("path");
const fs = require("fs");
const express = require("express");
const multer = require("multer");
const mammoth = require("mammoth");
const pdfParse = require("pdf-parse");

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Header,
  PageNumber,
  LevelFormat,
  convertMillimetersToTwip,
} = require("docx");

// ==========================================
// CONFIGURAÇÃO DA API (Azure AI Foundry)
// ==========================================
const AZURE_AI_ENDPOINT = (process.env.AZURE_AI_ENDPOINT || "").replace(/\/+$/, "");
const AZURE_AI_KEY = process.env.AZURE_AI_KEY;
const AZURE_AI_MODEL = process.env.AZURE_AI_MODEL;

if (!AZURE_AI_ENDPOINT || !AZURE_AI_KEY || !AZURE_AI_MODEL) {
  console.warn(
    "[AVISO] Configure AZURE_AI_ENDPOINT, AZURE_AI_KEY e AZURE_AI_MODEL no arquivo .env antes de usar a IA."
  );
}

/**
 * Chama o endpoint de chat completions da Azure AI Foundry
 * (equivalente à função chamar_gemini do script Python original).
 */
async function chamarIA(promptTexto) {
  const url = `${AZURE_AI_ENDPOINT}/chat/completions`;

  const body = {
    model: AZURE_AI_MODEL,
    messages: [{ role: "user", content: promptTexto }],
  };

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": AZURE_AI_KEY,
        // Algumas configurações da Azure AI Foundry (Entra ID) esperam Bearer
        // em vez de api-key. Se necessário, troque a linha acima por:
        // "Authorization": `Bearer ${AZURE_AI_KEY}`,
      },
      body: JSON.stringify(body),
    });

    const data = await response.json();

    if (!response.ok) {
      return `Erro na API (${response.status}): ${JSON.stringify(data)}`;
    }

    return data.choices?.[0]?.message?.content ?? "";
  } catch (e) {
    return `Falha na requisição: ${e.message}`;
  }
}

// ==========================================
// LEITURA DE ARQUIVOS (.txt, .docx, .pdf)
// ==========================================
async function lerArquivoBuffer(buffer, nomeOriginal) {
  if (!buffer) return "";
  const ext = path.extname(nomeOriginal || "").toLowerCase();

  try {
    if (ext === ".txt") {
      return buffer.toString("utf-8");
    }

    if (ext === ".docx") {
      const resultado = await mammoth.extractRawText({ buffer });
      return resultado.value || "";
    }

    if (ext === ".pdf") {
      const resultado = await pdfParse(buffer);
      return resultado.text || "";
    }

    return "";
  } catch (e) {
    return `[Erro ao ler o arquivo: ${e.message}]`;
  }
}

// ==========================================
// CONFIGURAÇÕES ABNT DO DOCUMENTO WORD
// ==========================================
const FONTE = "Arial";
const TAMANHO_TEXTO = 24; // 12pt em "half-points"
const SECOES_RESUMO = new Set(["resumo", "abstract"]);
const SECOES_REFERENCIAS = new Set([
  "referencias",
  "referencia bibliografica",
  "referencias bibliograficas",
]);

function normalizar(txt) {
  return txt
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "") // remove acentos
    .toLowerCase()
    .trim();
}

const MARGENS = {
  top: convertMillimetersToTwip(30), // 3 cm
  left: convertMillimetersToTwip(30), // 3 cm
  bottom: convertMillimetersToTwip(20), // 2 cm
  right: convertMillimetersToTwip(20), // 2 cm
};

const INDENT_PRIMEIRA_LINHA = convertMillimetersToTwip(12.5); // 1.25 cm

function headerComNumeroPagina() {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [
          new TextRun({ font: FONTE, size: TAMANHO_TEXTO, children: [PageNumber.CURRENT] }),
        ],
      }),
    ],
  });
}

function pCentralizado(texto, { tamanho = TAMANHO_TEXTO, negrito = false, espacoAntes = 0 } = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    spacing: { line: 240, before: espacoAntes, after: 0 },
    children: [new TextRun({ text: texto, font: FONTE, size: tamanho, bold: negrito })],
  });
}

function pBlocoDireita(texto, { larguraCm = 7.5, tamanho = TAMANHO_TEXTO } = {}) {
  const recuoEsquerdoCm = 15 - larguraCm - 2;
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: convertMillimetersToTwip(recuoEsquerdoCm * 10), firstLine: 0 },
    spacing: { line: 240, after: 0 },
    children: [new TextRun({ text: texto, font: FONTE, size: tamanho })],
  });
}

function paragrafosVazios(n) {
  const arr = [];
  for (let i = 0; i < n; i++) {
    arr.push(new Paragraph({ spacing: { line: 240 }, children: [] }));
  }
  return arr;
}

function naturezaTexto(tipoDocumento, curso, instituicao, orientador) {
  let txt = "";
  if (tipoDocumento === "TCC") {
    txt =
      `Trabalho de Conclusão de Curso apresentado ao curso de ${curso || "[curso]"} ` +
      `da ${instituicao}, como requisito parcial para obtenção do título de graduado.`;
  } else if (tipoDocumento === "Trabalho Acadêmico") {
    txt =
      `Trabalho apresentado à disciplina de ${curso || "[disciplina]"}, ` +
      `do curso de graduação da ${instituicao}, como requisito parcial de avaliação.`;
  } else if (tipoDocumento === "Relatório de Estágio") {
    txt =
      `Relatório de Estágio Supervisionado apresentado ao curso de ${curso || "[curso]"} ` +
      `da ${instituicao}, como requisito parcial para aprovação na disciplina ` +
      `de Estágio Supervisionado.`;
  }
  if (orientador && orientador.trim()) {
    txt += `\n\nOrientador(a): ${orientador.trim()}`;
  }
  return txt;
}

function adicionarCapa(instituicao, nomeAluno, tituloTrabalho, cidade, ano) {
  const paragrafos = [];
  paragrafos.push(pCentralizado(instituicao.toUpperCase()));
  paragrafos.push(...paragrafosVazios(6));
  paragrafos.push(pCentralizado(nomeAluno.toUpperCase()));
  paragrafos.push(...paragrafosVazios(6));
  paragrafos.push(pCentralizado(tituloTrabalho.toUpperCase(), { negrito: true }));
  paragrafos.push(...paragrafosVazios(10));
  paragrafos.push(pCentralizado(cidade));
  paragrafos.push(pCentralizado(String(ano)));
  paragrafos.push(new Paragraph({ children: [], pageBreakBefore: true }));
  return paragrafos;
}

function adicionarFolhaRosto(nomeAluno, tituloTrabalho, instituicao, tipoDocumento, curso, orientador, cidade, ano) {
  const paragrafos = [];
  paragrafos.push(pCentralizado(nomeAluno.toUpperCase()));
  paragrafos.push(...paragrafosVazios(6));
  paragrafos.push(pCentralizado(tituloTrabalho.toUpperCase(), { negrito: true }));
  paragrafos.push(...paragrafosVazios(6));

  const natureza = naturezaTexto(tipoDocumento, curso, instituicao, orientador);
  for (const bloco of natureza.split("\n\n")) {
    paragrafos.push(pBlocoDireita(bloco));
    paragrafos.push(new Paragraph({ spacing: { line: 240 }, children: [] }));
  }

  paragrafos.push(...paragrafosVazios(8));
  paragrafos.push(pCentralizado(cidade));
  paragrafos.push(pCentralizado(String(ano)));
  paragrafos.push(new Paragraph({ children: [], pageBreakBefore: true }));
  return paragrafos;
}

function adicionarCabecalhoArtigo(tituloTrabalho, nomeAluno, instituicao) {
  const paragrafos = [];
  paragrafos.push(pCentralizado(tituloTrabalho.toUpperCase(), { negrito: true }));
  paragrafos.push(...paragrafosVazios(1));
  paragrafos.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { firstLine: 0 },
      spacing: { line: 240 },
      children: [new TextRun({ text: nomeAluno, font: FONTE, size: TAMANHO_TEXTO })],
    })
  );
  paragrafos.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { firstLine: 0 },
      spacing: { line: 240 },
      children: [new TextRun({ text: instituicao, font: FONTE, size: 20 })], // 10pt
    })
  );
  paragrafos.push(...paragrafosVazios(1));
  return paragrafos;
}

function montarElementosPretextuais(tipoDocumento, nomeAluno, instituicao, tituloTrabalho, curso, orientador, cidade, ano) {
  if (["TCC", "Trabalho Acadêmico", "Relatório de Estágio"].includes(tipoDocumento)) {
    return [
      ...adicionarCapa(instituicao, nomeAluno, tituloTrabalho, cidade, ano),
      ...adicionarFolhaRosto(nomeAluno, tituloTrabalho, instituicao, tipoDocumento, curso, orientador, cidade, ano),
    ];
  }
  if (tipoDocumento === "Artigo Científico") {
    return adicionarCabecalhoArtigo(tituloTrabalho, nomeAluno, instituicao);
  }
  return [];
}

/**
 * Converte o texto em markdown simplificado (#, ##, ###) gerado pela IA
 * em parágrafos formatados do corpo do documento.
 */
function adicionarCorpo(textoIA) {
  const paragrafos = [];
  const linhas = textoIA.trim().split("\n");
  let modo = "normal";

  for (let linhaBruta of linhas) {
    const linha = linhaBruta.trim();
    if (!linha) continue;

    if (linha.startsWith("### ")) {
      paragrafos.push(
        new Paragraph({
          indent: { firstLine: 0 },
          spacing: { before: 240, after: 120 }, // 12pt / 6pt
          children: [
            new TextRun({ text: linha.slice(4).trim(), font: FONTE, size: TAMANHO_TEXTO, bold: true, italics: true }),
          ],
        })
      );
      continue;
    }

    if (linha.startsWith("## ")) {
      paragrafos.push(
        new Paragraph({
          indent: { firstLine: 0 },
          spacing: { before: 240, after: 120 },
          children: [new TextRun({ text: linha.slice(3).trim(), font: FONTE, size: TAMANHO_TEXTO, bold: true })],
        })
      );
      continue;
    }

    if (linha.startsWith("# ")) {
      const titulo = linha.slice(2).trim();
      const norm = normalizar(titulo);
      modo = SECOES_RESUMO.has(norm) ? "resumo" : SECOES_REFERENCIAS.has(norm) ? "referencias" : "normal";
      paragrafos.push(
        new Paragraph({
          indent: { firstLine: 0 },
          spacing: { before: 360, after: 240 }, // 18pt / 12pt
          children: [new TextRun({ text: titulo.toUpperCase(), font: FONTE, size: TAMANHO_TEXTO, bold: true })],
        })
      );
      continue;
    }

    if (modo === "resumo") {
      paragrafos.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { firstLine: 0 },
          spacing: { line: 240 },
          children: [new TextRun({ text: linha, font: FONTE, size: TAMANHO_TEXTO })],
        })
      );
    } else if (modo === "referencias") {
      paragrafos.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          indent: { firstLine: 0 },
          spacing: { line: 240, after: 120 },
          children: [new TextRun({ text: linha, font: FONTE, size: TAMANHO_TEXTO })],
        })
      );
    } else {
      // modo normal: usa o estilo padrão do documento (justificado, recuo, 1.5 linhas)
      paragrafos.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { firstLine: INDENT_PRIMEIRA_LINHA },
          spacing: { line: 360 },
          children: [new TextRun({ text: linha, font: FONTE, size: TAMANHO_TEXTO })],
        })
      );
    }
  }

  return paragrafos;
}

function criarDocumento(paragrafos, { comCabecalhoPagina = false } = {}) {
  return new Document({
    sections: [
      {
        properties: { page: { margin: MARGENS } },
        headers: comCabecalhoPagina ? { default: headerComNumeroPagina() } : undefined,
        children: paragrafos,
      },
    ],
  });
}

async function gerarDocumentoWordBuffer(
  textoIA,
  tipoDocumento,
  nomeAluno,
  instituicao,
  tituloTrabalho,
  { curso = "", orientador = "", cidade = "Cidade", ano = "2026" } = {}
) {
  const pretextuais = montarElementosPretextuais(
    tipoDocumento,
    nomeAluno,
    instituicao,
    tituloTrabalho,
    curso,
    orientador,
    cidade,
    ano
  );
  const corpo = adicionarCorpo(textoIA);
  const doc = criarDocumento([...pretextuais, ...corpo], { comCabecalhoPagina: true });
  return Packer.toBuffer(doc);
}

async function gerarDocumentoSimplesBuffer(texto) {
  const corpo = adicionarCorpo(texto);
  const doc = criarDocumento(corpo, { comCabecalhoPagina: false });
  return Packer.toBuffer(doc);
}

// ==========================================
// CONTEXTOS DA ÁREA ACADÊMICA
// ==========================================
const CONTEXTOS = {
  TCC: (
    "Você é um especialista em TCC e normas ABNT. Gere um trabalho acadêmico completo, " +
    "com conteúdo original e linguagem acadêmica.\n" +
    "Estruture OBRIGATORIAMENTE nesta ordem, cada uma como seção de nível 1 (#): " +
    "Resumo (com 'Palavras-chave:' ao final do parágrafo), Introdução, Desenvolvimento " +
    "(pode ter subseções ##), Conclusão, Referências."
  ),
  "Trabalho Acadêmico": (
    "Você é especialista em trabalhos acadêmicos. Estruture o trabalho, melhore a escrita " +
    "e organize os capítulos.\n" +
    "Estruture nesta ordem, como seções de nível 1 (#): Introdução, Desenvolvimento " +
    "(pode ter subseções ##), Conclusão, Referências."
  ),
  "Artigo Científico": (
    "Você é especialista em artigos científicos, seguindo a NBR 6022.\n" +
    "Estruture OBRIGATORIAMENTE nesta ordem, como seções de nível 1 (#): " +
    "Resumo (parágrafo único, seguido de 'Palavras-chave:'), " +
    "Abstract (versão em inglês do resumo, seguido de 'Keywords:'), " +
    "Introdução, Desenvolvimento (com subseções ## como Metodologia, Resultados e Discussão), " +
    "Conclusão, Referências."
  ),
  "Relatório de Estágio": (
    "Você é especialista em relatórios de estágio.\n" +
    "Estruture nesta ordem, como seções de nível 1 (#): Apresentação da Empresa, " +
    "Atividades Desenvolvidas, Resultados Obtidos, Considerações Finais, Referências."
  ),
};

const REGRA_FORMATACAO = `
IMPORTANTE - formate sua resposta usando OBRIGATORIAMENTE esta marcação,
pois ela será convertida automaticamente em um documento Word:
Use "# " no início da linha para cada seção de nível 1 (ex: "# Introdução", "# Resumo", "# Referências").
Use "## " para subseções e "### " para subitens, se necessário.
Escreva os títulos de seção SEM numeração (a numeração não é aplicada automaticamente).
Não use "#" para nada além de títulos de seção.
Não use markdown de negrito (**) ou itálico (*) no texto corrido.
Na seção de Referências, coloque cada referência em uma linha própria (sem numerar).
Não inclua capa, folha de rosto, nome do aluno ou instituição no texto: isso é gerado à parte.
`;

// ==========================================
// LÓGICA DA ÁREA DE CORREÇÃO
// ==========================================
async function correcao(textoUsuario = "", textoArquivo = "") {
  let conteudo = "";
  if (textoUsuario) conteudo += textoUsuario + "\n\n";
  if (textoArquivo) conteudo += textoArquivo;

  if (!conteudo.trim()) {
    return { resposta: "Digite um texto ou envie um arquivo para correção.", bufferDocx: null, nomeArquivo: null };
  }

  const prompt = `
Você é um professor universitário especialista em:
Revisão textual
Gramática
Ortografia
ABNT
Correção acadêmica

Primeiro analise automaticamente o conteúdo enviado.
Identifique se é:
• Trabalho Acadêmico
• TCC
• Artigo Científico
• Relatório
• Redação
• Lista de Exercícios
• Questionário
• Prova
• Outro

Depois siga estas regras.

====================================================
SE FOR UM TRABALHO
Corrija ortografia, gramática, clareza, coesão, normas ABNT, referências e citações.
Preserve totalmente o sentido original.

Depois produza duas partes:
PARTE 1: TEXTO CORRIGIDO
PARTE 2: SUGESTÕES DE MELHORIA (em tópicos)

====================================================
SE FOR UMA LISTA DE EXERCÍCIOS
Corrija questão por questão.
Para cada questão informe:
Questão X
Resposta Correta
Explicação
====================================================

Conteúdo enviado:
${conteudo}
`;

  let resposta = await chamarIA(prompt);
  const respostaLower = resposta.toLowerCase();

  let bufferDocx;
  let nomeArquivo;
  let textoFinal;

  if (respostaLower.includes("questão 1") || respostaLower.includes("questao 1")) {
    bufferDocx = await gerarDocumentoSimplesBuffer(resposta);
    nomeArquivo = "questoes_corrigidas.docx";
    textoFinal = resposta;
  } else {
    let textoCorrigido = resposta;
    let sugestoes = "";
    if (resposta.includes("SUGESTÕES DE MELHORIA")) {
      const partes = resposta.split("SUGESTÕES DE MELHORIA");
      textoCorrigido = partes[0];
      sugestoes = "SUGESTÕES DE MELHORIA\n\n" + partes.slice(1).join("SUGESTÕES DE MELHORIA");
    }
    bufferDocx = await gerarDocumentoSimplesBuffer(textoCorrigido);
    nomeArquivo = "trabalho_corrigido.docx";
    textoFinal = sugestoes;
  }

  return { resposta: textoFinal, bufferDocx, nomeArquivo };
}

// ==========================================
// SERVIDOR EXPRESS
// ==========================================
const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } });

app.use(express.static(path.join(__dirname))); // serve index.html e estáticos na raiz
app.use(express.json());

// Guarda em memória os últimos documentos gerados por sessão simples (id -> buffer)
const documentosGerados = new Map();

function registrarDocumento(buffer, nomeArquivo) {
  const id = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  documentosGerados.set(id, { buffer, nomeArquivo });
  // limpeza simples depois de 30 min
  setTimeout(() => documentosGerados.delete(id), 30 * 60 * 1000);
  return id;
}

app.get("/download/:id", (req, res) => {
  const item = documentosGerados.get(req.params.id);
  if (!item) return res.status(404).send("Arquivo não encontrado ou expirado.");
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );
  res.setHeader("Content-Disposition", `attachment; filename="${item.nomeArquivo}"`);
  res.send(item.buffer);
});

// ---- Rota: geração de documentos acadêmicos (TCC, Artigo, Relatório, Trabalho) ----
app.post("/api/academico", upload.single("arquivo"), async (req, res) => {
  try {
    const { tipoDoc, prompt = "", nomeAluno = "", instituicao = "", curso = "", orientador = "", cidade = "" } = req.body;

    if (!CONTEXTOS[tipoDoc]) {
      return res.status(400).json({ erro: "Tipo de documento inválido." });
    }
    if (!prompt.trim() && !req.file) {
      return res.status(400).json({ erro: "Digite a solicitação ou envie um arquivo com as orientações do trabalho." });
    }
    if (!nomeAluno.trim() || !instituicao.trim()) {
      return res.status(400).json({ erro: "Preencha o Nome e a Instituição para montar a capa ABNT." });
    }

    const textoArquivo = req.file ? await lerArquivoBuffer(req.file.buffer, req.file.originalname) : "";
    const promptCompleto = prompt + (textoArquivo ? "\n\nMaterial anexado:\n" + textoArquivo : "");

    const contexto = CONTEXTOS[tipoDoc];
    const promptFinal = `
${contexto}

${REGRA_FORMATACAO}

Tipo de Documento: ${tipoDoc}

Solicitação do usuário:
${promptCompleto}
`;

    const resposta = await chamarIA(promptFinal);
    if (resposta.startsWith("Erro") || resposta.startsWith("Falha")) {
      return res.status(502).json({ erro: resposta });
    }

    const bufferDocx = await gerarDocumentoWordBuffer(resposta, tipoDoc, nomeAluno, instituicao, tipoDoc, {
      curso,
      orientador,
      cidade: cidade || "Cidade",
      ano: "2026",
    });

    const nomeArquivo = `${tipoDoc.toLowerCase().replace(/\s+/g, "_")}.docx`;
    const id = registrarDocumento(bufferDocx, nomeArquivo);

    res.json({ resposta, downloadUrl: `/download/${id}`, nomeArquivo });
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: `Falha na requisição:\n\n${e.message}` });
  }
});

// ---- Rota: correção automatizada ----
app.post("/api/correcao", upload.single("arquivo"), async (req, res) => {
  try {
    const { prompt = "" } = req.body;
    const textoArquivo = req.file ? await lerArquivoBuffer(req.file.buffer, req.file.originalname) : "";

    const { resposta, bufferDocx, nomeArquivo } = await correcao(prompt, textoArquivo);

    if (!bufferDocx) {
      return res.json({ resposta, downloadUrl: null, nomeArquivo: null });
    }

    const id = registrarDocumento(bufferDocx, nomeArquivo);
    res.json({ resposta, downloadUrl: `/download/${id}`, nomeArquivo });
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: `Falha na requisição:\n\n${e.message}` });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Assistente Acadêmico rodando em http://localhost:${PORT}`);
});
