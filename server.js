const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

const NOME_PLANILHA = "personal fit pro dashboard 3.xlsx";
const DIRETORIO_DADOS = process.env.DATA_DIR || __dirname;
const PLANILHA = path.join(DIRETORIO_DADOS, NOME_PLANILHA);
const REPOSITORIO_GITHUB = process.env.GITHUB_REPOSITORY || "";
const BRANCH_GITHUB = process.env.GITHUB_BRANCH || "main";
const TOKEN_GITHUB = process.env.GITHUB_TOKEN || "";

const sessoes = new Map();
const DURACAO_SESSAO = 1000 * 60 * 60 * 8;

app.use(express.json());
app.set("trust proxy", 1);

if (!fs.existsSync(PLANILHA)) {
  fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
  fs.copyFileSync(path.join(__dirname, NOME_PLANILHA), PLANILHA);
}

function normalizarNome(nome) {
  return String(nome || "").trim().toLocaleLowerCase("pt-BR");
}

function criarSessao(aluno) {
  const token = crypto.randomBytes(32).toString("hex");
  sessoes.set(token, { aluno, expiraEm: Date.now() + DURACAO_SESSAO });
  return token;
}

function lerCookie(req, nome) {
  const cookies = String(req.headers.cookie || "").split(";");
  const cookie = cookies.find(item => item.trim().startsWith(`${nome}=`));
  return cookie ? decodeURIComponent(cookie.trim().slice(nome.length + 1)) : "";
}

function exigirSessao(req, res, next) {
  const token = lerCookie(req, "treinos_session");
  const sessao = sessoes.get(token);

  if (!sessao || sessao.expiraEm < Date.now()) {
    sessoes.delete(token);
    return res.status(401).json({ erro: "Sessão expirada. Faça login novamente." });
  }

  req.aluno = sessao.aluno;
  next();
}


// =====================================================
// LER A PLANILHA
// =====================================================

function carregarPlanilha() {
  const workbook = XLSX.readFile(PLANILHA);
  return workbook;
}

function githubConfigurado() {
  return Boolean(TOKEN_GITHUB && REPOSITORIO_GITHUB.includes("/"));
}

function urlArquivoGithub() {
  return `https://api.github.com/repos/${REPOSITORIO_GITHUB}/contents/${encodeURIComponent(NOME_PLANILHA)}`;
}

async function obterPlanilhaGithub() {
  const resposta = await fetch(`${urlArquivoGithub()}?ref=${encodeURIComponent(BRANCH_GITHUB)}`, {
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${TOKEN_GITHUB}`,
      "X-GitHub-Api-Version": "2022-11-28"
    }
  });

  if (!resposta.ok) throw new Error(`GitHub não conseguiu ler a planilha (${resposta.status}).`);
  const arquivo = await resposta.json();
  return { workbook: XLSX.read(Buffer.from(arquivo.content, "base64")), sha: arquivo.sha };
}

async function salvarPlanilhaGithub(workbook, sha) {
  const conteudo = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  const resposta = await fetch(urlArquivoGithub(), {
    method: "PUT",
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${TOKEN_GITHUB}`,
      "Content-Type": "application/json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    body: JSON.stringify({
      message: "Atualiza progresso do aluno",
      content: conteudo.toString("base64"),
      sha,
      branch: BRANCH_GITHUB
    })
  });

  if (!resposta.ok) throw new Error(`GitHub não conseguiu salvar a planilha (${resposta.status}).`);
  fs.writeFileSync(PLANILHA, conteudo);
}


// =====================================================
// LER ALUNOS
// =====================================================

function carregarAlunos() {
  const workbook = carregarPlanilha();

  const aba = workbook.Sheets["ALUNOS"];

  if (!aba) {
    throw new Error("A aba ALUNOS não foi encontrada.");
  }

  const dados = XLSX.utils.sheet_to_json(aba, {
    header: 1,
    defval: ""
  });

  const alunos = [];

  for (let i = 1; i < dados.length; i++) {

    const nome = String(dados[i][0] || "").trim();
    const senha = String(dados[i][1] || "").trim();

    if (!nome) continue;

    alunos.push({
      nome,
      senha
    });
  }

  return alunos;
}


// =====================================================
// LER TREINO DO ALUNO
// =====================================================

function carregarTreino(nomeAluno) {

  const workbook = carregarPlanilha();

  const nomeAba = workbook.SheetNames.find(
    nome =>
      nome.toLowerCase().trim() ===
      nomeAluno.toLowerCase().trim()
  );

  if (!nomeAba) {
    throw new Error(
      `Aba do aluno "${nomeAluno}" não encontrada.`
    );
  }

  const aba = workbook.Sheets[nomeAba];

  const dados = XLSX.utils.sheet_to_json(aba, {
    header: 1,
    defval: ""
  });

  const treinos = {};

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    const divisao = String(linha[0] || "")
      .trim()
      .toUpperCase();

    const exercicio = String(linha[1] || "")
      .trim();

    const series = linha[2] || "";
    const repeticoes = linha[3] || "";
    const video = String(linha[4] || "").trim();

    if (!divisao || !exercicio) continue;

    if (!treinos[divisao]) {
      treinos[divisao] = [];
    }

    treinos[divisao].push({

      exercicio,

      series,

      repeticoes,

      video

    });
  }

  return treinos;
}

function carregarProgresso(nomeAluno) {
  const workbook = carregarPlanilha();
  const aba = workbook.Sheets.PROGRESSO;
  if (!aba) return {};

  const dados = XLSX.utils.sheet_to_json(aba, { header: 1, defval: "" });
  const progresso = {};

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (normalizarNome(linha[0]) !== normalizarNome(nomeAluno)) continue;
    const chave = `${nomeAluno}_${linha[1]}_${linha[2]}`;
    progresso[chave] = String(linha[3]).toLowerCase() === "true";
  }

  return progresso;
}

async function salvarProgresso(nomeAluno, treino, indice, concluido) {
  let workbook;
  let sha;

  if (githubConfigurado()) {
    const remoto = await obterPlanilhaGithub();
    workbook = remoto.workbook;
    sha = remoto.sha;
  } else {
    workbook = carregarPlanilha();
  }

  const nomeAba = "PROGRESSO";
  const aba = workbook.Sheets[nomeAba];
  const dados = aba
    ? XLSX.utils.sheet_to_json(aba, { header: 1, defval: "" })
    : [["Aluno", "Treino", "Exercício", "Concluído", "Atualizado em"]];
  const nomeNormalizado = normalizarNome(nomeAluno);
  const linhaExistente = dados.findIndex((linha, indiceLinha) => (
    indiceLinha > 0 &&
    normalizarNome(linha[0]) === nomeNormalizado &&
    String(linha[1]) === String(treino) &&
    String(linha[2]) === String(indice)
  ));
  const novaLinha = [nomeAluno, treino, indice, Boolean(concluido), new Date().toISOString()];

  if (linhaExistente >= 0) dados[linhaExistente] = novaLinha;
  else dados.push(novaLinha);

  workbook.Sheets[nomeAba] = XLSX.utils.aoa_to_sheet(dados);
  if (!workbook.SheetNames.includes(nomeAba)) workbook.SheetNames.push(nomeAba);

  if (githubConfigurado()) {
    await salvarPlanilhaGithub(workbook, sha);
  } else {
    XLSX.writeFile(workbook, PLANILHA);
  }
}

async function sincronizarPlanilhaRemota(req, res, next) {
  if (!githubConfigurado()) return next();

  try {
    const remoto = await obterPlanilhaGithub();
    const conteudo = XLSX.write(remoto.workbook, { type: "buffer", bookType: "xlsx" });
    fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
    fs.writeFileSync(PLANILHA, conteudo);
    next();
  } catch (erro) {
    console.error(erro);
    res.status(503).json({ erro: "A base de dados está temporariamente indisponível." });
  }
}


// =====================================================
// API — LISTAR ALUNOS
// =====================================================

app.use("/api", sincronizarPlanilhaRemota);

app.get("/api/alunos", (req, res) => {

  try {

    const alunos = carregarAlunos();

    res.json(
      alunos.map(aluno => ({
        nome: aluno.nome
      }))
    );

  } catch (erro) {

    console.error(erro);

    res.status(500).json({
      erro: "Erro ao carregar alunos."
    });

  }

});


// =====================================================
// API — LOGIN
// =====================================================

app.post("/api/login", (req, res) => {

  try {

    const nome = String(req.body.nome || "").trim();
    const senha = String(req.body.senha || "").trim();

    const alunos = carregarAlunos();

    const aluno = alunos.find(
      item =>
        item.nome.toLowerCase() === nome.toLowerCase() &&
        item.senha === senha
    );

    if (!aluno) {

      return res.status(401).json({

        sucesso: false,

        erro: "Nome ou senha incorretos."

      });

    }

    const treinos = carregarTreino(aluno.nome);
    const progresso = carregarProgresso(aluno.nome);
    const token = criarSessao(aluno.nome);

    res.setHeader(
      "Set-Cookie",
      `treinos_session=${encodeURIComponent(token)}; HttpOnly; SameSite=Lax;${req.secure ? " Secure;" : ""} Path=/; Max-Age=${DURACAO_SESSAO / 1000}`
    );

    res.json({

      sucesso: true,

      aluno: {

        nome: aluno.nome

      },

      treinos,
      progresso

    });

  } catch (erro) {

    console.error(erro);

    res.status(500).json({

      sucesso: false,

      erro: "Erro ao realizar login."

    });

  }

});

app.post("/api/logout", (req, res) => {
  const token = lerCookie(req, "treinos_session");
  sessoes.delete(token);
  res.setHeader("Set-Cookie", "treinos_session=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0");
  res.json({ sucesso: true });
});

app.get("/api/sessao", exigirSessao, (req, res) => {
  res.json({
    sucesso: true,
    aluno: { nome: req.aluno },
    treinos: carregarTreino(req.aluno),
    progresso: carregarProgresso(req.aluno)
  });
});

app.post("/api/progresso", exigirSessao, async (req, res) => {
  try {
    const treino = String(req.body.treino || "").trim().toUpperCase();
    const indice = Number(req.body.indice);
    const concluido = Boolean(req.body.concluido);
    const treinos = carregarTreino(req.aluno);

    if (!treinos[treino] || !Number.isInteger(indice) || !treinos[treino][indice]) {
      return res.status(400).json({ erro: "Exercício inválido." });
    }

    await salvarProgresso(req.aluno, treino, indice, concluido);
    res.json({ sucesso: true });
  } catch (erro) {
    console.error(erro);
    res.status(500).json({ erro: "Não foi possível salvar o progresso." });
  }
});


// =====================================================
// SERVIR O INDEX.HTML
// =====================================================

app.get("/", (req, res) => {

  res.sendFile(
    path.join(__dirname, "index.html")
  );

});

app.get("/health", (req, res) => {
  res.status(200).json({ status: "ok" });
});

app.use("/api", (req, res) => {
  res.status(404).json({ erro: "Endpoint da API não encontrado." });
});

app.use((erro, req, res, next) => {
  console.error(erro);
  if (req.path.startsWith("/api")) {
    return res.status(500).json({ erro: "Erro interno do servidor." });
  }
  next(erro);
});


// =====================================================
// INICIAR SERVIDOR
// =====================================================

app.listen(PORT, "0.0.0.0", () => {

  console.log(
    `Treinos Personal rodando na porta ${PORT}`
  );

});