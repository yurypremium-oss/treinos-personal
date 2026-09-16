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

function salvarProgresso(nomeAluno, treino, indice, concluido) {
  const workbook = carregarPlanilha();
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
  XLSX.writeFile(workbook, PLANILHA);
}


// =====================================================
// API — LISTAR ALUNOS
// =====================================================

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

app.post("/api/progresso", exigirSessao, (req, res) => {
  try {
    const treino = String(req.body.treino || "").trim().toUpperCase();
    const indice = Number(req.body.indice);
    const concluido = Boolean(req.body.concluido);
    const treinos = carregarTreino(req.aluno);

    if (!treinos[treino] || !Number.isInteger(indice) || !treinos[treino][indice]) {
      return res.status(400).json({ erro: "Exercício inválido." });
    }

    salvarProgresso(req.aluno, treino, indice, concluido);
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


// =====================================================
// INICIAR SERVIDOR
// =====================================================

app.listen(PORT, () => {

  console.log(
    `Treinos Personal rodando na porta ${PORT}`
  );

});