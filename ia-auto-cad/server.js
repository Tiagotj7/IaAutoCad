const express = require("express");
const multer = require("multer");
const path = require("path");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
const port = 3000;

// Middleware para permitir CORS
const cors = require("cors");
app.use(cors());

// Configuração do upload de arquivos (DXF e Excel)
const storage = multer.diskStorage({
  destination: "uploads/",
  filename: (req, file, cb) => {
    cb(null, file.fieldname + "-" + Date.now() + path.extname(file.originalname));
  },
});

const upload = multer({ storage });

// Função para processar o DXF (simulação)
function processarDXF(filePath) {
  // Simula extração de dados do DXF
  return [
    { "Largura (cm)": 50, "Altura (cm)": 30, "Comprimento (m)": 2, "Área (m²)": 3, "Peso (kg)": 20, "Preço (R$)": 300 },
    { "Largura (cm)": 100, "Altura (cm)": 50, "Comprimento (m)": 1, "Área (m²)": 5, "Peso (kg)": 40, "Preço (R$)": 600 }
  ];
}

// Rota para processar arquivos
app.post("/api/processar", upload.fields([{ name: "arquivo_dxf" }, { name: "arquivo_excel" }]), async (req, res) => {
  try {
    if (!req.files["arquivo_dxf"] || !req.files["arquivo_excel"]) {
      return res.status(400).json({ erro: "Arquivos não encontrados" });
    }

    const arquivoDXF = req.files["arquivo_dxf"][0].path;
    const arquivoExcel = req.files["arquivo_excel"][0].path;

    // Processa o DXF e retorna os dados
    const materiais = processarDXF(arquivoDXF);

    // Atualiza o Excel
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(arquivoExcel);
    const sheet = workbook.getWorksheet(1) || workbook.addWorksheet("Orçamento");

    // Limpa e escreve os dados
    sheet.getRow(1).values = ["Largura (cm)", "Altura (cm)", "Comprimento (m)", "Área (m²)", "Peso (kg)", "Preço (R$)"];
    materiais.forEach((material, index) => {
      sheet.getRow(index + 2).values = Object.values(material);
    });

    await workbook.xlsx.writeFile(arquivoExcel);

    res.json({ mensagem: "Planilha atualizada com sucesso!", materiais });

    // Remove arquivos temporários
    fs.unlinkSync(arquivoDXF);
    fs.unlinkSync(arquivoExcel);
  } catch (error) {
    console.error("Erro:", error);
    res.status(500).json({ erro: "Erro ao processar arquivos." });
  }
});

// Inicia o servidor
app.listen(port, () => {
  console.log(`Servidor rodando em http://localhost:${port}`);
});
