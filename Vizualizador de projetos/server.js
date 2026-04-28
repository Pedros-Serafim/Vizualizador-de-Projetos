const express = require("express");
const axios = require("axios");
const path = require("path");

const app = express();
const PORT = 8080;

const ONEDRIVE_URL = "https://lifeservicos-my.sharepoint.com/personal/pedro_serafim_life_net_br/Documents/Vizualizador%20de%20projetos/Pr%C3%A9dios%20e%20Condom%C3%ADnios.xlsx?download=1";

app.use(express.static(__dirname));

app.get("/api/planilha", async (req, res) => {
    try {
        const response = await axios.get(ONEDRIVE_URL, {
            responseType: "arraybuffer"
        });

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );

        res.send(response.data);
    } catch (error) {
        console.error("Erro ao carregar planilha:", error.message);
        res.status(500).send("Erro ao carregar planilha");
    }
});

app.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
});