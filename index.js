const express = require("express");
const { google } = require("googleapis");
const app = express();

app.get("/", async (req, res) => {
  try {
    const auth = new google.auth.GoogleAuth({
      keyFile: "credentials.json",
      scopes: "https://www.googleapis.com/auth/spreadsheets",
    });

    const client = await auth.getClient();

    const googleSheets = google.sheets({ version: "v4", auth: client });

    const spreadsheetId = "19UwCb9QO5hwOOqgFaT2vyf-RMfDviVBVnADBjCETnOg";
    const sheetName = "engenharia_de_software";

    const response = await googleSheets.spreadsheets.values.get({
      auth,
      spreadsheetId,
      range: `${sheetName}!A:H`,
    });

    const data = response.data.values;

    // Verificar se há dados na planilha
    if (!data || data.length < 4) {
      res.status(400).send("Não há dados suficientes na planilha.");
      return;
    }

    for (let i = 3; i < data.length; i++) {
      const matricula = data[i][0];
      const faltas = parseInt(data[i][2]);
      const p1 = parseFloat(data[i][3]);
      const p2 = parseFloat(data[i][4]);
      const p3 = parseFloat(data[i][5]);

      // Calcular a média
      const media = (p1 + p2 + p3) / 3;

      if (faltas > 0.25 * 60) {
        situacao = "Reprovado por Falta";
        await googleSheets.spreadsheets.values.update({
          auth,
          spreadsheetId,
          range: `${sheetName}!H${i + 1}`,
          valueInputOption: "USER_ENTERED",
          resource: {
            values: [[0]],
          },
        });
      } else {
        if (media < 5) {
          situacao = "Reprovado por Nota";
          await googleSheets.spreadsheets.values.update({
            auth,
            spreadsheetId,
            range: `${sheetName}!H${i + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: {
              values: [[0]],
            },
          });
        } else if (media >= 5 && media < 7) {
          situacao = "Exame Final";

          const naf = Math.round((7 - media) * 2);

          await googleSheets.spreadsheets.values.update({
            auth,
            spreadsheetId,
            range: `${sheetName}!H${i + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: {
              values: [[naf]],
            },
          });
        } else {
          situacao = "Aprovado";
          await googleSheets.spreadsheets.values.update({
            auth,
            spreadsheetId,
            range: `${sheetName}!H${i + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: {
              values: [[0]],
            },
          });
        }
      }

      await googleSheets.spreadsheets.values.update({
        auth,
        spreadsheetId,
        range: `${sheetName}!G${i + 1}`,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: [[situacao]],
        },
      });
    }

    res.send("Situações calculadas e escritas com sucesso!");
  } catch (error) {
    console.error("Erro:", error.message);
    res.status(500).send("Erro ao processar a solicitação");
  }
});

app.listen(1337, () => {
  console.log("Server running at port 1337");
});
