const express = require("express");
const ExcelJS = require("exceljs");
const { v4: uuidv4 } = require("uuid");

const app = express();
app.use(express.json());

// In-memory storage (for demo; use DB/S3 in real apps)
const reports = {};
const dashboards = {};

// Your deployed URL (fallback if needed)
const BASE_URL = "https://ticket-report-generation.onrender.com";

/* =========================================================
   📄 GENERATE REPORT → RETURN DOWNLOAD LINK
========================================================= */
app.post("/generate-report", async (req, res) => {
    try {
        const data = req.body;
        const id = uuidv4();

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Report");

        sheet.addRow(["Ticket Report"]);
        sheet.addRow([]);
        sheet.addRow(["Total Tickets", data.total_tickets_last_2_months]);

        sheet.addRow([]);
        sheet.addRow(["Agent", "Resolved"]);

        for (const agent in data.agent_resolution_performance) {
            sheet.addRow([
                agent,
                data.agent_resolution_performance[agent].resolved
            ]);
        }

        sheet.addRow([]);
        sheet.addRow(["Issue Type", "Count"]);

        for (const issue in data.issue_type_trends) {
            sheet.addRow([issue, data.issue_type_trends[issue]]);
        }

        const buffer = await workbook.xlsx.writeBuffer();
        reports[id] = buffer;

        // Dynamic base URL (works everywhere)
        const baseUrl = req.protocol + "://" + req.get("host") || BASE_URL;

        const downloadUrl = `${baseUrl}/download-report/${id}`;

        res.json({
            message: "Report generated successfully",
            download_url: downloadUrl
        });

    } catch (err) {
        console.error(err);
        res.status(500).send("Error generating report");
    }
});

/* =========================================================
   📥 DOWNLOAD REPORT
========================================================= */
app.get("/download-report/:id", (req, res) => {
    const file = reports[req.params.id];

    if (!file) {
        return res.status(404).send("Report not found");
    }

    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
        "Content-Disposition",
        "attachment; filename=report.xlsx"
    );

    res.send(file);
});

/* =========================================================
   📊 GENERATE DASHBOARD → RETURN LIVE URL
========================================================= */
app.post("/dashboard", (req, res) => {
    const data = req.body;
    const id = uuidv4();

    dashboards[id] = data;

    const baseUrl = req.protocol + "://" + req.get("host") || BASE_URL;

    const dashboardUrl = `${baseUrl}/dashboard/${id}`;

    res.json({
        message: "Dashboard created",
        dashboard_url: dashboardUrl
    });
});

/* =========================================================
   🌐 DASHBOARD PAGE
========================================================= */
app.get("/dashboard/:id", (req, res) => {
    const data = dashboards[req.params.id];

    if (!data) {
        return res.status(404).send("Dashboard not found");
    }

    const agentLabels = Object.keys(data.agent_resolution_performance);
    const agentData = Object.values(data.agent_resolution_performance).map(a => a.resolved);

    const html = `
    <html>
    <head>
        <title>Dashboard</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            body { font-family: Arial; text-align: center; }
            canvas { max-width: 700px; margin: auto; }
        </style>
    </head>
    <body>
        <h1>📊 Ticket Dashboard</h1>
        <h2>Total Tickets: ${data.total_tickets_last_2_months}</h2>

        <canvas id="chart"></canvas>

        <script>
            new Chart(document.getElementById("chart"), {
                type: "bar",
                data: {
                    labels: ${JSON.stringify(agentLabels)},
                    datasets: [{
                        label: "Resolved Tickets",
                        data: ${JSON.stringify(agentData)}
                    }]
                }
            });
        </script>
    </body>
    </html>
    `;

    res.send(html);
});

/* =========================================================
   🚀 START SERVER
========================================================= */
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(\`Server running on port \${PORT}\`);
});
