const express = require("express");
const ExcelJS = require("exceljs");
const { v4: uuidv4 } = require("uuid");

const app = express();
app.use(express.json());

const reports = {};
const dashboards = {};

const BASE_URL = "https://ticket-report-generation.onrender.com";

/* =========================================================
   📄 GENERATE EXCEL REPORT (MULTI SHEET)
========================================================= */
app.post("/generate-report", async (req, res) => {
    try {
        const data = req.body;
        const id = uuidv4();

        const workbook = new ExcelJS.Workbook();

        /* ================= SUMMARY ================= */
        const summary = workbook.addWorksheet("Summary");
        summary.addRow(["Metric", "Value"]);
        summary.addRow(["Total Tickets", data.total_tickets_last_2_months]);
        summary.addRow(["Avg Resolution Time (hrs)", data.resolution_metrics.avg_resolution_time_hours]);
        summary.addRow(["Avg First Response (mins)", data.first_response_time.avg_minutes]);

        summary.addRow([]);
        //summary.addRow(["AI Insights"]);
        //data.ai_insights.forEach(i => summary.addRow([i]));

        /* ================= AGENTS ================= */
        const agents = workbook.addWorksheet("Agents");
        agents.addRow(["Agent", "Resolved", "Avg Resolution Time", "Efficiency"]);

        for (const agent in data.agent_performance) {
            agents.addRow([
                agent,
                data.agent_performance[agent].resolved,
                data.agent_performance[agent].avg_resolution_time,
                data.agent_efficiency[agent]
            ]);
        }

        /* ================= COMPANIES ================= */
        const companies = workbook.addWorksheet("Companies");
        companies.addRow(["Company", "Total", "Resolved", "Open", "Resolution %", "Avg Time"]);

        for (const comp in data.company_stats) {
            const c = data.company_stats[comp];
            companies.addRow([
                comp,
                c.total,
                c.resolved,
                c.open,
                c.resolution_rate,
                c.avg_resolution_time
            ]);
        }

        /* ================= ISSUES ================= */
        const issues = workbook.addWorksheet("Issues");
        issues.addRow(["Issue Type", "Count"]);

        for (const issue in data.issue_type_trends) {
            issues.addRow([issue, data.issue_type_trends[issue]]);
        }

        /* ================= BACKLOG ================= */
        const backlog = workbook.addWorksheet("Backlog & Risk");
        backlog.addRow(["Metric", "Value"]);
        backlog.addRow(["Open Tickets", data.backlog_analysis.total_open]);
        backlog.addRow([">2 Days", data.backlog_analysis.older_than_2_days]);
        backlog.addRow([">5 Days", data.backlog_analysis.older_than_5_days]);
        backlog.addRow([">10 Days", data.backlog_analysis.older_than_10_days]);

        backlog.addRow([]);
        backlog.addRow(["High Risk Tickets", data.risk_analysis.high_risk_tickets]);
        backlog.addRow(["Stuck Pending", data.risk_analysis.stuck_pending]);

        /* ================= TRENDS ================= */
        const trends = workbook.addWorksheet("Trends");
        trends.addRow(["Date", "Created", "Resolved"]);

        for (const day in data.daily_trends) {
            const d = data.daily_trends[day];
            trends.addRow([day, d.created, d.resolved]);
        }

        const buffer = await workbook.xlsx.writeBuffer();
        reports[id] = buffer;

        const baseUrl = req.protocol + "://" + req.get("host") || BASE_URL;

        res.json({
            message: "Report generated successfully",
            download_url: `${baseUrl}/download-report/${id}`
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
   📊 GENERATE DASHBOARD LINK
========================================================= */
app.post("/dashboard", (req, res) => {
    const data = req.body;
    const id = uuidv4();

    dashboards[id] = data;

    const baseUrl = req.protocol + "://" + req.get("host") || BASE_URL;

    res.json({
        message: "Dashboard created",
        dashboard_url: `${baseUrl}/dashboard/${id}`
    });
});

/* =========================================================
   🌐 ADVANCED DASHBOARD PAGE
========================================================= */
app.get("/dashboard/:id", (req, res) => {
    const data = dashboards[req.params.id];

    if (!data) {
        return res.status(404).send("Dashboard not found");
    }

    const html = `
    <html>
    <head>
        <title>Advanced Dashboard</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            body { font-family: Arial; padding: 20px; text-align: center; }
            .cards { display: flex; justify-content: space-around; margin-bottom: 20px; }
            .card { padding: 20px; background: #f4f4f4; border-radius: 10px; width: 200px; }
            canvas { max-width: 600px; margin: 20px auto; }
        </style>
    </head>
    <body>

        <h1>📊 Ticket Dashboard</h1>

        <div class="cards">
            <div class="card">Total Tickets<br><b>${data.total_tickets_last_2_months}</b></div>
            <div class="card">Avg Resolution<br><b>${data.resolution_metrics.avg_resolution_time_hours} hrs</b></div>
            <div class="card">High Risk<br><b>${data.risk_analysis.high_risk_tickets}</b></div>
        </div>

        <h3>Select Metric</h3>
        <select id="metricSelector">
            <option value="agent">Agent Performance</option>
            <option value="company">Company Tickets</option>
            <option value="issue">Issue Types</option>
        </select>

        <canvas id="chart"></canvas>

        <script>
            const data = ${JSON.stringify(data)};
            let chart;

            function renderChart(type) {

                let labels = [];
                let values = [];

                if (type === "agent") {
                    labels = Object.keys(data.agent_performance);
                    values = labels.map(a => data.agent_performance[a].resolved);
                }

                if (type === "company") {
                    labels = Object.keys(data.company_stats);
                    values = labels.map(c => data.company_stats[c].total);
                }

                if (type === "issue") {
                    labels = Object.keys(data.issue_type_trends);
                    values = Object.values(data.issue_type_trends);
                }

                if (chart) chart.destroy();

                chart = new Chart(document.getElementById("chart"), {
                    type: "bar",
                    data: {
                        labels: labels,
                        datasets: [{
                            label: type,
                            data: values
                        }]
                    }
                });
            }

            document.getElementById("metricSelector")
                .addEventListener("change", (e) => {
                    renderChart(e.target.value);
                });

            renderChart("agent");
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
    console.log("Server running on port " + PORT);
});
