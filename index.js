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

    if (!data) return res.status(404).send("Dashboard not found");

    const html = `
    <html>
    <head>
        <title>Pro Dashboard</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

        <style>
            body {
                font-family: Arial;
                padding: 20px;
                background: #f5f6fa;
            }

            h1 {
                text-align: center;
            }

            .filters {
                display: flex;
                gap: 10px;
                justify-content: center;
                margin-bottom: 20px;
            }

            .grid {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 20px;
            }

            .card {
                background: white;
                padding: 15px;
                border-radius: 10px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            }

            canvas {
                width: 100% !important;
                height: 300px !important;
            }
        </style>
    </head>

    <body>

        <h1>📊 Operations Dashboard</h1>

        <!-- FILTERS -->
        <div class="filters">
            <select id="metric">
                <option value="agent">Agent</option>
                <option value="company">Company</option>
                <option value="issue">Issue</option>
            </select>
        </div>

        <div class="grid">

            <!-- 1. DONUT CHART -->
            <div class="card">
                <h3>Tickets by Issue Type</h3>
                <canvas id="issueChart"></canvas>
            </div>

            <!-- 2. BAR CHART -->
            <div class="card">
                <h3>Agent Performance</h3>
                <canvas id="agentChart"></canvas>
            </div>

            <!-- 3. COMPANY -->
            <div class="card">
                <h3>Company Tickets</h3>
                <canvas id="companyChart"></canvas>
            </div>

            <!-- 4. DAILY TREND -->
            <div class="card">
                <h3>Daily Trends</h3>
                <canvas id="trendChart"></canvas>
            </div>

            <!-- 5. BACKLOG -->
            <div class="card">
                <h3>Backlog Aging</h3>
                <canvas id="backlogChart"></canvas>
            </div>

            <!-- 6. RISK -->
            <div class="card">
                <h3>Risk Analysis</h3>
                <canvas id="riskChart"></canvas>
            </div>

        </div>

        <script>
            const data = ${JSON.stringify(data)};

            // ISSUE DONUT
            new Chart(document.getElementById("issueChart"), {
                type: "doughnut",
                data: {
                    labels: Object.keys(data.issue_type_trends),
                    datasets: [{
                        data: Object.values(data.issue_type_trends)
                    }]
                }
            });

            // AGENT BAR
            new Chart(document.getElementById("agentChart"), {
                type: "bar",
                data: {
                    labels: Object.keys(data.agent_performance),
                    datasets: [{
                        label: "Resolved",
                        data: Object.values(data.agent_performance).map(a => a.resolved)
                    }]
                }
            });

            // COMPANY BAR
            new Chart(document.getElementById("companyChart"), {
                type: "bar",
                data: {
                    labels: Object.keys(data.company_stats),
                    datasets: [{
                        label: "Tickets",
                        data: Object.values(data.company_stats).map(c => c.total)
                    }]
                }
            });

            // TREND LINE
            const dates = Object.keys(data.daily_trends);
            const created = dates.map(d => data.daily_trends[d].created);
            const resolved = dates.map(d => data.daily_trends[d].resolved);

            new Chart(document.getElementById("trendChart"), {
                type: "line",
                data: {
                    labels: dates,
                    datasets: [
                        { label: "Created", data: created },
                        { label: "Resolved", data: resolved }
                    ]
                }
            });

            // BACKLOG
            new Chart(document.getElementById("backlogChart"), {
                type: "bar",
                data: {
                    labels: ["2 Days", "5 Days", "10 Days"],
                    datasets: [{
                        data: [
                            data.backlog_analysis.older_than_2_days,
                            data.backlog_analysis.older_than_5_days,
                            data.backlog_analysis.older_than_10_days
                        ]
                    }]
                }
            });

            // RISK
            new Chart(document.getElementById("riskChart"), {
                type: "pie",
                data: {
                    labels: ["High Risk", "Stuck", "Waiting"],
                    datasets: [{
                        data: [
                            data.risk_analysis.high_risk_tickets,
                            data.risk_analysis.stuck_pending,
                            data.risk_analysis.waiting_customer_long
                        ]
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
    console.log("Server running on port " + PORT);
});
