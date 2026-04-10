const express = require("express");
const ExcelJS = require("exceljs");
const { v4: uuidv4 } = require("uuid");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(express.json());

const BASE_URL = "https://ticket-report-generation.onrender.com";

// 📁 Folder setup
const REPORT_DIR = path.join(__dirname, "reports");
if (!fs.existsSync(REPORT_DIR)) {
  fs.mkdirSync(REPORT_DIR);
}

/* =========================================================
   📄 GENERATE EXCEL REPORT
========================================================= */
app.post("/generate-report", async (req, res) => {
  try {
    const data = req.body;
    const id = uuidv4();

    const workbook = new ExcelJS.Workbook();

    const summary = workbook.addWorksheet("Summary");
    summary.addRow(["Metric", "Value"]);
    summary.addRow(["Total Tickets", data.total_tickets_last_2_months]);
    summary.addRow(["Avg Resolution Time (hrs)", data.resolution_metrics.avg_resolution_time_hours]);
    summary.addRow(["Avg First Response (mins)", data.first_response_time.avg_minutes]);

    const agents = workbook.addWorksheet("Agents");
    agents.addRow(["Agent", "Resolved", "Avg Resolution Time", "Efficiency"]);
    for (const agent in data.agent_performance) {
      agents.addRow([
        agent,
        data.agent_performance[agent].resolved,
        data.agent_performance[agent].avg_resolution_time,
        data.agent_efficiency[agent],
      ]);
    }

    const companies = workbook.addWorksheet("Companies");
    companies.addRow(["Company", "Total", "Resolved", "Open", "Resolution %", "Avg Time"]);
    for (const comp in data.company_stats) {
      const c = data.company_stats[comp];
      companies.addRow([comp, c.total, c.resolved, c.open, c.resolution_rate, c.avg_resolution_time]);
    }

    const issues = workbook.addWorksheet("Issues");
    issues.addRow(["Issue Type", "Count"]);
    for (const issue in data.issue_type_trends) {
      issues.addRow([issue, data.issue_type_trends[issue]]);
    }

    const filePath = path.join(REPORT_DIR, `${id}.xlsx`);
    await workbook.xlsx.writeFile(filePath);

    res.json({
      download_url: `${BASE_URL}/download-report/${id}`,
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
  const filePath = path.join(REPORT_DIR, `${req.params.id}.xlsx`);
  if (!fs.existsSync(filePath)) {
    return res.status(404).send("Report not found");
  }
  res.download(filePath);
});

/* =========================================================
   📊 DASHBOARD GENERATION
========================================================= */
app.post("/dashboard", (req, res) => {
  const encoded = encodeURIComponent(JSON.stringify(req.body));
  res.json({
    dashboard_url: `${BASE_URL}/dashboard?data=${encoded}`,
  });
});

/* =========================================================
   🌐 DASHBOARD UI (FULL CHARTS BACK)
========================================================= */
app.get("/dashboard", (req, res) => {
  if (!req.query.data) return res.send("No data");

  const data = JSON.parse(decodeURIComponent(req.query.data));

  const html = `
  <html>
  <head>
    <title>Advanced Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

    <style>
      body { font-family: Arial; background:#f5f6fa; padding:20px; }
      h1 { text-align:center; }
      .cards { display:flex; justify-content:space-around; margin-bottom:20px; }
      .card { background:white; padding:15px; border-radius:10px; box-shadow:0 2px 5px rgba(0,0,0,0.1); }
      .grid { display:grid; grid-template-columns:repeat(3,1fr); gap:20px; }
      canvas { width:100% !important; height:300px !important; }
    </style>
  </head>

  <body>
    <h1>📊 Operations Dashboard</h1>

    <div class="cards">
      <div class="card">Total Tickets<br><b>${data.total_tickets_last_2_months}</b></div>
      <div class="card">Avg Resolution<br><b>${data.resolution_metrics.avg_resolution_time_hours} hrs</b></div>
      <div class="card">High Risk<br><b>${data.risk_analysis.high_risk_tickets}</b></div>
    </div>

    <div class="grid">
      <div class="card"><h3>Issue Types</h3><canvas id="issueChart"></canvas></div>
      <div class="card"><h3>Agent Performance</h3><canvas id="agentChart"></canvas></div>
      <div class="card"><h3>Company Tickets</h3><canvas id="companyChart"></canvas></div>

      <div class="card"><h3>Daily Trends</h3><canvas id="trendChart"></canvas></div>
      <div class="card"><h3>Backlog</h3><canvas id="backlogChart"></canvas></div>
      <div class="card"><h3>Risk</h3><canvas id="riskChart"></canvas></div>

      <div class="card"><h3>Company vs Issue</h3><canvas id="stackedChart"></canvas></div>
    </div>

    <script>
      const data = ${JSON.stringify(data)};

      new Chart(issueChart, {
        type: "doughnut",
        data: {
          labels: Object.keys(data.issue_type_trends),
          datasets: [{ data: Object.values(data.issue_type_trends) }]
        }
      });

      new Chart(agentChart, {
        type: "bar",
        data: {
          labels: Object.keys(data.agent_performance),
          datasets: [{
            data: Object.values(data.agent_performance).map(a => a.resolved)
          }]
        }
      });

      new Chart(companyChart, {
        type: "bar",
        data: {
          labels: Object.keys(data.company_stats),
          datasets: [{
            data: Object.values(data.company_stats).map(c => c.total)
          }]
        }
      });

      const dates = Object.keys(data.daily_trends);
      new Chart(trendChart, {
        type: "line",
        data: {
          labels: dates,
          datasets: [
            { label: "Created", data: dates.map(d => data.daily_trends[d].created) },
            { label: "Resolved", data: dates.map(d => data.daily_trends[d].resolved) }
          ]
        }
      });

      new Chart(backlogChart, {
        type: "bar",
        data: {
          labels: ["2d","5d","10d"],
          datasets: [{
            data: [
              data.backlog_analysis.older_than_2_days,
              data.backlog_analysis.older_than_5_days,
              data.backlog_analysis.older_than_10_days
            ]
          }]
        }
      });

      new Chart(riskChart, {
        type: "pie",
        data: {
          labels: ["High Risk","Stuck","Waiting"],
          datasets: [{
            data: [
              data.risk_analysis.high_risk_tickets,
              data.risk_analysis.stuck_pending,
              data.risk_analysis.waiting_customer_long
            ]
          }]
        }
      });

      // 🔥 STACKED CHART
      const compIssues = data.company_issue_trends;

      const issueSet = new Set();
      Object.values(compIssues).forEach(obj => {
        Object.keys(obj).forEach(issue => issueSet.add(issue));
      });

      const issues = Array.from(issueSet);
      const companies = Object.keys(compIssues);

      const datasets = issues.map(issue => ({
        label: issue,
        data: companies.map(c => compIssues[c][issue] || 0)
      }));

      new Chart(stackedChart, {
        type: "bar",
        data: {
          labels: companies,
          datasets: datasets
        },
        options: {
          scales: {
            x: { stacked: true },
            y: { stacked: true }
          }
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
