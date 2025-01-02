const express = require("express");
const https = require("https");
const axios = require("axios");
const { DateTime } = require("luxon");
const ExcelJS = require("exceljs");
const nodemailer = require("nodemailer");
const { HttpsProxyAgent } = require("https-proxy-agent"); // Import thư viện proxy

const app = express();
const port = 3000;

// Thông tin proxy
const proxyAgent = new HttpsProxyAgent(
	"http://hamia_oncall:OHLTJGkl75FMobrv@27.71.230.22:31281"
);

const agent = new https.Agent({
	rejectUnauthorized: false,
});

const statusCodeDescriptions = {
	0: "Normal Call Clearing",
	404: "Phone number Not Found",
	480: "Temporary Unavailable",
	484: "Address Incomplete",
	486: "Busy Here",
	487: "Request Terminated",
	503: "Service Unavailable",
};

async function getAccessToken() {
	try {
		const response = await axios.post(
			"https://rsv02.oncall.vn:8887/api/tokens",
			{
				username: "HNCX00609",
				password: "EG9693dyhamia",
				domain: "hncx00609.oncall",
			},
			{ httpsAgent: agent }
			// {
			// 	httpsAgent: proxyAgent, // Sử dụng proxy
			// }
		);
		return response.data.access_token;
	} catch (error) {
		throw new Error(error.response ? error.response.data : error.message);
	}
}

async function callInfor() {
	try {
		const accessToken = await getAccessToken();
		const yesterday = DateTime.now()
			.setZone("Asia/Bangkok")
			.minus({ days: 1 });
		const startOfDay = yesterday.startOf("day").toISO();
		const endOfDay = yesterday.endOf("day").toISO();
		const filter = `started_at gt '${startOfDay}' and started_at lt '${endOfDay}'`;

		const response = await axios.get(
			"https://rsv02.oncall.vn:8887/api/cdrs",
			{
				headers: {
					Authorization: `Bearer ${accessToken}`,
				},
				params: { filter },
				httpsAgent: agent,
				// httpsAgent: proxyAgent, // Sử dụng proxy
			}
		);

		return response.data.items
			.filter((call) => ["120", "123", "125", "180"].includes(call.caller))
			.map((call) => ({
				id: call.id,
				caller: call.caller,
				callee: call.callee,
				started_at: call.started_at,
				ended_at: call.ended_at,
				status_code: call.status_code,
				duration: call.duration,
			}));
	} catch (error) {
		throw new Error(error.response ? error.response.data : error.message);
	}
}

async function createDailyReport() {
	try {
		const calls = await callInfor();
		const workbook = new ExcelJS.Workbook();
		const worksheet = workbook.addWorksheet("Daily Report");

		worksheet.columns = [
			{ header: "ID", key: "id", width: 20 },
			{ header: "Caller", key: "caller", width: 20 },
			{ header: "Callee", key: "callee", width: 20 },
			{ header: "Start time", key: "started_at", width: 25 },
			{ header: "End time", key: "ended_at", width: 25 },
			{ header: "Status", key: "status_code", width: 30 },
			{ header: "Duration (s)", key: "duration", width: 15 },
		];

		calls.forEach((call) => {
			worksheet.addRow({
				...call,
				status_code:
					statusCodeDescriptions[call.status_code] ||
					call.status_code,
			});
		});

		const buffer = await workbook.xlsx.writeBuffer();
		return buffer;
	} catch (error) {
		throw new Error("Failed to create report: " + error.message);
	}
}

async function sendMail() {
	try {
		const reportBuffer = await createDailyReport();
		const transporter = nodemailer.createTransport({
			host: "smtp.gmail.com",
			port: 465,
			secure: true,
			auth: {
				user: "tda.ducanh@gmail.com",
				pass: "wlur rdtb rger zser",
			},
		});

		await transporter.sendMail({
			from: "tda.ducanh@gmail.com",
			to: "huyennt@ispeak.vn",
			subject: `Oncall báo cáo ${DateTime.now()
				.setZone("Asia/Bangkok")
				.toFormat("yyyy-MM-dd")}`,
			attachments: [
				{
					filename: `Oncall_daily_Report_${DateTime.now()
						.setZone("Asia/Bangkok")
						.toFormat("yyyy-MM-dd")}.xlsx`,
					content: reportBuffer,
				},
			],
		});

		console.log(
			`Email ${DateTime.now()
				.setZone("Asia/Bangkok")
				.toFormat("yyyy-MM-dd")} sent successfully.`
		);
	} catch (error) {
		console.error("Failed to send email:", error.message);
	}
}

app.get("/send-daily-report", async (req, res) => {
	try {
		await sendMail();
		res.status(200).send("Daily report sent successfully.");
	} catch (error) {
		res.status(500).json({
			message: "Failed to send daily report",
			error: error.message,
		});
	}
});

app.listen(port, () => {
	console.log(`Server is running on http://localhost:${port}`);
});
