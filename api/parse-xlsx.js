const XLSX = require("xlsx");

module.exports = async (req, res) => {
	if (req.method !== "POST") {
		return res.status(405).json({
			error: "Method not allowed",
		});
	}

	try {
		const { file, filename } = req.body;

		if (!file) {
			return res.status(400).json({
				error: "No XLSX file provided",
			});
		}

		// Base64 → Buffer
		const buffer = Buffer.from(file, "base64");

		console.log("Received file:", filename);
		console.log("Buffer size:", buffer.length);

		// Read workbook
		const workbook = XLSX.read(buffer, {
			type: "buffer",
		});

		if (!workbook.SheetNames.length) {
			return res.status(400).json({
				error: "XLSX file contains no sheets",
			});
		}

		const sheetName = workbook.SheetNames[0];
		const worksheet = workbook.Sheets[sheetName];

		console.log("Sheet:", sheetName);
		console.log("Range:", worksheet["!ref"]);

		// Read rows as arrays instead of objects
		const rows = XLSX.utils.sheet_to_json(worksheet, {
			header: 1,
			defval: null,
			raw: true,
		});

		console.log("Rows:", rows);

		if (rows.length === 0) {
			return res.status(200).json({
				success: true,
				filename: filename || null,
				sheet: sheetName,
				rowCount: 0,
				aidCount: 0,
				aids: [],
				message: "Worksheet contains no readable rows",
			});
		}

		// First row is the header
		const headers = rows[0];

		const aidIndex = headers.findIndex(
			(header) => String(header).trim().toUpperCase() === "AID",
		);

		if (aidIndex === -1) {
			return res.status(400).json({
				success: false,
				error: "AID column not found",
				headers: headers,
			});
		}

		// Extract AID values
		const aids = rows
			.slice(1)
			.map((row) => row[aidIndex])
			.filter(
				(value) =>
					value !== undefined &&
					value !== null &&
					String(value).trim() !== "",
			);

		return res.status(200).json({
			success: true,
			filename: filename || null,
			sheet: sheetName,
			rowCount: rows.length - 1,
			aidCount: aids.length,
			aids: aids,
		});
	} catch (error) {
		console.error("XLSX parsing error:", error);

		return res.status(500).json({
			success: false,
			error: "Failed to parse XLSX file",
			details: error.message,
		});
	}
};
