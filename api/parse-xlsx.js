const XLSX = require("xlsx");

module.exports = async (req, res) => {
	if (req.method !== "POST") {
		return res.status(405).json({
			success: false,
			error: "Method not allowed",
		});
	}

	try {
		const { file, filename } = req.body;

		if (!file) {
			return res.status(400).json({
				success: false,
				error: "No XLSX file provided",
			});
		}

		console.log("Received file:", filename);
		console.log("Input starts with:", file.substring(0, 20));

		/*
		 * Copilot is currently sending:
		 *
		 * Raw XLSX bytes
		 *      ↓ base64()
		 * VUVzRE...
		 *
		 * First decode:
		 * VUVzRE...
		 *      ↓
		 * UEsDB...
		 *
		 * Second decode:
		 * UEsDB...
		 *      ↓
		 * PK...
		 *
		 * Therefore we decode twice.
		 */

		// First Base64 decode
		let decoded = Buffer.from(file, "base64");

		console.log("After first decode:", decoded.toString("utf8", 0, 20));

		/*
		 * If the first decoded data starts with "UEsDB",
		 * it is still Base64 encoded.
		 */
		const decodedText = decoded.toString("utf8").trim();

		let buffer;

		if (decodedText.startsWith("UEsDB")) {
			console.log("Detected double Base64 encoding.");

			// Second Base64 decode → actual XLSX bytes
			buffer = Buffer.from(decodedText, "base64");
		} else {
			// Already actual XLSX bytes
			buffer = decoded;
		}

		console.log("Final buffer size:", buffer.length);

		// XLSX files are ZIP files and normally begin with PK
		console.log("File signature:", buffer.subarray(0, 4).toString("hex"));

		if (buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
			return res.status(400).json({
				success: false,
				error: "Decoded content is not a valid XLSX file",
				signature: buffer.subarray(0, 8).toString("hex"),
			});
		}

		// Read workbook
		const workbook = XLSX.read(buffer, {
			type: "buffer",
		});

		if (!workbook.SheetNames.length) {
			return res.status(400).json({
				success: false,
				error: "XLSX file contains no sheets",
			});
		}

		const sheetName = workbook.SheetNames[0];
		const worksheet = workbook.Sheets[sheetName];

		console.log("Sheet:", sheetName);
		console.log("Range:", worksheet["!ref"]);

		// Read worksheet as arrays
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

		// First row = headers
		const headers = rows[0];

		console.log("Headers:", headers);

		// Find AID column
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

		console.log("Extracted AIDs:", aids);

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
