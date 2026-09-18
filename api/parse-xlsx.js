const XLSX = require("xlsx");

module.exports = async (req, res) => {
	// Only allow POST
	if (req.method !== "POST") {
		return res.status(405).json({
			error: "Method not allowed",
		});
	}

	try {
		// Expect JSON body:
		// {
		//   "file": "<base64 string>",
		//   "filename": "testbook.xlsx"
		// }

		const { file, filename } = req.body;

		// Check whether file was provided
		if (!file) {
			return res.status(400).json({
				error: "No XLSX file provided",
			});
		}

		// Convert Base64 string into a Buffer
		const buffer = Buffer.from(file, "base64");

		// Read the Excel workbook
		const workbook = XLSX.read(buffer, {
			type: "buffer",
		});

		// Make sure workbook has at least one sheet
		if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
			return res.status(400).json({
				error: "XLSX file contains no sheets",
			});
		}

		// Get first sheet
		const sheetName = workbook.SheetNames[0];
		const worksheet = workbook.Sheets[sheetName];

		// Convert sheet to JSON
		const rows = XLSX.utils.sheet_to_json(worksheet);

		// Extract AID column
		const aids = rows
			.map((row) => row.AID)
			.filter(
				(value) =>
					value !== undefined && value !== null && value !== "",
			);

		// Return result
		return res.status(200).json({
			success: true,
			filename: filename || null,
			sheet: sheetName,
			rowCount: rows.length,
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
