const multer = require("multer");
const XLSX = require("xlsx");

const upload = multer({
	storage: multer.memoryStorage(),
});

const uploadMiddleware = upload.single("file");

module.exports = async (req, res) => {
	if (req.method !== "POST") {
		return res.status(405).json({
			error: "Method not allowed",
		});
	}

	uploadMiddleware(req, res, (err) => {
		if (err) {
			return res.status(400).json({
				error: err.message,
			});
		}

		try {
			if (!req.file) {
				return res.status(400).json({
					error: "No XLSX file provided",
				});
			}

			const workbook = XLSX.read(req.file.buffer, {
				type: "buffer",
			});

			const sheetName = workbook.SheetNames[0];
			const worksheet = workbook.Sheets[sheetName];

			const rows = XLSX.utils.sheet_to_json(worksheet);

			// Always extract AID
			const aids = rows
				.map((row) => row.AID)
				.filter((value) => value !== undefined && value !== null);

			return res.status(200).json(aids);
		} catch (error) {
			console.error(error);

			return res.status(500).json({
				error: "Failed to parse XLSX file",
			});
		}
	});
};
