import path from 'path';
import fs from 'fs';
import { parse } from 'url';

const uploadDir = path.join(process.cwd(), 'public/uploads');

export default async (req, res) => {
  try {
    const { query } = parse(req.url, true);
    const { filename } = query;

    if (!filename) {
      res.status(400).json({ message: 'No filename provided.' });
      return;
    }

    // Construct the full file path based on the filename and the upload directory
    const filePath = path.join(uploadDir, filename);

    // Check if the file exists
    if (!fs.existsSync(filePath)) {
      res.status(404).json({ message: 'File not found.' });
      return;
    }

    // Read the file data
    const fileData = fs.readFileSync(filePath);

    const workbook = XLSX.read(fileData, { type: 'buffer' });

    // Assuming you want to access the first sheet in the Excel file
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert the worksheet to an array of objects (representing rows)
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    console.log(jsonData)

    // Send the parsed data in the response
    res.status(200).json({ message: 'File retrieved successfully.', data: jsonData });
  } catch (error) {
    console.error('Error retrieving file:', error);
    res.status(500).json({ message: 'An error occurred while retrieving the file.' });
  }
};
