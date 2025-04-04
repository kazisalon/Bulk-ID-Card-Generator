import { useState } from 'react';
import axios from 'axios';
import { HexColorPicker } from 'react-colorful';

function App() {
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [preview, setPreview] = useState(null);
  const [generating, setGenerating] = useState(false);
  const [downloadLink, setDownloadLink] = useState(null);
  const [error, setError] = useState(null);

  // Design options
  const [design, setDesign] = useState({
    font: 'Helvetica',
    textColor: '#000000',
    bgColor: '#FFFFFF',
  });

  const handleFileChange = (e) => {
    if (e.target.files[0]) {
      setFile(e.target.files[0]);
      setPreview(null);
      setDownloadLink(null);
      setError(null);
    }
  };

  const handleUpload = async () => {
    if (!file) {
      setError('Please select a file first');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setUploading(true);
    setError(null);

    try {
      const response = await axios.post('/api/upload', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });

      setPreview(response.data);
    } catch (err) {
      console.error('Upload error:', err);
      setError(err.response?.data?.error || 'Error uploading file');
    } finally {
      setUploading(false);
    }
  };

  const handleGeneratePDF = async () => {
    if (!preview) {
      setError('Please upload and process a file first');
      return;
    }

    setGenerating(true);
    setError(null);

    try {
      const response = await axios.post('/api/generate-pdf', {
        filePath: file.name,
        design: design,
      });

      setDownloadLink(response.data);
    } catch (err) {
      console.error('Generation error:', err);
      setError(err.response?.data?.error || 'Error generating PDF');
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <div className="max-w-4xl mx-auto bg-white rounded-lg shadow-md p-6">
        <h1 className="text-2xl font-bold mb-6 text-center">ID Card Generator</h1>

        {error && (
          <div className="mb-4 p-3 bg-red-100 text-red-700 rounded-md">
            {error}
          </div>
        )}

        <div className="mb-6">
          <h2 className="text-lg font-semibold mb-2">1. Upload Excel Sheet</h2>
          <p className="text-sm text-gray-600 mb-4">
            The file should include columns for Name, Photo Path, and other details.
          </p>
          
          <div className="flex flex-col md:flex-row gap-4">
            <input
              type="file"
              onChange={handleFileChange}
              className="border rounded-md p-2 w-full md:w-auto"
              accept=".xlsx,.xls"
            />
            
            <button
              onClick={handleUpload}
              disabled={!file || uploading}
              className={`px-4 py-2 rounded-md ${
                !file || uploading
                  ? 'bg-gray-300 cursor-not-allowed'
                  : 'bg-blue-500 hover:bg-blue-600 text-white'
              }`}
            >
              {uploading ? 'Uploading...' : 'Upload & Process'}
            </button>
          </div>
        </div>

        {preview && (
          <>
            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-2">2. Preview Data</h2>
              <div className="overflow-x-auto">
                <table className="min-w-full border-collapse border border-gray-300">
                  <thead>
                    <tr className="bg-gray-100">
                      {preview.columns.map((column) => (
                        <th key={column} className="border border-gray-300 px-4 py-2 text-left">
                          {column}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {preview.preview.map((row, index) => (
                      <tr key={index}>
                        {preview.columns.map((column) => (
                          <td key={column} className="border border-gray-300 px-4 py-2">
                            {row[column]?.toString() || ''}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-2">3. Customize ID Design</h2>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <div className="mb-4">
                    <label className="block text-sm font-medium mb-1">Font</label>
                    <select
                      value={design.font}
                      onChange={(e) => setDesign({ ...design, font: e.target.value })}
                      className="w-full p-2 border rounded-md"
                    >
                      <option value="Helvetica">Helvetica</option>
                      <option value="Times-Roman">Times Roman</option>
                      <option value="Courier">Courier</option>
                    </select>
                  </div>
                  
                  <div className="mb-4">
                    <label className="block text-sm font-medium mb-1">Text Color</label>
                    <div className="flex flex-col gap-2">
                      <HexColorPicker
                        color={design.textColor}
                        onChange={(color) => setDesign({ ...design, textColor: color })}
                      />
                      <input
                        type="text"
                        value={design.textColor}
                        onChange={(e) => setDesign({ ...design, textColor: e.target.value })}
                        className="p-2 border rounded-md"
                      />
                    </div>
                  </div>
                </div>
                
                <div>
                  <div className="mb-4">
                    <label className="block text-sm font-medium mb-1">Background Color</label>
                    <div className="flex flex-col gap-2">
                      <HexColorPicker
                        color={design.bgColor}
                        onChange={(color) => setDesign({ ...design, bgColor: color })}
                      />
                      <input
                        type="text"
                        value={design.bgColor}
                        onChange={(e) => setDesign({ ...design, bgColor: e.target.value })}
                        className="p-2 border rounded-md"
                      />
                    </div>
                  </div>
                </div>
              </div>
              
              <div className="mt-4">
                <div className="bg-gray-100 p-4 rounded-md">
                  <h3 className="font-medium mb-2">Preview of Design Settings:</h3>
                  <div 
                    className="border rounded-md p-4 flex items-center justify-between" 
                    style={{ backgroundColor: design.bgColor }}
                  >
                    <div style={{ color: design.textColor, fontFamily: design.font }}>
                      <div className="text-lg font-bold">Sample ID Card</div>
                      <div>Name: John Doe</div>
                      <div>ID: 12345</div>
                      <div>Department: Engineering</div>
                    </div>
                    <div className="w-20 h-24 bg-gray-300 flex items-center justify-center">
                      <span className="text-xs text-gray-600">Photo</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-2">4. Generate ID Cards</h2>
              <button
                onClick={handleGeneratePDF}
                disabled={generating}
                className={`px-4 py-2 rounded-md ${
                  generating
                    ? 'bg-gray-300 cursor-not-allowed'
                    : 'bg-green-500 hover:bg-green-600 text-white'
                }`}
              >
                {generating ? 'Generating...' : 'Generate ID Cards PDF'}
              </button>
            </div>
          </>
        )}

        {downloadLink && (
          <div className="mb-6 p-4 bg-green-100 rounded-md">
            <h2 className="text-lg font-semibold mb-2">5. Download ID Cards</h2>
            <p className="mb-2">Your ID cards have been generated successfully!</p>
            <a
              href={downloadLink.downloadUrl}
              download={downloadLink.filename}
              className="inline-block px-4 py-2 bg-blue-500 hover:bg-blue-600 text-white rounded-md"
            >
              Download PDF
            </a>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;