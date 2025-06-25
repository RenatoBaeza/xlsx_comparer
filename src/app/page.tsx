'use client';

import { useState } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import ExcelComparer from '../components/ExcelComparer';

interface ExcelFile {
  name: string;
  data: XLSX.WorkBook;
}

export default function Home() {
  const [file1, setFile1] = useState<ExcelFile | null>(null);
  const [file2, setFile2] = useState<ExcelFile | null>(null);
  const [isComparing, setIsComparing] = useState(false);
  const [headerRow, setHeaderRow] = useState<number>(1);

  const onDrop1 = (acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      const file = acceptedFiles[0];
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        setFile1({ name: file.name, data: workbook });
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const onDrop2 = (acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      const file = acceptedFiles[0];
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        setFile2({ name: file.name, data: workbook });
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const { getRootProps: getRootProps1, getInputProps: getInputProps1, isDragActive: isDragActive1 } = useDropzone({
    onDrop: onDrop1,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false
  });

  const { getRootProps: getRootProps2, getInputProps: getInputProps2, isDragActive: isDragActive2 } = useDropzone({
    onDrop: onDrop2,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false
  });

  const handleCompare = () => {
    if (file1 && file2) {
      setIsComparing(true);
    }
  };

  const resetFiles = () => {
    setFile1(null);
    setFile2(null);
    setIsComparing(false);
  };

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-900 mb-4">
            Excel File Comparer
          </h1>
          <p className="text-lg text-gray-600">
            Upload two Excel files to compare their contents across all sheets
          </p>
        </div>

        {!isComparing ? (
          <div className="space-y-8">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              {/* File 1 Upload */}
              <div className="bg-white rounded-lg shadow-md p-6">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">
                  First Excel File
                </h2>
                <div
                  {...getRootProps1()}
                  className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${
                    isDragActive1
                      ? 'border-blue-400 bg-blue-50'
                      : file1
                      ? 'border-green-400 bg-green-50'
                      : 'border-gray-300 hover:border-gray-400'
                  }`}
                >
                  <input {...getInputProps1()} />
                  {file1 ? (
                    <div>
                      <div className="text-green-600 font-medium mb-2">
                        ✓ File uploaded successfully
                      </div>
                      <div className="text-sm text-gray-600">{file1.name}</div>
                      <div className="text-xs text-gray-500 mt-1">
                        {file1.data.SheetNames.length} sheet(s)
                      </div>
                    </div>
                  ) : (
                    <div>
                      <div className="text-gray-400 mb-2">
                        <svg className="mx-auto h-12 w-12" stroke="currentColor" fill="none" viewBox="0 0 48 48">
                          <path d="M28 8H12a4 4 0 00-4 4v20m32-12v8m0 0v8a4 4 0 01-4 4H12a4 4 0 01-4-4v-4m32-4l-3.172-3.172a4 4 0 00-5.656 0L28 28M8 32l9.172-9.172a4 4 0 015.656 0L28 28m0 0l4 4m4-24h8m-4-4v8m-12 4h.02" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                      </div>
                      <p className="text-gray-600">
                        {isDragActive1 ? 'Drop the file here' : 'Drag & drop an Excel file here, or click to select'}
                      </p>
                      <p className="text-xs text-gray-500 mt-2">Supports .xlsx and .xls files</p>
                    </div>
                  )}
                </div>
              </div>

              {/* File 2 Upload */}
              <div className="bg-white rounded-lg shadow-md p-6">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">
                  Second Excel File
                </h2>
                <div
                  {...getRootProps2()}
                  className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${
                    isDragActive2
                      ? 'border-blue-400 bg-blue-50'
                      : file2
                      ? 'border-green-400 bg-green-50'
                      : 'border-gray-300 hover:border-gray-400'
                  }`}
                >
                  <input {...getInputProps2()} />
                  {file2 ? (
                    <div>
                      <div className="text-green-600 font-medium mb-2">
                        ✓ File uploaded successfully
                      </div>
                      <div className="text-sm text-gray-600">{file2.name}</div>
                      <div className="text-xs text-gray-500 mt-1">
                        {file2.data.SheetNames.length} sheet(s)
                      </div>
                    </div>
                  ) : (
                    <div>
                      <div className="text-gray-400 mb-2">
                        <svg className="mx-auto h-12 w-12" stroke="currentColor" fill="none" viewBox="0 0 48 48">
                          <path d="M28 8H12a4 4 0 00-4 4v20m32-12v8m0 0v8a4 4 0 01-4 4H12a4 4 0 01-4-4v-4m32-4l-3.172-3.172a4 4 0 00-5.656 0L28 28M8 32l9.172-9.172a4 4 0 015.656 0L28 28m0 0l4 4m4-24h8m-4-4v8m-12 4h.02" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                      </div>
                      <p className="text-gray-600">
                        {isDragActive2 ? 'Drop the file here' : 'Drag & drop an Excel file here, or click to select'}
                      </p>
                      <p className="text-xs text-gray-500 mt-2">Supports .xlsx and .xls files</p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Header Row Configuration */}
            <div className="bg-white rounded-lg shadow-md p-6">
              <h2 className="text-xl font-semibold text-gray-900 mb-4">
                Comparison Settings
              </h2>
              <div className="max-w-md">
                <label htmlFor="headerRow" className="block text-sm font-medium text-gray-700 mb-2">
                  Header Row Number
                </label>
                <div className="flex items-center space-x-4">
                  <input
                    type="number"
                    id="headerRow"
                    min="1"
                    value={headerRow}
                    onChange={(e) => setHeaderRow(Math.max(1, parseInt(e.target.value) || 1))}
                    className="block w-20 px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500"
                  />
                  <span className="text-sm text-gray-600">
                    Which row contains the column headers? (Default: 1)
                  </span>
                </div>
                <p className="text-xs text-gray-500 mt-1">
                  If your headers are not in the first row, specify the correct row number here.
                </p>
              </div>
            </div>

            {/* Compare Button */}
            <div className="text-center">
              <button
                onClick={handleCompare}
                disabled={!file1 || !file2}
                className={`px-8 py-3 rounded-lg font-medium transition-colors ${
                  file1 && file2
                    ? 'bg-blue-600 text-white hover:bg-blue-700'
                    : 'bg-gray-300 text-gray-500 cursor-not-allowed'
                }`}
              >
                Compare Files
              </button>
            </div>
          </div>
        ) : (
          <div>
            <div className="mb-6 flex justify-between items-center">
              <h2 className="text-2xl font-bold text-gray-900">Comparison Results</h2>
              <button
                onClick={resetFiles}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition-colors"
              >
                Upload New Files
              </button>
            </div>
            <ExcelComparer file1={file1!} file2={file2!} headerRow={headerRow} />
          </div>
        )}
      </div>
    </div>
  );
}
