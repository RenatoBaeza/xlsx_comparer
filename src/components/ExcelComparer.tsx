'use client';

import { useState, useMemo } from 'react';
import * as ExcelJS from 'exceljs';

interface ExcelFile {
  name: string;
  data: ExcelJS.Workbook;
}

// Type for Excel cell values
type ExcelCellValue = string | number | boolean | null | undefined;

// Type for Excel row data
type ExcelRow = ExcelCellValue[];

// Type for Excel sheet data
type ExcelSheetData = ExcelRow[];

interface CellDifference {
  row: number;
  col: number;
  value1: string;
  value2: string;
  type: 'added' | 'removed' | 'modified';
}

interface RowDifference {
  rowNumber: number;
  cells: {
    col: number;
    value1: string;
    value2: string;
    isDifferent: boolean;
    differenceType?: 'added' | 'removed' | 'modified';
  }[];
  hasDifferences: boolean;
}

interface SheetComparison {
  sheetName: string;
  differences: CellDifference[];
  rowDifferences: RowDifference[];
  totalDifferences: number;
  hasDifferences: boolean;
  headers: string[];
}

interface ExcelComparerProps {
  file1: ExcelFile;
  file2: ExcelFile;
  headerRow: number;
}

// Helper function to convert ExcelJS worksheet to array data
const worksheetToArray = (worksheet: ExcelJS.Worksheet): ExcelSheetData => {
  const data: ExcelSheetData = [];
  const dimensions = worksheet.dimensions;
  
  if (!dimensions) return data;
  
  for (let row = dimensions.top; row <= dimensions.bottom; row++) {
    const rowData: ExcelRow = [];
    for (let col = dimensions.left; col <= dimensions.right; col++) {
      const cell = worksheet.getCell(row, col);
      let value: ExcelCellValue;
      
      if (cell.value === null || cell.value === undefined) {
        value = '';
      } else if (typeof cell.value === 'object' && 'result' in cell.value) {
        // Handle formula results
        value = cell.value.result?.toString() || '';
      } else {
        value = cell.value.toString();
      }
      
      rowData.push(value);
    }
    data.push(rowData);
  }
  
  return data;
};

export default function ExcelComparer({ file1, file2, headerRow }: ExcelComparerProps) {
  const [activeTab, setActiveTab] = useState<string>('');
  const [expandedRows, setExpandedRows] = useState<Set<number>>(new Set());

  const comparisonResults = useMemo(() => {
    const results: SheetComparison[] = [];
    
    // Get all unique sheet names from both files
    const allSheetNames = new Set([
      ...file1.data.worksheets.map(ws => ws.name),
      ...file2.data.worksheets.map(ws => ws.name)
    ]);

    allSheetNames.forEach(sheetName => {
      const sheet1 = file1.data.getWorksheet(sheetName);
      const sheet2 = file2.data.getWorksheet(sheetName);
      
      const differences: CellDifference[] = [];
      const rowDifferences: RowDifference[] = [];
      
      // Get headers from the specified header row
      let headers: string[] = [];
      if (sheet1) {
        const jsonData1 = worksheetToArray(sheet1);
        if (jsonData1.length >= headerRow) {
          headers = jsonData1[headerRow - 1].map((cell: ExcelCellValue) => cell !== undefined && cell !== null ? String(cell) : '');
        }
      } else if (sheet2) {
        const jsonData2 = worksheetToArray(sheet2);
        if (jsonData2.length >= headerRow) {
          headers = jsonData2[headerRow - 1].map((cell: ExcelCellValue) => cell !== undefined && cell !== null ? String(cell) : '');
        }
      }
      
      if (!sheet1 && sheet2) {
        // Sheet exists only in file2
        const jsonData = worksheetToArray(sheet2);
        // Skip rows up to and including the header row when comparing
        const dataRows = jsonData.slice(headerRow);
        dataRows.forEach((row: ExcelRow, rowIndex: number) => {
          const rowDiff: RowDifference = {
            rowNumber: rowIndex + headerRow + 1, // +1 because rowIndex starts at 0, +headerRow to account for skipped rows
            cells: [],
            hasDifferences: false
          };
          
          row.forEach((cell: ExcelCellValue, colIndex: number) => {
            if (cell !== undefined && cell !== null && cell !== '') {
              rowDiff.cells.push({
                col: colIndex + 1,
                value1: '',
                value2: String(cell),
                isDifferent: true,
                differenceType: 'added'
              });
              rowDiff.hasDifferences = true;
              
              differences.push({
                row: rowIndex + headerRow + 1,
                col: colIndex + 1,
                value1: '',
                value2: String(cell),
                type: 'added'
              });
            } else {
              rowDiff.cells.push({
                col: colIndex + 1,
                value1: '',
                value2: '',
                isDifferent: false
              });
            }
          });
          
          if (rowDiff.hasDifferences) {
            rowDifferences.push(rowDiff);
          }
        });
      } else if (sheet1 && !sheet2) {
        // Sheet exists only in file1
        const jsonData = worksheetToArray(sheet1);
        // Skip rows up to and including the header row when comparing
        const dataRows = jsonData.slice(headerRow);
        dataRows.forEach((row: ExcelRow, rowIndex: number) => {
          const rowDiff: RowDifference = {
            rowNumber: rowIndex + headerRow + 1, // +1 because rowIndex starts at 0, +headerRow to account for skipped rows
            cells: [],
            hasDifferences: false
          };
          
          row.forEach((cell: ExcelCellValue, colIndex: number) => {
            if (cell !== undefined && cell !== null && cell !== '') {
              rowDiff.cells.push({
                col: colIndex + 1,
                value1: String(cell),
                value2: '',
                isDifferent: true,
                differenceType: 'removed'
              });
              rowDiff.hasDifferences = true;
              
              differences.push({
                row: rowIndex + headerRow + 1,
                col: colIndex + 1,
                value1: String(cell),
                value2: '',
                type: 'removed'
              });
            } else {
              rowDiff.cells.push({
                col: colIndex + 1,
                value1: '',
                value2: '',
                isDifferent: false
              });
            }
          });
          
          if (rowDiff.hasDifferences) {
            rowDifferences.push(rowDiff);
          }
        });
      } else if (sheet1 && sheet2) {
        // Both sheets exist, compare them
        const jsonData1 = worksheetToArray(sheet1);
        const jsonData2 = worksheetToArray(sheet2);
        
        // Skip rows up to and including the header row when comparing
        const dataRows1 = jsonData1.slice(headerRow);
        const dataRows2 = jsonData2.slice(headerRow);
        
        const maxRows = Math.max(dataRows1.length, dataRows2.length);
        const maxCols = Math.max(
          ...dataRows1.map((row: ExcelRow) => row.length),
          ...dataRows2.map((row: ExcelRow) => row.length)
        );
        
        // Track rows that have differences
        const rowsWithDifferences = new Set<number>();
        
        for (let row = 0; row < maxRows; row++) {
          for (let col = 0; col < maxCols; col++) {
            const cell1 = dataRows1[row]?.[col];
            const cell2 = dataRows2[row]?.[col];
            
            const value1 = cell1 !== undefined && cell1 !== null ? String(cell1) : '';
            const value2 = cell2 !== undefined && cell2 !== null ? String(cell2) : '';
            
            if (value1 !== value2) {
              rowsWithDifferences.add(row + headerRow + 1); // +1 because rowIndex starts at 0, +headerRow to account for skipped rows
              differences.push({
                row: row + headerRow + 1,
                col: col + 1,
                value1,
                value2,
                type: value1 === '' ? 'added' : value2 === '' ? 'removed' : 'modified'
              });
            }
          }
        }
        
        // Create row differences for rows that have changes
        rowsWithDifferences.forEach(rowNum => {
          const rowIndex = rowNum - headerRow - 1; // Convert back to data row index
          const rowDiff: RowDifference = {
            rowNumber: rowNum,
            cells: [],
            hasDifferences: true
          };
          
          for (let col = 0; col < maxCols; col++) {
            const cell1 = dataRows1[rowIndex]?.[col];
            const cell2 = dataRows2[rowIndex]?.[col];
            
            const value1 = cell1 !== undefined && cell1 !== null ? String(cell1) : '';
            const value2 = cell2 !== undefined && cell2 !== null ? String(cell2) : '';
            const isDifferent = value1 !== value2;
            
            let differenceType: 'added' | 'removed' | 'modified' | undefined;
            if (isDifferent) {
              differenceType = value1 === '' ? 'added' : value2 === '' ? 'removed' : 'modified';
            }
            
            rowDiff.cells.push({
              col: col + 1,
              value1,
              value2,
              isDifferent,
              differenceType
            });
          }
          
          rowDifferences.push(rowDiff);
        });
        
        // Sort rows by row number
        rowDifferences.sort((a, b) => a.rowNumber - b.rowNumber);
      }
      
      results.push({
        sheetName,
        differences,
        rowDifferences,
        totalDifferences: differences.length,
        hasDifferences: differences.length > 0,
        headers
      });
    });
    
    // Set the first sheet with differences as active tab, or the first sheet if no differences
    const firstSheetWithDifferences = results.find(r => r.hasDifferences);
    if (firstSheetWithDifferences && !activeTab) {
      setActiveTab(firstSheetWithDifferences.sheetName);
    } else if (results.length > 0 && !activeTab) {
      setActiveTab(results[0].sheetName);
    }
    
    return results;
  }, [file1, file2, headerRow, activeTab]);

  const totalDifferences = comparisonResults.reduce((sum, sheet) => sum + sheet.totalDifferences, 0);

  const getColumnLetter = (colIndex: number): string => {
    let result = '';
    while (colIndex > 0) {
      colIndex--;
      result = String.fromCharCode(65 + (colIndex % 26)) + result;
      colIndex = Math.floor(colIndex / 26);
    }
    return result;
  };

  const getDifferenceTypeColor = (type: string) => {
    switch (type) {
      case 'added':
        return 'bg-green-100 text-green-800 border-green-200';
      case 'removed':
        return 'bg-red-100 text-red-800 border-red-200';
      case 'modified':
        return 'bg-yellow-100 text-yellow-800 border-yellow-200';
      default:
        return 'bg-gray-100 text-gray-800 border-gray-200';
    }
  };

  const getDifferenceTypeLabel = (type: string) => {
    switch (type) {
      case 'added':
        return 'Added';
      case 'removed':
        return 'Removed';
      case 'modified':
        return 'Modified';
      default:
        return 'Unknown';
    }
  };

  const getCellBackgroundColor = (isDifferent: boolean, differenceType?: string) => {
    if (!isDifferent) return 'bg-white';
    switch (differenceType) {
      case 'added':
        return 'bg-green-50';
      case 'removed':
        return 'bg-red-50';
      case 'modified':
        return 'bg-yellow-50';
      default:
        return 'bg-white';
    }
  };

  const toggleRowExpansion = (rowNumber: number) => {
    const newExpandedRows = new Set(expandedRows);
    if (newExpandedRows.has(rowNumber)) {
      newExpandedRows.delete(rowNumber);
    } else {
      newExpandedRows.add(rowNumber);
    }
    setExpandedRows(newExpandedRows);
  };

  const getRowSummary = (rowDiff: RowDifference) => {
    const added = rowDiff.cells.filter(cell => cell.differenceType === 'added').length;
    const removed = rowDiff.cells.filter(cell => cell.differenceType === 'removed').length;
    const modified = rowDiff.cells.filter(cell => cell.differenceType === 'modified').length;
    
    const parts = [];
    if (added > 0) parts.push(`${added} added`);
    if (removed > 0) parts.push(`${removed} removed`);
    if (modified > 0) parts.push(`${modified} modified`);
    
    return parts.join(', ');
  };

  const getHeaderName = (colIndex: number, headers: string[]) => {
    const header = headers[colIndex - 1];
    return header || getColumnLetter(colIndex);
  };

  return (
    <div className="bg-white rounded-lg shadow-md">
      {/* Summary */}
      <div className="p-6 border-b border-gray-200">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-semibold text-gray-900">Comparison Summary</h3>
          <div className="text-sm text-gray-600">
            Total differences found: <span className="font-semibold">{totalDifferences}</span>
          </div>
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="bg-blue-50 p-4 rounded-lg">
            <div className="text-sm font-medium text-blue-800">File 1</div>
            <div className="text-lg font-semibold text-blue-900">{file1.name}</div>
            <div className="text-sm text-blue-600">{file1.data.worksheets.length} sheets</div>
          </div>
          <div className="bg-green-50 p-4 rounded-lg">
            <div className="text-sm font-medium text-green-800">File 2</div>
            <div className="text-lg font-semibold text-green-900">{file2.name}</div>
            <div className="text-sm text-green-600">{file2.data.worksheets.length} sheets</div>
          </div>
          <div className="bg-purple-50 p-4 rounded-lg">
            <div className="text-sm font-medium text-purple-800">Sheets Compared</div>
            <div className="text-lg font-semibold text-purple-900">{comparisonResults.length}</div>
            <div className="text-sm text-purple-600">
              {comparisonResults.filter(r => r.hasDifferences).length} with differences
            </div>
          </div>
        </div>
        
        <div className="mt-4 p-3 bg-gray-50 rounded-lg">
          <div className="text-sm text-gray-600">
            <span className="font-medium">Header Row:</span> Row {headerRow} is being used for column headers
          </div>
        </div>
      </div>

      {/* Sheet Tabs */}
      <div className="border-b border-gray-200">
        <div className="flex overflow-x-auto">
          {comparisonResults.map((sheet) => (
            <button
              key={sheet.sheetName}
              onClick={() => setActiveTab(sheet.sheetName)}
              className={`px-4 py-3 text-sm font-medium border-b-2 whitespace-nowrap ${
                activeTab === sheet.sheetName
                  ? 'border-blue-500 text-blue-600'
                  : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
              }`}
            >
              {sheet.sheetName}
              {sheet.hasDifferences && (
                <span className="ml-2 inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-800">
                  {sheet.totalDifferences}
                </span>
              )}
            </button>
          ))}
        </div>
      </div>

      {/* Sheet Content */}
      {activeTab && (
        <div className="p-6">
          <div className="mb-4">
            <h4 className="text-lg font-semibold text-gray-900 mb-2">
              Sheet: {activeTab}
            </h4>
            {comparisonResults.find(r => r.sheetName === activeTab)?.hasDifferences ? (
              <div className="text-sm text-gray-600">
                {comparisonResults.find(r => r.sheetName === activeTab)?.totalDifferences} differences found across{' '}
                {comparisonResults.find(r => r.sheetName === activeTab)?.rowDifferences.length} rows
              </div>
            ) : (
              <div className="text-sm text-green-600 font-medium">✓ No differences found</div>
            )}
          </div>

          {comparisonResults.find(r => r.sheetName === activeTab)?.hasDifferences && (
            <div className="space-y-3">
              {comparisonResults
                .find(r => r.sheetName === activeTab)
                ?.rowDifferences.map((rowDiff, rowIndex) => {
                  const currentSheet = comparisonResults.find(r => r.sheetName === activeTab);
                  const headers = currentSheet?.headers || [];
                  
                  return (
                    <div key={rowIndex} className="border border-gray-200 rounded-lg overflow-hidden">
                      {/* Row Summary Header */}
                      <div 
                        className="bg-gray-50 px-4 py-3 border-b border-gray-200 cursor-pointer hover:bg-gray-100 transition-colors"
                        onClick={() => toggleRowExpansion(rowDiff.rowNumber)}
                      >
                        <div className="flex items-center justify-between">
                          <div className="flex items-center space-x-3">
                            <svg
                              className={`w-4 h-4 text-gray-500 transition-transform ${
                                expandedRows.has(rowDiff.rowNumber) ? 'rotate-90' : ''
                              }`}
                              fill="none"
                              stroke="currentColor"
                              viewBox="0 0 24 24"
                            >
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                            </svg>
                            <h5 className="text-sm font-medium text-gray-900">
                              Row {rowDiff.rowNumber}
                            </h5>
                            <span className="text-sm text-gray-600">
                              ({getRowSummary(rowDiff)})
                            </span>
                          </div>
                          <div className="flex items-center space-x-2">
                            {rowDiff.cells.filter(cell => cell.differenceType === 'added').length > 0 && (
                              <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-green-100 text-green-800">
                                +{rowDiff.cells.filter(cell => cell.differenceType === 'added').length}
                              </span>
                            )}
                            {rowDiff.cells.filter(cell => cell.differenceType === 'removed').length > 0 && (
                              <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-red-100 text-red-800">
                                -{rowDiff.cells.filter(cell => cell.differenceType === 'removed').length}
                              </span>
                            )}
                            {rowDiff.cells.filter(cell => cell.differenceType === 'modified').length > 0 && (
                              <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-yellow-100 text-yellow-800">
                                ~{rowDiff.cells.filter(cell => cell.differenceType === 'modified').length}
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                      
                      {/* Expanded Row Details */}
                      {expandedRows.has(rowDiff.rowNumber) && (
                        <div className="overflow-x-auto">
                          <table className="min-w-full">
                            <thead className="bg-gray-100">
                              <tr>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider border-r border-gray-200">
                                  Cell
                                </th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider border-r border-gray-200">
                                  Header
                                </th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider border-r border-gray-200">
                                  {file1.name}
                                </th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                                  {file2.name}
                                </th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-200">
                              {rowDiff.cells.map((cell, cellIndex) => (
                                <tr key={cellIndex} className={getCellBackgroundColor(cell.isDifferent, cell.differenceType)}>
                                  <td className="px-3 py-2 text-sm font-mono text-gray-900 border-r border-gray-200">
                                    {getColumnLetter(cell.col)}{rowDiff.rowNumber}
                                    {cell.isDifferent && (
                                      <span className={`ml-2 inline-flex items-center px-1.5 py-0.5 rounded text-xs font-medium border ${getDifferenceTypeColor(cell.differenceType!)}`}>
                                        {getDifferenceTypeLabel(cell.differenceType!)}
                                      </span>
                                    )}
                                  </td>
                                  <td className="px-3 py-2 text-sm text-gray-600 border-r border-gray-200">
                                    {getHeaderName(cell.col, headers)}
                                  </td>
                                  <td className="px-3 py-2 text-sm text-gray-900 border-r border-gray-200">
                                    {cell.value1 || <span className="text-gray-400 italic">(empty)</span>}
                                  </td>
                                  <td className="px-3 py-2 text-sm text-gray-900">
                                    {cell.value2 || <span className="text-gray-400 italic">(empty)</span>}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })}
            </div>
          )}
        </div>
      )}
    </div>
  );
} 